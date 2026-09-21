"""The browser session and the Teams and SharePoint calls made through it.

Requests are issued from inside pages the user is already signed in to, so the
session carries no credentials of its own. Teams API calls need a bearer token,
which is captured from the requests the Teams tab makes while it loads.
SharePoint transcript content needs only cookies.
"""
from __future__ import annotations

import asyncio
import json
import urllib.parse
from datetime import datetime, timedelta, timezone

from .browser import launch_context, log
from .config import TEAMS_URL
from .errors import TranscriptError
from .model import (
    APPROVED_TEAMS_HOSTS,
    MCPS_RE,
    MT_RE,
    UrlValidationError,
    iso_z,
    is_approved_auth_url,
    jwt_claims,
    quote_path_component,
    sanitize_terminal,
    thread_id_from,
    validate_sharepoint_host,
    validate_sharepoint_url,
    validate_teams_url,
    web_recap_url,
)

try:
    from playwright.async_api import async_playwright
except ImportError as exc:  # pragma: no cover
    raise TranscriptError("Playwright is missing. Install this tool with 'pip install msteams-transcripts'.") from exc

FETCH_JS = """async ([url, headers]) => {
  const r = await fetch(url, {headers, credentials: 'include'});
  const text = await r.text();
  return {status: r.status, ct: r.headers.get('content-type') || '', url: r.url, text};
}"""


def _safe_url_label(url: str) -> str:
    """Return an origin/path label without query strings or fragments."""
    try:
        parsed = urllib.parse.urlsplit(url)
        host = parsed.hostname or "unknown-host"
        path = parsed.path or "/"
        return sanitize_terminal(f"{host}{path}")[:160]
    except (TypeError, ValueError):
        return "unknown-url"


class TeamsSession:
    """A browser process owned by this session plus captured tokens.

    Use as an async context manager::

        async with TeamsSession() as s:
            events = await s.calendar(start, end)
    """

    def __init__(self) -> None:
        self.pw = None
        self.ctx = None
        self.teams = None
        self.tokens: dict[str, str] = {}
        self.bases: dict[str, str] = {}
        self.sp_pages: dict[str, object] = {}
        self.mcps_cache: dict[str, dict] = {}

    async def __aenter__(self):
        self.pw = await async_playwright().start()
        try:
            self.ctx = await launch_context(self.pw)
        except BaseException:
            await self.pw.stop()
            self.pw = None
            raise
        return self

    async def __aexit__(self, *exc):
        try:
            if self.ctx is not None:
                await self.ctx.close()
        finally:
            if self.pw is not None:
                await self.pw.stop()

    # -- token capture -------------------------------------------------------
    def _on_request(self, req) -> None:
        auth = req.headers.get("authorization")
        if not auth or not auth.startswith("Bearer "):
            return
        for key, rx in (("mt", MT_RE), ("mcps", MCPS_RE)):
            m = rx.match(req.url)
            if m:
                self.tokens[key] = auth
                self.bases[key] = m.group(1)

    async def teams_page(self):
        if self.teams is not None:
            return self.teams
        page = None
        for candidate in self.ctx.pages:
            try:
                validate_teams_url(candidate.url)
            except UrlValidationError:
                continue
            page = candidate
            break
        if page is None:
            page = await self.ctx.new_page()
            await self._goto(page, TEAMS_URL, self._allow_teams_or_auth_origin)
        page.on("request", self._on_request)
        self.teams = page
        return page

    @staticmethod
    def _allow_teams_or_auth_origin(url: str) -> None:
        try:
            validate_teams_url(url)
            return
        except UrlValidationError as exc:
            if is_approved_auth_url(url):
                return
            raise TranscriptError("The browser navigated to an unexpected Teams origin.") from exc

    @staticmethod
    def _allow_sharepoint_or_auth_origin(url: str, expected_host: str) -> None:
        try:
            validate_sharepoint_url(url, expected_host)
            return
        except UrlValidationError as exc:
            if is_approved_auth_url(url):
                return
            raise TranscriptError("The browser navigated to an unexpected SharePoint origin.") from exc

    @staticmethod
    async def _close_page(page) -> None:
        try:
            await page.close()
        except Exception:
            pass

    async def _goto(self, page, url: str, validate_final) -> None:
        try:
            response = await page.goto(url, wait_until="domcontentloaded")
        except Exception as exc:
            await self._close_page(page)
            raise TranscriptError(
                f"Browser navigation failed for {_safe_url_label(url)} ({type(exc).__name__})."
            ) from exc
        try:
            response_url = getattr(response, "url", "")
            if response_url:
                validate_final(response_url)
            validate_final(page.url)
        except Exception as exc:
            await self._close_page(page)
            raise TranscriptError(
                f"Browser navigation ended at an unexpected origin for {_safe_url_label(url)}."
            ) from exc

    @staticmethod
    def _sharepoint_host(host: str) -> str:
        try:
            return validate_sharepoint_host(host)
        except UrlValidationError as exc:
            raise TranscriptError("The transcript returned an invalid SharePoint host.") from exc

    async def _dismiss_launcher(self, page) -> None:
        for txt in ("Use the web app instead", "Use the web app"):
            try:
                btn = page.get_by_text(txt).first
                if await btn.is_visible(timeout=300):
                    await btn.click()
                    return
            except Exception:
                pass

    async def _wait_token(self, key: str, seconds: int) -> bool:
        page = await self.teams_page()
        for _ in range(seconds * 2):
            self._allow_teams_or_auth_origin(page.url)
            if key in self.tokens:
                return True
            await self._dismiss_launcher(page)
            await asyncio.sleep(0.5)
        return key in self.tokens

    async def ensure_mt_token(self) -> None:
        page = await self.teams_page()
        if await self._wait_token("mt", 3):
            return
        log("Waiting for Teams to sign in and load the calendar ...")
        await self._goto(page, TEAMS_URL, self._allow_teams_or_auth_origin)
        if not await self._wait_token("mt", 90):
            raise TranscriptError("Could not capture a Teams token. Is Teams signed in inside the dedicated browser window?")

    async def ensure_mcps_token(self, thread_id: str = "", ical: str = "") -> None:
        """Get the meeting-content token.

        Teams issues it while the calendar loads. If it does not appear, opening a
        meeting recap always requests one, so a few recent meetings are tried.
        """
        if await self._wait_token("mcps", 2):
            return
        await self.ensure_mt_token()
        page = await self.teams_page()
        if "calendar" not in (await page.title()).lower():
            await self._goto(page, TEAMS_URL + "v2/#/calendarv2", self._allow_teams_or_auth_origin)
        log("Waiting for the meeting-content token ...")
        if await self._wait_token("mcps", 30):
            return
        tenant = jwt_claims(self.tokens.get("mt", "")).get("tid", "")
        candidates = [(thread_id, ical)] if thread_id and ical else []
        now = datetime.now(timezone.utc)
        for e in reversed(await self.calendar(now - timedelta(days=60), now)):
            t = thread_id_from(e.get("skypeTeamsMeetingUrl", ""))
            if e.get("isOnlineMeeting") and t and e.get("iCalUID") and e.get("endTime", "") < now.isoformat():
                candidates.append((t, e["iCalUID"]))
            if len(candidates) >= 4:
                break
        for t, ic in candidates:
            log("Opening a meeting recap to obtain the meeting-content token ...")
            await self._goto(page, web_recap_url(t, ic, tenant), self._allow_teams_or_auth_origin)
            if await self._wait_token("mcps", 40):
                return
        raise TranscriptError(
            "Could not capture the meeting-content token. Open any meeting Recap tab in the browser window, then try again."
        )

    # -- Teams API ----------------------------------------------------------
    async def _fetch(self, page, url: str, headers: dict) -> dict:
        try:
            request = urllib.parse.urlsplit(url)
            if request.hostname in APPROVED_TEAMS_HOSTS:
                request_host = validate_teams_url(url)
                page_host = validate_teams_url(page.url)
            else:
                request_host = validate_sharepoint_url(url)
                page_host = validate_sharepoint_url(page.url, request_host)
        except (UrlValidationError, ValueError) as exc:
            raise TranscriptError("Refusing a request after an unexpected browser redirect or URL.") from exc
        if request_host != page_host:
            raise TranscriptError("Refusing a request after an unexpected browser redirect or URL.")
        try:
            response = await page.evaluate(FETCH_JS, [url, headers])
        except Exception as exc:
            raise TranscriptError(
                f"Browser request failed for {_safe_url_label(url)} ({type(exc).__name__})."
            ) from exc
        if not isinstance(response, dict) or not isinstance(response.get("url"), str):
            raise TranscriptError("The browser returned a response without a valid final URL.")
        try:
            if request.hostname in APPROVED_TEAMS_HOSTS:
                response_host = validate_teams_url(response["url"])
            else:
                response_host = validate_sharepoint_url(response["url"], request_host)
        except (UrlValidationError, ValueError) as exc:
            raise TranscriptError("Refusing a response after an unexpected redirect origin.") from exc
        if response_host != request_host:
            raise TranscriptError("Refusing a response after an unexpected redirect origin.")
        return response

    async def calendar(self, start: datetime, end: datetime) -> list[dict]:
        """Calendar events between two instants, following pagination."""
        await self.ensure_mt_token()
        page = await self.teams_page()
        base = self.bases["mt"]
        items: list[dict] = []
        skip, top = 0, 200
        while True:
            q = {
                "startDate": iso_z(start),
                "endDate": iso_z(end),
                "$top": top,
                "$count": "true",
                "$skip": skip,
                "$orderby": "startTime asc",
                "$filter": "isAppointment eq false and isAllDayEvent eq false",
            }
            url = f"{base}/v2.1/me/calendars/calendarView?" + urllib.parse.urlencode(q, quote_via=urllib.parse.quote)
            r = await self._fetch(page, url, {"authorization": self.tokens["mt"], "x-ms-migration": "True"})
            if r["status"] != 200:
                raise TranscriptError(f"Reading the calendar failed: HTTP {r['status']} {r['text'][:200]}")
            batch = json.loads(r["text"]).get("value", [])
            items.extend(batch)
            skip += len(batch)
            # The response's own count is the page size, not the total.
            if len(batch) < top:
                break
        return items

    async def meeting_contents(self, thread_id: str, ical: str = "") -> dict:
        """Recap resources for a meeting: transcripts, summaries, chat files."""
        if thread_id in self.mcps_cache:
            return self.mcps_cache[thread_id]
        await self.ensure_mcps_token(thread_id, ical)
        page = await self.teams_page()
        url = f"{self.bases['mcps']}/contents/?threadId={urllib.parse.quote(thread_id, safe='')}"
        r = await self._fetch(page, url, {"authorization": self.tokens["mcps"]})
        if r["status"] == 401:
            self.tokens.pop("mcps", None)
            await self.ensure_mcps_token(thread_id, ical)
            r = await self._fetch(page, url, {"authorization": self.tokens["mcps"]})
        if r["status"] == 404:
            data = {"resources": []}  # no recap content for this meeting
        elif r["status"] != 200:
            log(f"  meeting contents failed for {thread_id}: HTTP {r['status']} {r['text'][:120]}")
            data = {"resources": []}
        else:
            data = json.loads(r["text"])
        self.mcps_cache[thread_id] = data
        return data

    async def event_details(self, object_id: str = "", ical: str = "") -> dict:
        """Full calendar event: attendees with response status, location, invitation body."""
        if not object_id and not ical:
            return {}
        await self.ensure_mt_token()
        page = await self.teams_page()
        if object_id:
            path = f"events/{quote_path_component(object_id)}"
        else:
            path = f"events/iCalUId/{quote_path_component(ical)}"
        url = f"{self.bases['mt']}/v2.0/me/calendars/{path}?shouldDecryptData=true"
        r = await self._fetch(page, url, {"authorization": self.tokens["mt"], "x-ms-migration": "True"})
        if r["status"] != 200:
            log(f"  event details unavailable: HTTP {r['status']}")
            return {}
        return json.loads(r["text"])

    # -- SharePoint ---------------------------------------------------------
    async def sharepoint_page(self, host: str):
        """A tab on the SharePoint host, used for its cookies."""
        host = self._sharepoint_host(host)
        if host in self.sp_pages:
            return self.sp_pages[host]
        page = await self.ctx.new_page()
        await self._goto(
            page,
            f"https://{host}/",
            lambda final_url: self._allow_sharepoint_or_auth_origin(final_url, host),
        )
        for _ in range(60):
            try:
                validate_sharepoint_url(page.url, host)
            except UrlValidationError:
                if not is_approved_auth_url(page.url):
                    await self._close_page(page)
                    raise TranscriptError("SharePoint redirected to an unexpected origin.")
                await asyncio.sleep(1)
                continue
            if "login" not in page.url.lower():
                r = await self._fetch(page, f"https://{host}/_api/v2.1/me?select=id", {"accept": "application/json"})
                if r["status"] in (200, 400, 404):
                    self.sp_pages[host] = page
                    return page
            await asyncio.sleep(1)
        await self._close_page(page)
        raise TranscriptError("SharePoint did not return to the expected signed-in tenant origin.")

    async def list_transcripts(self, host: str, drive_id: str, item_id: str) -> list[dict]:
        host = self._sharepoint_host(host)
        page = await self.sharepoint_page(host)
        url = (
            f"https://{host}/_api/v2.1/drives/{quote_path_component(drive_id)}"
            f"/items/{quote_path_component(item_id)}/versions/current/media/transcripts"
        )
        r = await self._fetch(page, url, {"accept": "application/json"})
        if r["status"] != 200:
            raise TranscriptError(f"Listing transcripts failed: HTTP {r['status']} {r['text'][:200]}")
        return json.loads(r["text"]).get("value", [])

    async def item_created(self, host: str, drive_id: str, item_id: str) -> str:
        host = self._sharepoint_host(host)
        page = await self.sharepoint_page(host)
        url = (
            f"https://{host}/_api/v2.1/drives/{quote_path_component(drive_id)}"
            f"/items/{quote_path_component(item_id)}?select=createdDateTime"
        )
        r = await self._fetch(page, url, {"accept": "application/json"})
        if r["status"] != 200:
            return ""
        return json.loads(r["text"]).get("createdDateTime", "")

    async def fetch_transcript(self, host: str, drive_id: str, item_id: str, transcript_id: str, fmt: str) -> str:
        """Transcript content as served to the recap page. fmt is 'json' or 'vtt'."""
        if fmt not in {"json", "vtt"}:
            raise TranscriptError("The transcript format is not supported.")
        host = self._sharepoint_host(host)
        page = await self.sharepoint_page(host)
        url = (
            f"https://{host}/_api/v2.1/drives/{quote_path_component(drive_id)}"
            f"/items/{quote_path_component(item_id)}"
            f"/media/transcripts/{quote_path_component(transcript_id)}"
            f"/content?format={quote_path_component(fmt)}"
        )
        r = await self._fetch(page, url, {})
        if r["status"] != 200:
            raise TranscriptError(f"Downloading the transcript failed: HTTP {r['status']} {r['text'][:200]}")
        return r["text"]
