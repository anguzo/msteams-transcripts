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

from .browser import ensure_browser, log
from .config import SETTINGS, TEAMS_URL
from .errors import TranscriptError
from .model import MCPS_RE, MT_RE, iso_z, jwt_claims, thread_id_from, web_recap_url

try:
    from playwright.async_api import async_playwright
except ImportError as exc:  # pragma: no cover
    raise TranscriptError("Playwright is missing. Install this tool with 'pip install msteams-transcripts'.") from exc

FETCH_JS = """async ([url, headers]) => {
  const r = await fetch(url, {headers, credentials: 'include'});
  const text = await r.text();
  return {status: r.status, ct: r.headers.get('content-type') || '', text};
}"""


class TeamsSession:
    """An attached browser plus the tokens captured from it.

    Use as an async context manager::

        async with TeamsSession() as s:
            events = await s.calendar(start, end)
    """

    def __init__(self) -> None:
        self.pw = None
        self.browser = None
        self.ctx = None
        self.teams = None
        self.tokens: dict[str, str] = {}
        self.bases: dict[str, str] = {}
        self.sp_pages: dict[str, object] = {}
        self.mcps_cache: dict[str, dict] = {}

    async def __aenter__(self):
        ensure_browser()
        self.pw = await async_playwright().start()
        self.browser = await self.pw.chromium.connect_over_cdp(SETTINGS.cdp_url)
        self.ctx = self.browser.contexts[0] if self.browser.contexts else await self.browser.new_context()
        return self

    async def __aexit__(self, *exc):
        try:
            for pg in self.sp_pages.values():
                await pg.close()
        except Exception:
            pass
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
        page = next((p for p in self.ctx.pages if "teams.cloud.microsoft" in p.url), None)
        if page is None:
            page = await self.ctx.new_page()
            await page.goto(TEAMS_URL, wait_until="domcontentloaded")
        page.on("request", self._on_request)
        self.teams = page
        return page

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
        await page.goto(TEAMS_URL, wait_until="domcontentloaded")
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
            await page.goto(TEAMS_URL + "v2/#/calendarv2", wait_until="domcontentloaded")
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
            await page.goto(web_recap_url(t, ic, tenant), wait_until="domcontentloaded")
            if await self._wait_token("mcps", 40):
                return
        raise TranscriptError(
            "Could not capture the meeting-content token. Open any meeting Recap tab in the browser window, then try again."
        )

    # -- Teams API ----------------------------------------------------------
    async def _fetch(self, page, url: str, headers: dict) -> dict:
        return await page.evaluate(FETCH_JS, [url, headers])

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
            path = f"events/{urllib.parse.quote(object_id, safe='')}"
        else:
            path = f"events/iCalUId/{ical}"
        url = f"{self.bases['mt']}/v2.0/me/calendars/{path}?shouldDecryptData=true"
        r = await self._fetch(page, url, {"authorization": self.tokens["mt"], "x-ms-migration": "True"})
        if r["status"] != 200:
            log(f"  event details unavailable: HTTP {r['status']}")
            return {}
        return json.loads(r["text"])

    # -- SharePoint ---------------------------------------------------------
    async def sharepoint_page(self, host: str):
        """A tab on the SharePoint host, used for its cookies."""
        if host in self.sp_pages:
            return self.sp_pages[host]
        page = await self.ctx.new_page()
        await page.goto(f"https://{host}/", wait_until="domcontentloaded")
        for _ in range(60):
            if host in page.url and "login" not in page.url:
                r = await self._fetch(page, f"https://{host}/_api/v2.1/me?select=id", {"accept": "application/json"})
                if r["status"] in (200, 400, 404):
                    break
            await asyncio.sleep(1)
        self.sp_pages[host] = page
        return page

    async def list_transcripts(self, host: str, drive_id: str, item_id: str) -> list[dict]:
        page = await self.sharepoint_page(host)
        url = f"https://{host}/_api/v2.1/drives/{drive_id}/items/{item_id}/versions/current/media/transcripts"
        r = await self._fetch(page, url, {"accept": "application/json"})
        if r["status"] != 200:
            raise TranscriptError(f"Listing transcripts failed: HTTP {r['status']} {r['text'][:200]}")
        return json.loads(r["text"]).get("value", [])

    async def item_created(self, host: str, drive_id: str, item_id: str) -> str:
        page = await self.sharepoint_page(host)
        url = f"https://{host}/_api/v2.1/drives/{drive_id}/items/{item_id}?select=createdDateTime"
        r = await self._fetch(page, url, {"accept": "application/json"})
        if r["status"] != 200:
            return ""
        return json.loads(r["text"]).get("createdDateTime", "")

    async def fetch_transcript(self, host: str, drive_id: str, item_id: str, transcript_id: str, fmt: str) -> str:
        """Transcript content as served to the recap page. fmt is 'json' or 'vtt'."""
        page = await self.sharepoint_page(host)
        url = (
            f"https://{host}/_api/v2.1/drives/{drive_id}/items/{item_id}"
            f"/media/transcripts/{transcript_id}/content?format={fmt}"
        )
        r = await self._fetch(page, url, {})
        if r["status"] != 200:
            raise TranscriptError(f"Downloading the transcript failed: HTTP {r['status']} {r['text'][:200]}")
        return r["text"]
