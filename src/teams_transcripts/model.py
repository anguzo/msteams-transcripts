"""Pure functions over Teams and SharePoint data: identifiers, dates, references.

Nothing here touches the network or a browser, which makes it the part that is
easy to test.
"""
from __future__ import annotations

import base64
import ipaddress
import json
import re
import unicodedata
import urllib.parse
from datetime import datetime, timedelta, timezone
from pathlib import Path

from .errors import TranscriptError

MT_RE = re.compile(r"^(https://teams\.cloud\.microsoft/api/mt/part/[^/]+)/")
MCPS_RE = re.compile(r"^(https://teams\.cloud\.microsoft/api/mcps/[^/]+)/")
THREAD_RE = re.compile(r"19(?::|%3a)meeting_[A-Za-z0-9_-]+(?:@|%40)thread\.v2", re.I)
APPROVED_TEAMS_HOSTS = frozenset({"teams.cloud.microsoft", "teams.microsoft.com"})
APPROVED_AUTH_HOSTS = frozenset(
    {
        "login.microsoftonline.com",
    }
)
SHAREPOINT_SUFFIXES = (
    "sharepoint.com",
)
_HOST_LABEL_RE = re.compile(r"^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$")


class UrlValidationError(TranscriptError, ValueError):
    """An externally supplied URL or host is outside the supported origins."""


def _parse_https_url(url: str, what: str) -> urllib.parse.SplitResult:
    if not isinstance(url, str) or not url or any(unicodedata.category(c) == "Cc" for c in url):
        raise UrlValidationError(f"The {what} must be a valid HTTPS URL.")
    try:
        parsed = urllib.parse.urlsplit(url)
        if parsed.scheme.lower() != "https" or not parsed.netloc:
            raise UrlValidationError(f"The {what} must be a valid HTTPS URL.")
        if parsed.username is not None or parsed.password is not None:
            raise UrlValidationError(f"The {what} must not contain a username or password.")
        if ":" in parsed.netloc:
            raise UrlValidationError(f"The {what} must not contain an explicit port.")
        # Accessing .port rejects malformed values and also lets us reject an
        # explicit port, including the otherwise harmless :443.
        if parsed.port is not None:
            raise UrlValidationError(f"The {what} must not contain an explicit port.")
        if not parsed.hostname:
            raise UrlValidationError(f"The {what} must have a valid host.")
    except (ValueError, UnicodeError) as exc:
        if isinstance(exc, UrlValidationError):
            raise
        raise UrlValidationError(f"The {what} must be a valid HTTPS URL.") from exc
    return parsed


def validate_teams_host(host: str) -> str:
    """Return a normalised approved Teams host or reject it."""
    try:
        parsed = urllib.parse.urlsplit(f"//{host}")
        if (
            not isinstance(host, str)
            or not host
            or parsed.path
            or parsed.query
            or parsed.fragment
            or ":" in host
            or parsed.username is not None
            or parsed.password is not None
            or parsed.port is not None
            or not parsed.hostname
        ):
            raise UrlValidationError("The Teams host is not approved.")
        hostname = parsed.hostname.lower()
    except (ValueError, UnicodeError) as exc:
        if isinstance(exc, UrlValidationError):
            raise
        raise UrlValidationError("The Teams host is not approved.") from exc
    if hostname not in APPROVED_TEAMS_HOSTS:
        raise UrlValidationError("The URL must use an approved Teams host.")
    return hostname


def validate_teams_url(url: str) -> str:
    """Validate an HTTPS URL whose origin is an approved Teams host."""
    parsed = _parse_https_url(url, "Teams URL")
    return validate_teams_host(parsed.netloc)


def is_approved_auth_url(url: str) -> bool:
    """Whether a browser URL is a permitted commercial Entra auth origin."""
    try:
        parsed = urllib.parse.urlsplit(url)
        return (
            parsed.scheme.lower() == "https"
            and parsed.hostname in APPROVED_AUTH_HOSTS
            and parsed.username is None
            and parsed.password is None
            and parsed.port is None
        )
    except (TypeError, ValueError):
        return False


def validate_sharepoint_host(host: str) -> str:
    """Return a normalised tenant SharePoint host or reject it."""
    try:
        parsed = urllib.parse.urlsplit(f"//{host}")
        if (
            not isinstance(host, str)
            or not host
            or parsed.path
            or parsed.query
            or parsed.fragment
            or ":" in host
            or parsed.username is not None
            or parsed.password is not None
            or parsed.port is not None
            or not parsed.hostname
        ):
            raise UrlValidationError("The SharePoint host is not valid.")
        hostname = parsed.hostname.lower()
    except (ValueError, UnicodeError) as exc:
        if isinstance(exc, UrlValidationError):
            raise
        raise UrlValidationError("The SharePoint host is not valid.") from exc

    try:
        ipaddress.ip_address(hostname)
    except ValueError:
        pass
    else:
        raise UrlValidationError("The SharePoint host must be a tenant host, not an IP address.")

    if hostname == "localhost" or hostname.endswith(".localhost"):
        raise UrlValidationError("The SharePoint host must not be localhost.")
    if any(unicodedata.category(c) == "Cc" for c in hostname):
        raise UrlValidationError("The SharePoint host is not valid.")
    if not re.fullmatch(r"[a-z0-9.-]+", hostname):
        raise UrlValidationError("The SharePoint host is not valid.")

    for suffix in SHAREPOINT_SUFFIXES:
        marker = f".{suffix}"
        if not hostname.endswith(marker):
            continue
        tenant = hostname[: -len(marker)]
        if "." not in tenant and _HOST_LABEL_RE.fullmatch(tenant):
            return hostname
    raise UrlValidationError("The SharePoint host is not a valid tenant host.")


def validate_sharepoint_url(url: str, expected_host: str = "") -> str:
    """Validate an HTTPS SharePoint URL and optionally require its exact host."""
    parsed = _parse_https_url(url, "SharePoint URL")
    host = validate_sharepoint_host(parsed.netloc)
    if expected_host and host != validate_sharepoint_host(expected_host):
        raise UrlValidationError("The SharePoint URL changed to an unexpected host.")
    return host


def validate_recap_url(url: str) -> str:
    """Validate a user-supplied recap URL before reading its query values."""
    parsed = _parse_https_url(url, "recap URL")
    validate_teams_host(parsed.netloc)
    return url


def sanitize_terminal(value: object) -> str:
    """Remove terminal control characters from user- or service-derived text."""
    return "".join(c for c in str(value) if unicodedata.category(c) not in {"Cc", "Cf"})


def quote_path_component(value: object) -> str:
    """Quote one externally derived URL path component, including slashes."""
    return urllib.parse.quote(str(value), safe="")


def jwt_claims(bearer: str) -> dict:
    """Claims of a bearer token, or an empty dict if it cannot be read."""
    try:
        payload = bearer.split(" ", 1)[1].split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload))
    except Exception:
        return {}


def ticks_to_dt(ticks) -> datetime | None:
    """.NET ticks (100 ns since year 1) to an aware datetime."""
    try:
        t = int(ticks)
        if t <= 0:
            return None
        return datetime(1, 1, 1, tzinfo=timezone.utc) + timedelta(microseconds=t // 10)
    except Exception:
        return None


def parse_date(s: str, end: bool = False) -> datetime:
    """'2026-09-01' or 'today' to a UTC midnight boundary."""
    if s.lower() == "today":
        d = datetime.now(timezone.utc).date()
    else:
        d = datetime.strptime(s, "%Y-%m-%d").date()
    dt = datetime(d.year, d.month, d.day, tzinfo=timezone.utc)
    return dt + timedelta(days=1) if end else dt


def iso_z(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")


def safe_name(s: str, limit: int = 80) -> str:
    """A file name that Windows and POSIX both accept."""
    s = re.sub(r"[\\/:*?\"<>|\r\n\t]+", " ", sanitize_terminal(s)).strip()
    s = re.sub(r"\s+", " ", s)
    return s[:limit].rstrip(" .") or "meeting"


def thread_id_from(text: str) -> str | None:
    """Pull a meeting thread id out of a join link, a recap link, or raw text.

    The middle of the identifier is case-sensitive base64 and must be preserved
    exactly; only the prefix and suffix are normalised.
    """
    m = THREAD_RE.search(text or "")
    if not m:
        return None
    tid = urllib.parse.unquote(m.group(0))
    return "19:" + tid[3:-10] + "@thread.v2"


def shared_files_from_contents(contents: dict) -> list[dict]:
    """Files posted in the meeting chat, as listed by the meeting-content service."""
    files = []
    for res in contents.get("resources", []):
        if str(res.get("type", "")) == "MeetingChat" and res.get("location"):
            md = res.get("metadata", {}) or {}
            location = res["location"]
            validate_sharepoint_url(location)
            files.append(
                {
                    "title": md.get("fileTitle") or Path(urllib.parse.urlsplit(location).path).name,
                    "url": location,
                    "type": md.get("fileType", ""),
                }
            )
    return files


def transcript_refs_from_contents(contents: dict, subject: str = "") -> list[dict]:
    """Transcript references (drive, item, transcript id, host) from a meeting-content response."""
    refs = []
    ical_any = next(
        (r.get("metadata", {}).get("iCalUid", "") for r in contents.get("resources", []) if (r.get("metadata") or {}).get("iCalUid")),
        "",
    )
    for res in contents.get("resources", []):
        if "transcript" not in str(res.get("type", "")).lower():
            continue
        md = res.get("metadata", {}) or {}
        loc = res.get("location") or ""
        host = validate_sharepoint_url(loc) if loc else ""
        tid = md.get("transcriptId") or ""
        if not tid and loc:
            parts = [urllib.parse.unquote(part) for part in urllib.parse.urlsplit(loc).path.split("/")]
            try:
                index = next(i for i, part in enumerate(parts) if part.lower() == "transcripts")
            except StopIteration:
                pass
            else:
                if index + 1 < len(parts):
                    tid = parts[index + 1]
        start = ticks_to_dt(md.get("startTime"))
        web_url = md.get("webUrl", "")
        if web_url:
            validate_sharepoint_url(web_url, host)
        refs.append(
            {
                "subject": subject or md.get("fileTitle") or "",
                "fileTitle": md.get("fileTitle", ""),
                "driveId": md.get("driveId", ""),
                "driveItemId": md.get("driveItemId", ""),
                "transcriptId": tid,
                "host": host,
                "location": loc,
                "start": start.isoformat() if start else "",
                "end": (ticks_to_dt(md.get("endTime")) or start).isoformat() if start else "",
                "webUrl": web_url,
                "iCalUID": md.get("iCalUid") or ical_any,
                "threadId": md.get("threadId", ""),
            }
        )
    return refs


def refs_from_recap_url(url: str) -> dict | None:
    """Everything needed to download a transcript, taken from a recap link.

    Returns None when the link carries no drive identifiers, in which case the
    caller should fall back to looking the meeting up by its thread id.
    """
    validate_recap_url(url)
    q = urllib.parse.parse_qs(urllib.parse.urlsplit(url).query)

    def g(k: str) -> str:
        return q.get(k, [""])[0]

    if not g("driveId") or not g("driveItemId"):
        return None
    site = g("sitePath")
    file_url = g("fileUrl")
    site_host = validate_sharepoint_url(site) if site else ""
    file_host = validate_sharepoint_url(file_url) if file_url else ""
    if site_host and file_host and site_host != file_host:
        raise UrlValidationError("The recap locations use different SharePoint hosts.")
    # Some recap links carry only drive identifiers. Keep the host empty so a
    # validated --sharepoint-host override can supply it at download time.
    host = site_host or file_host
    tid = ""
    if site:
        parts = [urllib.parse.unquote(part) for part in urllib.parse.urlsplit(site).path.split("/")]
        try:
            index = next(i for i, part in enumerate(parts) if part.lower() == "transcripts")
        except StopIteration:
            pass
        else:
            if index + 1 < len(parts):
                tid = parts[index + 1]
    title = ""
    if file_url:
        title = Path(urllib.parse.urlsplit(file_url).path).name
        title = urllib.parse.unquote(title).rsplit(".", 1)[0]
    return {
        "subject": title,
        "fileTitle": title,
        "driveId": g("driveId"),
        "driveItemId": g("driveItemId"),
        "transcriptId": tid,
        "host": host,
        "location": site,
        "start": "",
        "end": "",
        "webUrl": g("fileUrl"),
        "iCalUID": g("iCalUid"),
        "threadId": thread_id_from(g("threadId")) or "",
    }


def web_recap_url(thread_id: str, ical: str, tenant: str = "") -> str:
    """A recap link that opens in the web client instead of the desktop-app prompt.

    The "/_#/l/" form bypasses the launcher page that asks whether to open the
    Teams desktop app; the documented suppression flags do not.
    """
    q = {"threadId": thread_id, "iCalUid": ical, "threadType": "meeting", "meetingType": "Scheduled"}
    if tenant:
        q["tenantId"] = tenant
    return "https://teams.cloud.microsoft/_#/l/meetingrecap?" + urllib.parse.urlencode(q, quote_via=urllib.parse.quote)
