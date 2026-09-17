"""Pure functions over Teams and SharePoint data: identifiers, dates, references.

Nothing here touches the network or a browser, which makes it the part that is
easy to test.
"""
from __future__ import annotations

import base64
import json
import re
import urllib.parse
from datetime import datetime, timedelta, timezone
from pathlib import Path

MT_RE = re.compile(r"^(https://teams\.cloud\.microsoft/api/mt/part/[^/]+)/")
MCPS_RE = re.compile(r"^(https://teams\.cloud\.microsoft/api/mcps/[^/]+)/")
THREAD_RE = re.compile(r"19(?::|%3a)meeting_[A-Za-z0-9_-]+(?:@|%40)thread\.v2", re.I)


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
    s = re.sub(r"[\\/:*?\"<>|\r\n\t]+", " ", s).strip()
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
            files.append(
                {
                    "title": md.get("fileTitle") or Path(urllib.parse.urlsplit(res["location"]).path).name,
                    "url": res["location"],
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
        host = urllib.parse.urlsplit(loc).netloc if loc else ""
        tid = md.get("transcriptId") or ""
        if not tid and "/transcripts/" in loc:
            tid = loc.split("/transcripts/")[1].split("/")[0]
        start = ticks_to_dt(md.get("startTime"))
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
                "webUrl": md.get("webUrl", ""),
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
    q = urllib.parse.parse_qs(urllib.parse.urlsplit(url).query)

    def g(k: str) -> str:
        return q.get(k, [""])[0]

    if not g("driveId") or not g("driveItemId"):
        return None
    site = g("sitePath")
    host = urllib.parse.urlsplit(site).netloc if site else urllib.parse.urlsplit(g("fileUrl")).netloc
    tid = site.split("/transcripts/")[1].split("/")[0] if "/transcripts/" in site else ""
    title = ""
    if g("fileUrl"):
        title = Path(urllib.parse.urlsplit(g("fileUrl")).path).name
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
