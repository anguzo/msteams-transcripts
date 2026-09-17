"""Writing a transcript to disk, optionally with the meeting details."""
from __future__ import annotations

import json
import re
from pathlib import Path

from .browser import log
from .errors import TranscriptError
from .formats import invite_description, speakers_of, transcript_to_text
from .model import safe_name, shared_files_from_contents
from .session import TeamsSession

# Recording file names look like "<subject>-20260903_141403-Transcription....mp4"
RECORDING_NAME_RE = re.compile(r"^(.*?)-(\d{8})_(\d{6})(?:-.*)?$")


async def collect_details(s: TeamsSession, ref: dict) -> dict:
    """Calendar event and shared files for a transcript reference."""
    details = {"event": await s.event_details(ref.get("objectId", ""), ref.get("iCalUID", ""))}
    if ref.get("threadId"):
        try:
            details["files"] = shared_files_from_contents(await s.meeting_contents(ref["threadId"], ref.get("iCalUID", "")))
        except TranscriptError:
            details["files"] = []
    return details


def _meta_document(ref: dict, det: dict, doc: dict, subject: str, when: str) -> dict:
    ev = det.get("event") or {}
    return {
        "subject": subject,
        "start": when,
        "organizer": {"name": ev.get("organizerName", ""), "email": ev.get("organizerAddress", "")},
        "location": ev.get("location", ""),
        "attendees": ev.get("attendees", []),
        "speakers": [{"name": n, "segments": c} for n, c in speakers_of(doc)],
        "files": det.get("files", []),
        "description": invite_description(ev.get("bodyContent", "")),
        "threadId": ref.get("threadId", ""),
        "iCalUID": ref.get("iCalUID", ""),
        "source": ref.get("webUrl") or ref.get("location") or "",
    }


async def download_ref(
    s: TeamsSession,
    ref: dict,
    out_dir: Path,
    fmt: str = "txt",
    details: bool = False,
) -> list[Path]:
    """Download one transcript reference. Returns the files written.

    fmt is 'txt', 'json', 'vtt' or 'all'. With details, a header is added to the
    text file and a '.meta.json' sidecar is written alongside it.
    """
    host, drive, item = ref["host"], ref["driveId"], ref["driveItemId"]
    if not host:
        raise TranscriptError("No SharePoint host is known for this transcript; pass --sharepoint-host.")
    tids = [ref["transcriptId"]] if ref.get("transcriptId") else [t["id"] for t in await s.list_transcripts(host, drive, item)]
    if not tids:
        log(f"  no transcript found for {ref.get('subject')}")
        return []

    subject = ref.get("subject") or ref.get("fileTitle") or ""
    name_match = RECORDING_NAME_RE.match(subject)
    if name_match:
        subject = name_match.group(1)
    when = (ref.get("start") or ref.get("meetingStart") or "")[:16].replace("T", " ")
    if not when:
        # The recording's creation time is UTC, matching the times shown by 'list'.
        when = (await s.item_created(host, drive, item) or "")[:16].replace("T", " ")
    if not when and name_match:
        d, t = name_match.group(2), name_match.group(3)
        when = f"{d[:4]}-{d[4:6]}-{d[6:]} {t[:2]}:{t[2:4]}"

    det = await collect_details(s, ref) if details else None
    if det and det.get("event", {}).get("subject"):
        subject = det["event"]["subject"]

    stem = safe_name(f"{when.replace(':', '')} - {subject or 'meeting'}".strip(" -"))
    out_dir.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []
    for i, tid in enumerate(tids):
        suffix = f" ({i + 1})" if len(tids) > 1 else ""
        base = out_dir / (stem + suffix)
        doc_json = await s.fetch_transcript(host, drive, item, tid, "json")
        doc = json.loads(doc_json)
        source = ref.get("webUrl") or ref.get("location") or ""
        if fmt in ("txt", "all"):
            p = base.with_suffix(".txt")
            p.write_text(transcript_to_text(doc, subject or stem, when or "unknown", source, det), encoding="utf-8")
            written.append(p)
        if det:
            p = base.with_suffix(".meta.json")
            p.write_text(json.dumps(_meta_document(ref, det, doc, subject, when), indent=1, ensure_ascii=False), encoding="utf-8")
            written.append(p)
        if fmt in ("json", "all"):
            p = base.with_suffix(".json")
            p.write_text(doc_json, encoding="utf-8")
            written.append(p)
        if fmt in ("vtt", "all"):
            p = base.with_suffix(".vtt")
            p.write_text(await s.fetch_transcript(host, drive, item, tid, "vtt"), encoding="utf-8")
            written.append(p)
        log(f"  saved {written[-1].name}  ({len(doc.get('entries', []))} entries)")
    return written
