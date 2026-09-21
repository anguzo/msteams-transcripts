"""Command line: browser, list, get, batch."""

from __future__ import annotations

import argparse
import asyncio
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from . import __version__
from .browser import keep_browser_open, log
from .config import PROG, SETTINGS, configure
from .download import download_ref
from .errors import TranscriptError
from .model import (
    parse_date,
    refs_from_recap_url,
    sanitize_terminal,
    thread_id_from,
    transcript_refs_from_contents,
    validate_sharepoint_host,
)
from .session import TeamsSession
from .storage import atomic_write_text, ensure_private_dir, read_text_without_symlink

DESCRIPTION = """\
Download Microsoft Teams meeting transcripts from the command line.

The list, get and batch commands launch a dedicated browser process for their
session and close it when finished. Use the 'browser' command to sign in
interactively first; it remains open until you close it or interrupt it.
"""

EPILOG = f"""\
examples:
  {PROG} browser
  {PROG} list --from 2026-09-01 --to today
  {PROG} get 3 --details
  {PROG} get "https://teams.cloud.microsoft/l/meetingrecap?driveId=...&driveItemId=..." -f all
  {PROG} batch --from 2026-09-01 --to 2026-09-30 --details -o ./transcripts
"""


def add_deprecated_port_option(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--port",
        type=int,
        default=argparse.SUPPRESS,
        help="deprecated and ignored; no TCP CDP port is opened",
    )


def sharepoint_host_arg(value: str) -> str:
    try:
        return validate_sharepoint_host(value)
    except TranscriptError as exc:
        raise argparse.ArgumentTypeError(sanitize_terminal(str(exc))) from exc


def default_out() -> str:
    return str(Path.cwd() / "transcripts")


async def cmd_list(args) -> list[dict]:
    """Print one row per transcript and cache the rows for 'get N'."""
    start, end = parse_date(args.frm), parse_date(args.to, end=True)
    async with TeamsSession() as s:
        events = await s.calendar(start, end)
        now = datetime.now(timezone.utc)
        meetings = []
        for e in events:
            if not e.get("isOnlineMeeting") or e.get("isCancelled"):
                continue
            tid = thread_id_from(e.get("skypeTeamsMeetingUrl", "")) or thread_id_from(
                json.dumps(e.get("skypeTeamsDataObject") or e.get("skypeTeamsData") or "")
            )
            if not tid:
                continue
            try:
                ends = datetime.fromisoformat(e.get("endTime", "").replace("Z", "+00:00"))
            except Exception:
                ends = now
            if ends > now:
                continue
            meetings.append(
                {
                    "subject": e.get("subject", ""),
                    "start": e.get("startTime", ""),
                    "end": e.get("endTime", ""),
                    "threadId": tid,
                    "iCalUID": e.get("iCalUID", ""),
                    "organizer": e.get("organizerName", ""),
                    "objectId": e.get("objectId", ""),
                }
            )
        log(f"{len(events)} calendar events, {len(meetings)} past online meetings. Checking for transcripts ...")

        rows: list[dict] = []
        seen_threads: set[str] = set()
        for m in meetings:
            if m["threadId"] in seen_threads:
                continue
            seen_threads.add(m["threadId"])
            contents = await s.meeting_contents(m["threadId"], m["iCalUID"])
            same_thread = [x for x in meetings if x["threadId"] == m["threadId"]]
            for ref in transcript_refs_from_contents(contents, m["subject"]):
                # Match a transcript to the occurrence whose window contains its start.
                occ = next((x for x in same_thread if ref["start"] and x["start"] <= ref["start"] <= x["end"]), None)
                ref["subject"] = m["subject"]
                ref["meetingStart"] = (occ or m)["start"]
                ref["threadId"] = m["threadId"]
                ref["organizer"] = m["organizer"]
                ref["objectId"] = (occ or m)["objectId"]
                ref["iCalUID"] = (occ or m)["iCalUID"]
                if start.isoformat() <= (ref["start"] or ref["meetingStart"]) <= end.isoformat():
                    rows.append(ref)
        rows.sort(key=lambda r: r["start"] or r["meetingStart"])
        for i, r in enumerate(rows, 1):
            r["n"] = i

        ensure_private_dir(SETTINGS.state_dir)
        atomic_write_text(SETTINGS.last_list, json.dumps(rows, indent=1))
        if args.json:
            print(json.dumps(rows, indent=1))
        else:
            print(f"{'#':>3}  {'Start (UTC)':<17} {'Subject':<50} Organizer")
            for r in rows:
                st = sanitize_terminal((r["start"] or r["meetingStart"])[:16].replace("T", " "))
                subject = sanitize_terminal(str(r["subject"]))[:50]
                organizer = sanitize_terminal(str(r["organizer"]))
                print(f"{r['n']:>3}  {st:<17} {subject:<50} {organizer}")
            log(f"{len(rows)} transcript(s). Use:  {PROG} get <#>   or   {PROG} batch --from ... --to ...")
        return rows


async def _download_by_thread(s: TeamsSession, thread_id: str, args, sharepoint_host: str = "") -> None:
    refs = transcript_refs_from_contents(await s.meeting_contents(thread_id))
    if not refs:
        raise TranscriptError("No transcript was found for that meeting.")
    for r in refs:
        if sharepoint_host:
            r["host"] = sharepoint_host
        await download_ref(s, r, Path(args.out), args.format, args.details)


async def cmd_get(args) -> None:
    """Download one transcript, named by row number, recap link or thread id."""
    target: str = args.target
    sharepoint_host = validate_sharepoint_host(args.sharepoint_host) if args.sharepoint_host else ""
    async with TeamsSession() as s:
        if target.isdigit():
            if not SETTINGS.last_list.exists():
                raise TranscriptError(f"No previous list. Run:  {PROG} list --from ... --to ...")
            rows = json.loads(read_text_without_symlink(SETTINGS.last_list))
            ref = next((r for r in rows if r["n"] == int(target)), None)
            if not ref:
                raise TranscriptError(f"There is no row {target} in the last list.")
        elif "://" in target:
            ref = refs_from_recap_url(target)
            if ref is None:
                tid = thread_id_from(target)
                if not tid:
                    raise TranscriptError("That link has neither drive identifiers nor a meeting thread id.")
                await _download_by_thread(s, tid, args, sharepoint_host)
                return
        else:
            await _download_by_thread(s, thread_id_from(target) or target, args, sharepoint_host)
            return
        if sharepoint_host:
            ref["host"] = sharepoint_host
        await download_ref(s, ref, Path(args.out), args.format, args.details)


async def cmd_batch(args) -> None:
    """Download every transcript in a date range."""
    rows = await cmd_list(argparse.Namespace(frm=args.frm, to=args.to, json=False))
    if args.only:
        want = {int(x) for x in re.split(r"[,\s]+", args.only.strip()) if x}
        rows = [r for r in rows if r["n"] in want]
    if not rows:
        return
    async with TeamsSession() as s:
        ok = fail = 0
        for r in rows:
            try:
                await download_ref(s, r, Path(args.out), args.format, args.details)
                ok += 1
            except Exception as ex:
                fail += 1
                log(f"  FAILED {r.get('subject')}: {ex}")
        log(f"Done: {ok} downloaded, {fail} failed. Output: {Path(args.out).resolve()}")


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog=PROG,
        description=DESCRIPTION,
        epilog=EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    ap.add_argument("--port", type=int, default=None, help="deprecated and ignored; no TCP CDP port is opened")
    ap.add_argument("--profile", default="", help="browser profile directory (or TT_PROFILE_DIR)")
    ap.add_argument("--browser", default="", help="path to the Edge or Chrome executable (or TT_BROWSER_EXE)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    browser_parser = sub.add_parser("browser", help="open the dedicated browser for interactive sign-in")
    add_deprecated_port_option(browser_parser)

    p = sub.add_parser("list", help="list past online meetings that have transcripts")
    add_deprecated_port_option(p)
    p.add_argument("--from", dest="frm", required=True, metavar="YYYY-MM-DD")
    p.add_argument("--to", required=True, metavar="YYYY-MM-DD", help="a date, or 'today'")
    p.add_argument("--json", action="store_true", help="print JSON instead of a table")

    p = sub.add_parser("get", help="download one transcript")
    add_deprecated_port_option(p)
    p.add_argument("target", help="row number from 'list', a recap URL, or 19:meeting_...@thread.v2")
    p.add_argument("-o", "--out", default=None, help="output directory (default ./transcripts)")
    p.add_argument("-f", "--format", choices=["txt", "json", "vtt", "all"], default="txt")
    p.add_argument(
        "--sharepoint-host",
        type=sharepoint_host_arg,
        default=None,
        help="validated commercial tenant host override, e.g. contoso-my.sharepoint.com",
    )
    p.add_argument(
        "-d",
        "--details",
        action="store_true",
        help="add organizer, location, invited attendees, speakers, shared files and invitation text; also writes .meta.json",
    )

    p = sub.add_parser("batch", help="download every transcript in a date range")
    add_deprecated_port_option(p)
    p.add_argument("--from", dest="frm", required=True, metavar="YYYY-MM-DD")
    p.add_argument("--to", required=True, metavar="YYYY-MM-DD")
    p.add_argument("-o", "--out", default=None, help="output directory (default ./transcripts)")
    p.add_argument("-f", "--format", choices=["txt", "json", "vtt", "all"], default="txt")
    p.add_argument("-d", "--details", action="store_true", help="same as for 'get'")
    p.add_argument("--only", default="", help="comma-separated row numbers instead of all, e.g. 3,7,12")
    return ap


def main(argv: list[str] | None = None) -> None:
    args = build_parser().parse_args(argv)
    configure(port=args.port, profile=args.profile, browser=args.browser)
    if getattr(args, "out", None) is None:
        args.out = default_out()
    try:
        if args.cmd == "browser":
            asyncio.run(keep_browser_open())
            return
        runner = {"list": cmd_list, "get": cmd_get, "batch": cmd_batch}[args.cmd]
        asyncio.run(runner(args))
    except (TranscriptError, ValueError) as ex:
        sys.exit(sanitize_terminal(str(ex)))
    except KeyboardInterrupt:
        sys.exit(130)


if __name__ == "__main__":
    main()
