"""Turning the raw transcript document into readable text."""
from __future__ import annotations

import re
from html.parser import HTMLParser


def fmt_offset(off: str | None) -> str:
    """'00:01:02.3450000' to '00:01:02'."""
    if not off:
        return "--:--:--"
    return off.split(".")[0]


class _Text(HTMLParser):
    """Minimal HTML to text, keeping link targets and dropping style blocks."""

    SKIP = ("style", "script", "head", "title")

    def __init__(self):
        super().__init__()
        self.parts: list[str] = []
        self.skip = 0
        self.href = ""

    def handle_data(self, d):
        if not self.skip:
            self.parts.append(d)

    def handle_starttag(self, tag, attrs):
        if tag in self.SKIP:
            self.skip += 1
        elif tag in ("br", "p", "div", "li", "tr"):
            self.parts.append("\n")
        elif tag == "a":
            href = dict(attrs).get("href", "") or ""
            keep = href.startswith("http") and "/meetup-join/" not in href and "/meet/" not in href
            self.href = href if keep else ""

    def handle_endtag(self, tag):
        if tag in self.SKIP and self.skip:
            self.skip -= 1
        elif tag == "a" and self.href:
            self.parts.append(f" <{self.href}>")
            self.href = ""


def html_to_text(html: str) -> str:
    t = _Text()
    t.feed(html or "")
    lines = [re.sub(r"\s+", " ", line).strip() for line in "".join(t.parts).splitlines()]
    return "\n".join(line for line in lines if re.search(r"[A-Za-z0-9]", line))


def invite_description(body_html: str) -> str:
    """The invitation text, with the Teams join boilerplate cut off."""
    txt = html_to_text(body_html)
    marks = [i for i in (txt.find("Microsoft Teams meeting"), txt.find("____"), txt.find("Join the meeting now")) if i >= 0]
    return txt[: min(marks)].strip() if marks else txt.strip()


def speakers_of(doc: dict) -> list[tuple[str, int]]:
    """Who spoke, and how many segments each, most talkative first."""
    counts: dict[str, int] = {}
    for e in doc.get("entries", []):
        n = (e.get("speakerDisplayName") or e.get("speakerId") or "Unknown").strip()
        counts[n] = counts.get(n, 0) + 1
    return sorted(counts.items(), key=lambda x: -x[1])


def details_header(details: dict, doc: dict) -> list[str]:
    """Comment lines describing the meeting, placed above the transcript body."""
    ev = details.get("event") or {}
    out: list[str] = []
    if ev.get("organizerName") or ev.get("organizerAddress"):
        out.append(f"# Organizer: {ev.get('organizerName', '')} <{ev.get('organizerAddress', '')}>".replace(" <>", ""))
    if ev.get("location"):
        out.append(f"# Location: {ev['location']}")
    att = ev.get("attendees") or []
    if att:
        out.append(f"# Invited ({len(att)}):")
        for a in att:
            kind = a.get("type") or "Required"
            resp = (a.get("status") or {}).get("response") or ""
            resp = "" if resp in ("None", "NotResponded", "") else f", {resp}"
            out.append(f"#   {a.get('name', '')} <{a.get('address', '')}> ({kind}{resp})")
    sp = speakers_of(doc)
    if sp:
        out.append(f"# Speakers in transcript ({len(sp)}):")
        out.extend(f"#   {n} ({c} segments)" for n, c in sp)
    files = details.get("files") or []
    if files:
        out.append(f"# Files shared in meeting chat ({len(files)}):")
        out.extend(f"#   {f['title']}  {f['url']}" for f in files)
    is_html = (ev.get("bodyContentType") or "html").lower() == "html"
    desc = invite_description(ev.get("bodyContent", "")) if is_html else (ev.get("bodyContent") or "")
    if desc:
        out.append("# Invitation text:")
        out.extend(f"#   {line}" for line in desc.splitlines()[:40])
    return out


def transcript_to_text(doc: dict, title: str, when: str, source: str, details: dict | None = None) -> str:
    """Plain text: a comment header, then '[hh:mm:ss] Speaker:' with their lines."""
    lines = [f"# {title}", f"# Date: {when}", f"# Source: {source}"]
    if details:
        lines += details_header(details, doc)
    lines.append("")
    prev_speaker = None
    for e in doc.get("entries", []):
        speaker = (e.get("speakerDisplayName") or e.get("speakerId") or "Unknown").strip()
        text = (e.get("text") or "").strip()
        if not text:
            continue
        if speaker != prev_speaker:
            lines.append("")
            lines.append(f"[{fmt_offset(e.get('startOffset'))}] {speaker}:")
            prev_speaker = speaker
        lines.append(f"    {text}")
    return "\n".join(lines).strip() + "\n"
