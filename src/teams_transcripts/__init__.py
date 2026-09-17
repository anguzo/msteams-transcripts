"""List and download Microsoft Teams meeting transcripts through a signed-in browser.

The Teams web client already receives the whole transcript when you open a meeting
Recap tab; the Download button is only a UI policy. This package attaches to a
browser you are signed in to and replays the same requests, so it sees exactly what
you can see and nothing more.

Library use::

    import asyncio
    from teams_transcripts import TeamsSession, download_ref, refs_from_recap_url

    async def main():
        async with TeamsSession() as s:
            ref = refs_from_recap_url("https://teams.cloud.microsoft/l/meetingrecap?...")
            await download_ref(s, ref, Path("out"), "txt", details=True)

    asyncio.run(main())
"""
from __future__ import annotations

__version__ = "0.1.0"

from .errors import TranscriptError

__all__ = [
    "__version__",
    "TranscriptError",
    "TeamsSession",
    "download_ref",
    "refs_from_recap_url",
    "transcript_refs_from_contents",
    "transcript_to_text",
]

_LAZY = {
    "TeamsSession": ".session",
    "download_ref": ".download",
    "refs_from_recap_url": ".model",
    "transcript_refs_from_contents": ".model",
    "transcript_to_text": ".formats",
}


def __getattr__(name: str):
    """Import the browser-driving parts only when they are actually used."""
    module = _LAZY.get(name)
    if module is None:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    from importlib import import_module

    return getattr(import_module(module, __name__), name)


def __dir__() -> list[str]:
    return sorted(__all__)
