"""Launching the dedicated browser through Playwright.

The dedicated profile is opened by Playwright itself. No remote-debugging
listener is exposed, and callers own the lifetime of the returned context.
"""
from __future__ import annotations

import asyncio
import sys
import urllib.parse
from pathlib import Path

from .config import SETTINGS, TEAMS_URL, browser_candidates
from .errors import TranscriptError
from .model import is_approved_auth_url, sanitize_terminal, validate_teams_url
from .storage import ensure_private_dir

try:
    from playwright.async_api import async_playwright
except ImportError as exc:  # pragma: no cover
    raise TranscriptError("Playwright is missing. Install this tool with 'pip install msteams-transcripts'.") from exc


def log(msg: str) -> None:
    """Progress goes to stderr so that stdout stays machine-readable."""
    print(sanitize_terminal(msg), file=sys.stderr, flush=True)


def find_browser() -> str:
    exe = next((p for p in browser_candidates(SETTINGS.browser_exe) if Path(p).is_file()), None)
    if not exe:
        raise TranscriptError(
            "No Edge or Chrome executable found. Set TT_BROWSER_EXE or pass --browser with the full path."
        )
    return exe


def _allow_teams_or_auth_origin(url: str) -> None:
    try:
        validate_teams_url(url)
        return
    except ValueError as exc:
        if is_approved_auth_url(url):
            return
        raise TranscriptError("The browser navigated to an unexpected Teams origin.") from exc


def _safe_url_label(url: str) -> str:
    try:
        parsed = urllib.parse.urlsplit(url)
        return sanitize_terminal(f"{parsed.hostname or 'unknown-host'}{parsed.path or '/'}")[:160]
    except (TypeError, ValueError):
        return "unknown-url"


async def _goto(page, url: str, validate_final) -> None:
    try:
        response = await page.goto(url, wait_until="domcontentloaded")
    except Exception as exc:
        raise TranscriptError(
            f"Browser navigation failed for {_safe_url_label(url)} ({type(exc).__name__})."
        ) from exc
    try:
        response_url = getattr(response, "url", "")
        if response_url:
            validate_final(response_url)
        validate_final(page.url)
    except Exception as exc:
        raise TranscriptError(
            f"Browser navigation ended at an unexpected origin for {_safe_url_label(url)}."
        ) from exc


async def launch_context(playwright, open_teams: bool = True):
    """Launch one owner-managed persistent context using the dedicated profile."""
    profile = ensure_private_dir(SETTINGS.profile_dir)
    exe = find_browser()
    log(f"Starting {Path(exe).name} with profile {profile} ...")
    try:
        context = await playwright.chromium.launch_persistent_context(
            user_data_dir=str(profile),
            executable_path=exe,
            headless=False,
            args=["--no-first-run", "--no-default-browser-check"],
        )
    except Exception as exc:
        raise TranscriptError(
            f"Could not launch the dedicated browser with profile {profile}. Close any other window using it and retry."
        ) from exc

    try:
        if open_teams:
            page = context.pages[0] if context.pages else await context.new_page()
            if not page.url or page.url == "about:blank":
                await _goto(page, TEAMS_URL, _allow_teams_or_auth_origin)
            else:
                try:
                    _allow_teams_or_auth_origin(page.url)
                except TranscriptError:
                    await _goto(page, TEAMS_URL, _allow_teams_or_auth_origin)
            _allow_teams_or_auth_origin(page.url)
            log("Browser is up. If Teams asks you to sign in, do it in that window.")
    except BaseException:
        try:
            await context.close()
        except Exception:
            pass
        raise
    return context


async def keep_browser_open() -> None:
    """Run the interactive sign-in browser until it is closed or interrupted."""
    pw = await async_playwright().start()
    context = None
    try:
        context = await launch_context(pw)
        closed = asyncio.Event()

        def on_close(*_args) -> None:
            closed.set()

        context.on("close", on_close)
        log("Leave this browser open while signing in. Close it when you are finished.")
        await closed.wait()
    finally:
        if context is not None:
            try:
                await context.close()
            except Exception:
                pass
        await pw.stop()
