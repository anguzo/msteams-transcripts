"""Starting and finding the dedicated browser instance.

Recent Edge and Chrome refuse remote debugging on the user's normal profile, so
this always uses a separate profile directory. The user signs in there once;
corporate single sign-on usually does it without a prompt.
"""
from __future__ import annotations

import os
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

from .config import SETTINGS, TEAMS_URL, browser_candidates
from .errors import TranscriptError


def log(msg: str) -> None:
    """Progress goes to stderr so that stdout stays machine-readable."""
    print(msg, file=sys.stderr, flush=True)


def cdp_alive() -> bool:
    try:
        with urllib.request.urlopen(f"{SETTINGS.cdp_url}/json/version", timeout=2) as r:
            return r.status == 200
    except Exception:
        return False


def find_browser() -> str:
    exe = next((p for p in browser_candidates(SETTINGS.browser_exe) if Path(p).exists()), None)
    if not exe:
        raise TranscriptError(
            "No Edge or Chrome executable found. Set TT_BROWSER_EXE or pass --browser with the full path."
        )
    return exe


def ensure_browser(wait_seconds: int = 30) -> None:
    """Start the dedicated browser if it is not already listening on the debug port."""
    if cdp_alive():
        return
    exe = find_browser()
    SETTINGS.profile_dir.mkdir(parents=True, exist_ok=True)
    log(f"Starting {Path(exe).name} with profile {SETTINGS.profile_dir} on port {SETTINGS.cdp_port} ...")
    kwargs: dict = {"stdout": subprocess.DEVNULL, "stderr": subprocess.DEVNULL, "close_fds": True}
    if os.name == "nt":
        kwargs["creationflags"] = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
    else:
        kwargs["start_new_session"] = True
    subprocess.Popen(
        [
            exe,
            f"--user-data-dir={SETTINGS.profile_dir}",
            f"--remote-debugging-port={SETTINGS.cdp_port}",
            "--no-first-run",
            "--no-default-browser-check",
            TEAMS_URL,
        ],
        **kwargs,
    )
    for _ in range(wait_seconds):
        time.sleep(1)
        if cdp_alive():
            log("Browser is up. If Teams asks you to sign in, do it in that window.")
            return
    raise TranscriptError(
        "The browser did not expose its debugging port. "
        f"Close any window already using the profile at {SETTINGS.profile_dir}, then try again."
    )
