"""Runtime settings: where the browser is, which profile it uses, where state lives.

Every value can be set by an environment variable, and the command line can
override them again before any command runs.
"""
from __future__ import annotations

import os
import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path

PROG = "msteams-transcripts"
TEAMS_URL = "https://teams.cloud.microsoft/"

WINDOWS_BROWSERS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]
MACOS_BROWSERS = [
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
]
LINUX_BROWSERS = ["microsoft-edge", "microsoft-edge-stable", "google-chrome", "chromium", "chromium-browser"]


def _default_profile_dir() -> Path:
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share")
    return base / "TeamsTranscriptProfile"


def browser_candidates(explicit: str = "") -> list[str]:
    """Browser executables to try, most preferred first."""
    found = [explicit] if explicit else []
    if sys.platform == "win32":
        found += WINDOWS_BROWSERS
    elif sys.platform == "darwin":
        found += MACOS_BROWSERS
    else:
        found += [p for name in LINUX_BROWSERS if (p := shutil.which(name))]
    return [p for p in found if p]


@dataclass
class Settings:
    cdp_port: int = field(default_factory=lambda: int(os.environ.get("TT_CDP_PORT", "9222")))
    profile_dir: Path = field(
        default_factory=lambda: Path(os.environ["TT_PROFILE_DIR"]).expanduser()
        if os.environ.get("TT_PROFILE_DIR")
        else _default_profile_dir()
    )
    browser_exe: str = field(default_factory=lambda: os.environ.get("TT_BROWSER_EXE", ""))
    state_dir: Path = field(
        default_factory=lambda: Path(os.environ["TT_STATE_DIR"]).expanduser()
        if os.environ.get("TT_STATE_DIR")
        else Path.home() / ".teams_transcripts"
    )

    @property
    def cdp_url(self) -> str:
        return f"http://127.0.0.1:{self.cdp_port}"

    @property
    def last_list(self) -> Path:
        """Where the rows printed by the last `list` run are cached."""
        return self.state_dir / "last_list.json"


SETTINGS = Settings()


def configure(port: int | None = None, profile: str = "", browser: str = "") -> Settings:
    """Apply command-line overrides to the process-wide settings."""
    if port:
        SETTINGS.cdp_port = port
    if profile:
        SETTINGS.profile_dir = Path(profile).expanduser()
    if browser:
        SETTINGS.browser_exe = browser
    return SETTINGS
