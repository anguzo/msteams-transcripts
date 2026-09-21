from __future__ import annotations

import asyncio
import os
import stat
from pathlib import Path

import pytest

from teams_transcripts import model
from teams_transcripts.browser import keep_browser_open, launch_context, log
from teams_transcripts.cli import build_parser
from teams_transcripts.config import SETTINGS, configure
from teams_transcripts.errors import TranscriptError
from teams_transcripts.session import FETCH_JS, TeamsSession
from teams_transcripts.storage import atomic_write_text, ensure_private_dir, read_text_without_symlink


def run(coro):
    return asyncio.run(coro)


def assert_posix_mode(path: Path, mode: int) -> None:
    # Windows mode bits are best-effort and do not assert ACL enforcement.
    if os.name != "nt":
        assert stat.S_IMODE(path.stat().st_mode) == mode


def symlink_or_skip(link: Path, target: Path, *, target_is_directory: bool = False) -> None:
    try:
        link.symlink_to(target, target_is_directory=target_is_directory)
    except OSError as exc:
        pytest.skip(f"symlink support is unavailable: {exc}")


class FakePage:
    def __init__(self, url: str = "about:blank"):
        self.url = url
        self.goto_calls = []

    async def goto(self, url, **kwargs):
        self.goto_calls.append((url, kwargs))
        self.url = url


class FakeContext:
    def __init__(self):
        self.pages = []
        self.closed = False
        self.listeners = {}

    async def new_page(self):
        page = FakePage()
        self.pages.append(page)
        return page

    def on(self, name, callback):
        self.listeners[name] = callback

    async def close(self):
        self.closed = True


class FakeChromium:
    def __init__(self, context):
        self.context = context
        self.calls = []

    async def launch_persistent_context(self, **kwargs):
        self.calls.append(kwargs)
        return self.context


class FakePlaywright:
    def __init__(self, context):
        self.chromium = FakeChromium(context)


def test_launch_context_uses_playwright_persistent_profile_without_cdp(tmp_path, monkeypatch):
    profile = tmp_path / "profile"
    executable = tmp_path / "msedge"
    executable.touch()
    context = FakeContext()
    playwright = FakePlaywright(context)
    monkeypatch.setattr("teams_transcripts.browser.SETTINGS.profile_dir", profile)
    monkeypatch.setattr("teams_transcripts.browser.SETTINGS.browser_exe", str(executable))

    run(launch_context(playwright))

    call = playwright.chromium.calls[0]
    assert call["user_data_dir"] == str(profile)
    assert call["executable_path"] == str(executable)
    assert call["headless"] is False
    assert all("debug" not in arg.lower() and "port" not in arg.lower() for arg in call["args"])
    assert context.pages[0].goto_calls[0][0] == "https://teams.cloud.microsoft/"
    assert_posix_mode(profile, 0o700)


def test_session_closes_context_and_playwright(monkeypatch):
    context = FakeContext()

    class FakePw:
        def __init__(self):
            self.stopped = False

        async def stop(self):
            self.stopped = True

    pw = FakePw()
    starter = type("Starter", (), {"start": lambda self: None})()

    async def start():
        return pw

    starter.start = start
    monkeypatch.setattr("teams_transcripts.session.async_playwright", lambda: starter)

    async def fake_launch(_pw):
        assert _pw is pw
        return context

    monkeypatch.setattr("teams_transcripts.session.launch_context", fake_launch)

    async def exercise():
        async with TeamsSession():
            pass

    run(exercise())
    assert context.closed is True
    assert pw.stopped is True


def test_interactive_browser_waits_for_context_close(monkeypatch):
    context = FakeContext()

    class FakePw:
        def __init__(self):
            self.stopped = False

        async def stop(self):
            self.stopped = True

    pw = FakePw()
    starter = type("Starter", (), {})()

    async def start():
        return pw

    starter.start = start

    def on(name, callback):
        context.listeners[name] = callback
        asyncio.get_running_loop().call_soon(callback, context)

    context.on = on

    async def fake_launch(_pw):
        return context

    monkeypatch.setattr("teams_transcripts.browser.async_playwright", lambda: starter)
    monkeypatch.setattr("teams_transcripts.browser.launch_context", fake_launch)

    run(keep_browser_open())
    assert context.closed is True
    assert pw.stopped is True


@pytest.mark.parametrize("failure", ["navigation", "origin"])
def test_launch_context_closes_context_when_setup_fails(tmp_path, monkeypatch, failure):
    class SetupPage(FakePage):
        async def goto(self, url, **kwargs):
            if failure == "navigation":
                raise RuntimeError("browser failure with secret token")
            self.url = "https://evil.example/"

    context = FakeContext()
    context.pages = [SetupPage()]
    executable = tmp_path / "msedge"
    executable.touch()
    playwright = FakePlaywright(context)
    monkeypatch.setattr("teams_transcripts.browser.SETTINGS.profile_dir", tmp_path / "profile")
    monkeypatch.setattr("teams_transcripts.browser.SETTINGS.browser_exe", str(executable))

    with pytest.raises(TranscriptError):
        run(launch_context(playwright))
    assert context.closed is True


@pytest.mark.parametrize(
    "url",
    [
        "http://teams.cloud.microsoft/l/meetingrecap?driveId=d&driveItemId=i",
        "https://teams.cloud.microsoft.evil.example/l/meetingrecap?driveId=d&driveItemId=i",
        "https://user:pass@teams.cloud.microsoft/l/meetingrecap?driveId=d&driveItemId=i",
        "https://teams.cloud.microsoft:443/l/meetingrecap?driveId=d&driveItemId=i",
        "https://teams.cloud.microsoft:bad/l/meetingrecap?driveId=d&driveItemId=i",
    ],
)
def test_recap_url_requires_approved_https_origin(url):
    with pytest.raises(model.UrlValidationError):
        model.refs_from_recap_url(url)


def test_deprecated_port_is_accepted_but_does_not_configure_cdp():
    before = SETTINGS.__dict__.copy()
    configure(port=9222)
    assert SETTINGS.__dict__ == before

    before_position = build_parser().parse_args(["--port", "9222", "list", "--from", "2026-09-01", "--to", "today"])
    after_position = build_parser().parse_args(["list", "--from", "2026-09-01", "--to", "today", "--port", "9222"])
    assert before_position.port == after_position.port == 9222


def test_sharepoint_host_override_is_retained_as_a_validated_option():
    args = build_parser().parse_args(
        [
            "get",
            "19:meeting_x@thread.v2",
            "--sharepoint-host",
            "contoso-my.sharepoint.com",
        ]
    )
    assert args.sharepoint_host == "contoso-my.sharepoint.com"
    with pytest.raises(SystemExit):
        build_parser().parse_args(["get", "19:meeting_x@thread.v2", "--sharepoint-host", "localhost"])


def test_sparse_recap_url_can_use_sharepoint_host_override():
    ref = model.refs_from_recap_url("https://teams.cloud.microsoft/l/meetingrecap?driveId=drive&driveItemId=item")

    assert ref is not None
    assert ref["host"] == ""


@pytest.mark.parametrize(
    "host",
    [
        "contoso.sharepoint.com",
        "contoso-my.sharepoint.com",
    ],
)
def test_sharepoint_tenant_suffixes_are_accepted(host):
    assert model.validate_sharepoint_host(host) == host


@pytest.mark.parametrize("host", ["contoso.sharepoint.us", "contoso.sharepoint.de", "contoso.sharepoint.cn"])
def test_unsupported_sovereign_sharepoint_hosts_are_rejected(host):
    with pytest.raises(model.UrlValidationError):
        model.validate_sharepoint_host(host)


@pytest.mark.parametrize(
    "host",
    [
        "localhost",
        "127.0.0.1",
        "[::1]",
        "contoso.sharepoint.com:443",
        "contoso.sharepoint.com:",
        "user:pass@contoso.sharepoint.com",
        "contoso.evil.example",
        "contoso.sharepoint.com/path",
    ],
)
def test_sharepoint_host_rejects_local_or_malformed_values(host):
    with pytest.raises(model.UrlValidationError):
        model.validate_sharepoint_host(host)


def test_fetch_refuses_a_redirected_page_before_evaluate():
    session = TeamsSession()
    evaluated = False

    class Page:
        url = "https://evil.example/"

        async def evaluate(self, *_args):
            nonlocal evaluated
            evaluated = True

    async def exercise():
        with pytest.raises(TranscriptError):
            await session._fetch(Page(), "https://contoso.sharepoint.com/_api/v2.1/me", {})

    run(exercise())
    assert evaluated is False


def test_fetch_accepts_and_checks_final_response_url():
    assert "url: r.url" in FETCH_JS
    session = TeamsSession()

    class Page:
        url = "https://contoso.sharepoint.com/"

        async def evaluate(self, *_args):
            return {
                "status": 200,
                "ct": "application/json",
                "url": "https://contoso.sharepoint.com/_api/v2.1/me",
                "text": "{}",
            }

    result = run(session._fetch(Page(), "https://contoso.sharepoint.com/_api/v2.1/me", {}))
    assert result["url"].startswith("https://contoso.sharepoint.com/")


@pytest.mark.parametrize(
    "response",
    [
        {"status": 200, "ct": "text/html", "url": "https://evil.example/login", "text": "secret"},
        {"status": 200, "ct": "text/html", "text": "secret"},
    ],
)
def test_fetch_rejects_unexpected_or_missing_final_response_url(response):
    session = TeamsSession()

    class Page:
        url = "https://contoso.sharepoint.com/"

        async def evaluate(self, *_args):
            return response

    with pytest.raises(TranscriptError):
        run(session._fetch(Page(), "https://contoso.sharepoint.com/_api/v2.1/me", {}))


def test_fetch_wraps_playwright_evaluation_failures_without_secrets():
    session = TeamsSession()

    class Page:
        url = "https://contoso.sharepoint.com/"

        async def evaluate(self, *_args):
            raise RuntimeError("Bearer very-secret-token")

    with pytest.raises(TranscriptError) as error:
        run(session._fetch(Page(), "https://contoso.sharepoint.com/_api/v2.1/me", {}))
    assert "very-secret-token" not in str(error.value)
    assert "contoso.sharepoint.com" in str(error.value)


def test_session_navigation_failure_is_a_transcript_error_without_secrets():
    class FailingPage(FakePage):
        async def goto(self, _url, **_kwargs):
            raise RuntimeError("Bearer very-secret-token")

        async def close(self):
            self.closed = True

    page = FailingPage()

    class Context:
        pages = []

        async def new_page(self):
            return page

    session = TeamsSession()
    session.ctx = Context()

    with pytest.raises(TranscriptError) as error:
        run(session.teams_page())
    assert "very-secret-token" not in str(error.value)
    assert "teams.cloud.microsoft" in str(error.value)
    assert page.closed is True


def test_session_navigation_rejects_unexpected_response_origin():
    class Response:
        url = "https://evil.example/redirected"

    class ResponsePage(FakePage):
        async def goto(self, url, **_kwargs):
            self.url = url
            return Response()

        async def close(self):
            self.closed = True

    page = ResponsePage()

    class Context:
        pages = []

        async def new_page(self):
            return page

    session = TeamsSession()
    session.ctx = Context()

    with pytest.raises(TranscriptError):
        run(session.teams_page())
    assert page.closed is True


def test_sharepoint_page_rejects_an_unexpected_redirect():
    class RedirectPage(FakePage):
        def __init__(self):
            super().__init__()
            self.closed = False

        async def goto(self, _url, **_kwargs):
            self.url = "https://evil.example/"

        async def close(self):
            self.closed = True

    page = RedirectPage()

    class Context:
        async def new_page(self):
            return page

    session = TeamsSession()
    session.ctx = Context()

    with pytest.raises(TranscriptError):
        run(session.sharepoint_page("contoso.sharepoint.com"))
    assert page.closed is True


def test_path_components_are_encoded_before_sharepoint_requests():
    session = TeamsSession()
    page = FakePage("https://contoso.sharepoint.com/")
    urls = []

    async def sharepoint_page(_host):
        return page

    async def fetch(_page, url, _headers):
        urls.append(url)
        return {"status": 200, "text": '{"value": []}'}

    session.sharepoint_page = sharepoint_page
    session._fetch = fetch

    run(session.list_transcripts("contoso.sharepoint.com", "drive/one", "item?two"))
    run(session.fetch_transcript("contoso.sharepoint.com", "drive/one", "item?two", "transcript/three", "vtt"))

    assert "/drives/drive%2Fone/items/item%3Ftwo/" in urls[0]
    assert "/drives/drive%2Fone/items/item%3Ftwo/media/transcripts/transcript%2Fthree/" in urls[1]
    assert "format=vtt" in urls[1]


def test_atomic_write_replaces_symlink_without_touching_target(tmp_path):
    target = tmp_path / "target.txt"
    destination = tmp_path / "result.txt"
    target.write_text("original", encoding="utf-8")
    symlink_or_skip(destination, target)

    atomic_write_text(destination, "replacement")

    assert destination.is_symlink() is False
    assert destination.read_text(encoding="utf-8") == "replacement"
    assert target.read_text(encoding="utf-8") == "original"
    assert_posix_mode(destination, 0o600)


def test_private_directory_is_owner_only_but_existing_output_directory_is_not_chmodded(tmp_path):
    private = ensure_private_dir(tmp_path / "state")
    assert_posix_mode(private, 0o700)

    output = tmp_path / "output"
    output.mkdir(mode=0o755)
    os.chmod(output, 0o755)
    atomic_write_text(output / "file.txt", "data")
    assert_posix_mode(output, 0o755)


def test_private_directory_rejects_symlinked_parent(tmp_path):
    real_parent = tmp_path / "real"
    real_parent.mkdir()
    link_parent = tmp_path / "link"
    symlink_or_skip(link_parent, real_parent, target_is_directory=True)

    with pytest.raises(TranscriptError):
        ensure_private_dir(link_parent / "state")


def test_atomic_write_rejects_symlinked_parent_without_following_it(tmp_path):
    real_parent = tmp_path / "real"
    real_parent.mkdir()
    link_parent = tmp_path / "link"
    symlink_or_skip(link_parent, real_parent, target_is_directory=True)

    with pytest.raises(TranscriptError):
        atomic_write_text(link_parent / "output.txt", "data")
    assert not (real_parent / "output.txt").exists()


def test_cache_reads_reject_destination_symlinks(tmp_path):
    target = tmp_path / "target.json"
    cache = tmp_path / "last_list.json"
    target.write_text("[]", encoding="utf-8")
    symlink_or_skip(cache, target)

    with pytest.raises(TranscriptError):
        read_text_without_symlink(cache)


def test_control_characters_are_removed_from_names_and_progress(capsys):
    value = "subject\x1b[31m\nnext\x7f"
    assert "\x1b" not in model.safe_name(value)
    assert "\n" not in model.safe_name(value)

    log(value)
    output = capsys.readouterr().err
    assert "\x1b" not in output
    assert "\n" not in output.rstrip("\n")
