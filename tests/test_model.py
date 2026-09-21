from __future__ import annotations

from datetime import datetime, timezone

import pytest

from teams_transcripts.model import (
    MT_RE,
    parse_date,
    refs_from_recap_url,
    safe_name,
    shared_files_from_contents,
    thread_id_from,
    ticks_to_dt,
    transcript_refs_from_contents,
    web_recap_url,
)

THREAD = "19:meeting_ZDQwNGFmZmYtNTY1ZC00M2NhLWE2OTgtY2E4ZWMxMzgyZGEy@thread.v2"


def test_mt_api_base_uses_the_tenant_region():
    match = MT_RE.match("https://teams.cloud.microsoft/api/mt/emea/v2.1/me/calendars/calendarView")

    assert match is not None
    assert match.group(1) == "https://teams.cloud.microsoft/api/mt/emea"


class TestThreadId:
    def test_plain_identifier_is_unchanged(self):
        assert thread_id_from(THREAD) == THREAD

    def test_extracted_from_a_join_link(self):
        url = (
            "https://teams.microsoft.com/l/meetup-join/"
            "19%3ameeting_MTRkNmFiMzQtNjk2My00YzY4LTlmOTUtYjNjNTI2OTQyZGY0%40thread.v2/0?context=x"
        )
        assert thread_id_from(url) == "19:meeting_MTRkNmFiMzQtNjk2My00YzY4LTlmOTUtYjNjNTI2OTQyZGY0@thread.v2"

    def test_case_of_the_identifier_is_preserved(self):
        # The middle is base64 and the service rejects a lowercased form.
        assert "ZDQwNGFmZmYt" in thread_id_from(THREAD.replace("19:", "19%3a"))

    def test_text_without_an_identifier(self):
        assert thread_id_from("no identifier here") is None
        assert thread_id_from("") is None


class TestDates:
    def test_start_of_day(self):
        assert parse_date("2026-09-01") == datetime(2026, 9, 1, tzinfo=timezone.utc)

    def test_end_is_exclusive_next_midnight(self):
        assert parse_date("2026-09-01", end=True) == datetime(2026, 9, 2, tzinfo=timezone.utc)

    def test_today_keyword(self):
        assert parse_date("today").date() == datetime.now(timezone.utc).date()

    def test_rejects_other_formats(self):
        with pytest.raises(ValueError):
            parse_date("01/09/2026")

    def test_dotnet_ticks(self):
        # 2026-09-03T12:14:04Z expressed as .NET ticks
        assert ticks_to_dt(639240344442109056).year == 2026

    def test_bad_ticks_are_none(self):
        assert ticks_to_dt(0) is None
        assert ticks_to_dt("nonsense") is None


class TestSafeName:
    def test_strips_characters_windows_rejects(self):
        assert safe_name('a/b\\c:d*e?f"g<h>i|j') == "a b c d e f g h i j"

    def test_collapses_whitespace_and_trims(self):
        assert safe_name("  two   words .  ") == "two words"

    def test_never_returns_empty(self):
        assert safe_name("///") == "meeting"

    def test_respects_the_limit(self):
        assert len(safe_name("x" * 300)) == 80


class TestRecapUrl:
    URL = (
        "https://teams.cloud.microsoft/l/meetingrecap?driveId=b%21abq260&driveItemId=012D3K2L"
        "&sitePath=https%3A%2F%2Fcontoso-my.sharepoint.com%2Fpersonal%2Fa_b_c%2F_api%2Fv2.1%2Fdrives%2Fb%21abq260"
        "%2Fitems%2F012D3K2L%2Fversions%2Fcurrent%2Fmedia%2Ftranscripts%2F79262c7e-de76-42a2-ae2c-0fa871c38d17%2Fcontent"
        # Teams writes spaces in fileUrl as "+", which the query parser turns back into spaces.
        "&fileUrl=https%3A%2F%2Fcontoso-my.sharepoint.com%2Fpersonal%2Fa_b_c%2FDocuments%2FRecordings%2F"
        "Weekly+sync-20260903_141403-Meeting+Recording.mp4%3Fweb%3D1"
        "&iCalUid=040000008200E000&threadId=19%3Ameeting_ZDQw%40thread.v2"
    )

    def test_pulls_out_the_identifiers(self):
        ref = refs_from_recap_url(self.URL)
        assert ref["driveId"] == "b!abq260"
        assert ref["driveItemId"] == "012D3K2L"
        assert ref["transcriptId"] == "79262c7e-de76-42a2-ae2c-0fa871c38d17"
        assert ref["host"] == "contoso-my.sharepoint.com"
        assert ref["iCalUID"] == "040000008200E000"
        assert ref["threadId"] == "19:meeting_ZDQw@thread.v2"

    def test_subject_comes_from_the_recording_file_name(self):
        assert refs_from_recap_url(self.URL)["subject"] == "Weekly sync-20260903_141403-Meeting Recording"

    def test_link_without_drive_identifiers(self):
        assert refs_from_recap_url("https://teams.cloud.microsoft/l/meetingrecap?threadId=19%3Ameeting_x%40thread.v2") is None

    def test_web_form_skips_the_desktop_app_prompt(self):
        url = web_recap_url(THREAD, "040000", "tenant-id")
        assert url.startswith("https://teams.cloud.microsoft/_#/l/meetingrecap?")
        assert "tenantId=tenant-id" in url


CONTENTS = {
    "resources": [
        {"id": "a-AISummary", "type": "AISummary", "metadata": {"driveId": "d1"}},
        {
            "id": "a-TranscriptV2",
            "type": "TranscriptV2",
            "location": "https://contoso-my.sharepoint.com/personal/x/_api/v2.1/drives/d1/items/i1/versions/current/media/transcripts/t1/content",
            "metadata": {
                "driveId": "d1",
                "driveItemId": "i1",
                "transcriptId": "t1",
                "startTime": "639240344442109056",
                "endTime": "639240378409283601",
                "iCalUid": "040000008200E000",
                "threadId": THREAD,
                "fileTitle": "Weekly sync-20260903_141403-Meeting Recording.mp4",
                "webUrl": "https://contoso-my.sharepoint.com/personal/x/Documents/Recordings/Weekly%20sync.mp4",
            },
        },
        {
            "id": "a-chat",
            "type": "MeetingChat",
            "subType": "Excel",
            "location": "https://contoso.sharepoint.com/sites/team/Shared%20Documents/Plan.xlsx",
            "metadata": {"fileTitle": "Plan.xlsx", "fileType": "xlsx"},
        },
    ]
}


class TestContents:
    def test_only_transcript_resources_are_returned(self):
        refs = transcript_refs_from_contents(CONTENTS, "Weekly sync")
        assert len(refs) == 1
        assert refs[0]["subject"] == "Weekly sync"
        assert refs[0]["transcriptId"] == "t1"
        assert refs[0]["host"] == "contoso-my.sharepoint.com"
        assert refs[0]["start"].startswith("2026-09-03")

    def test_transcript_id_falls_back_to_the_location(self):
        contents = {"resources": [dict(CONTENTS["resources"][1], metadata={k: v for k, v in CONTENTS["resources"][1]["metadata"].items() if k != "transcriptId"})]}
        assert transcript_refs_from_contents(contents)[0]["transcriptId"] == "t1"

    def test_meeting_without_recap_content(self):
        assert transcript_refs_from_contents({"resources": []}) == []

    def test_chat_files_are_listed_separately(self):
        files = shared_files_from_contents(CONTENTS)
        assert len(files) == 1
        assert files[0]["title"] == "Plan.xlsx"
        assert files[0]["url"].endswith("Plan.xlsx")
