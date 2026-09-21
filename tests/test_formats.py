from __future__ import annotations

from teams_transcripts.formats import (
    details_header,
    fmt_offset,
    html_to_text,
    invite_description,
    speakers_of,
    transcript_to_text,
)

DOC = {
    "entries": [
        {"text": "Good morning.", "speakerDisplayName": "Ada Lovelace", "startOffset": "00:00:03.2936180"},
        {"text": "Let us begin.", "speakerDisplayName": "Ada Lovelace", "startOffset": "00:00:06.7510000"},
        {"text": "Agreed.", "speakerDisplayName": "Alan Turing", "startOffset": "00:01:02.3450000"},
        {"text": "", "speakerDisplayName": "Alan Turing", "startOffset": "00:01:05.0000000"},
        {"text": "One more thing.", "speakerDisplayName": "Ada Lovelace", "startOffset": "00:02:00.0000000"},
    ]
}


class TestOffsets:
    def test_fraction_is_dropped(self):
        assert fmt_offset("00:01:02.3450000") == "00:01:02"

    def test_missing_offset(self):
        assert fmt_offset(None) == "--:--:--"


class TestTranscriptText:
    def test_header_and_first_speaker(self):
        out = transcript_to_text(DOC, "Weekly sync", "2026-09-03 14:14", "https://example/rec.mp4")
        assert out.startswith("# Weekly sync\n# Date: 2026-09-03 14:14\n# Source: https://example/rec.mp4")
        assert "[00:00:03] Ada Lovelace:" in out

    def test_consecutive_lines_share_one_speaker_heading(self):
        out = transcript_to_text(DOC, "t", "d", "s")
        assert out.count("Ada Lovelace:") == 2  # she speaks, Alan interrupts, she resumes
        assert out.count("Alan Turing:") == 1

    def test_empty_segments_are_skipped(self):
        assert transcript_to_text(DOC, "t", "d", "s").count("    ") == 4

    def test_transcript_without_entries(self):
        assert transcript_to_text({}, "t", "d", "s").endswith("# Source: s\n")


class TestSpeakers:
    def test_counted_and_sorted_by_volume(self):
        assert speakers_of(DOC) == [("Ada Lovelace", 3), ("Alan Turing", 2)]

    def test_falls_back_to_the_identifier(self):
        assert speakers_of({"entries": [{"speakerId": "x@y", "text": "hi"}]}) == [("x@y", 1)]


class TestHtml:
    def test_style_blocks_are_dropped(self):
        html = "<html><head><style>@font-face {font-family:Wingdings}</style></head><body><p>Agenda</p></body></html>"
        assert html_to_text(html) == "Agenda"

    def test_link_targets_are_kept(self):
        out = html_to_text('<p>See <a href="https://example.com/notes">shared notes</a></p>')
        assert "https://example.com/notes" in out

    def test_join_links_are_not_kept(self):
        out = html_to_text('<a href="https://teams.microsoft.com/l/meetup-join/19%3ameeting_x">Join</a>')
        assert "meetup-join" not in out

    def test_boilerplate_is_cut_from_the_invitation(self):
        html = "<p>Agenda: roadmap</p><p>____________</p><p>Microsoft Teams meeting</p><p>Join: x</p>"
        assert invite_description(html) == "Agenda: roadmap"

    def test_invitation_without_boilerplate(self):
        assert invite_description("<p>Just an agenda</p>") == "Just an agenda"


class TestDetailsHeader:
    DETAILS = {
        "event": {
            "organizerName": "Ada Lovelace",
            "organizerAddress": "ada@example.com",
            "location": "Microsoft Teams Meeting",
            "attendees": [
                {
                    "name": "Ada Lovelace",
                    "address": "ada@example.com",
                    "type": "Organizer",
                    "status": {"response": "None"},
                },
                {
                    "name": "Alan Turing",
                    "address": "alan@example.com",
                    "type": "Required",
                    "status": {"response": "Accepted"},
                },
                {
                    "name": "Grace Hopper",
                    "address": "grace@example.com",
                    "type": "Optional",
                    "status": {"response": "NotResponded"},
                },
            ],
            "bodyContent": "<p>Agenda: roadmap</p><p>____</p><p>Microsoft Teams meeting</p>",
        },
        "files": [{"title": "Plan.xlsx", "url": "https://example/Plan.xlsx"}],
    }

    def test_every_section_is_present(self):
        lines = "\n".join(details_header(self.DETAILS, DOC))
        assert "# Organizer: Ada Lovelace <ada@example.com>" in lines
        assert "# Location: Microsoft Teams Meeting" in lines
        assert "# Invited (3):" in lines
        assert "# Speakers in transcript (2):" in lines
        assert "# Files shared in meeting chat (1):" in lines
        assert "#   Agenda: roadmap" in lines

    def test_response_status_is_shown_only_when_answered(self):
        lines = "\n".join(details_header(self.DETAILS, DOC))
        assert "Alan Turing <alan@example.com> (Required, Accepted)" in lines
        assert "Grace Hopper <grace@example.com> (Optional)" in lines

    def test_header_is_placed_above_the_body(self):
        out = transcript_to_text(DOC, "Weekly sync", "2026-09-03 14:14", "s", self.DETAILS)
        assert out.index("# Invited (3):") < out.index("[00:00:03] Ada Lovelace:")

    def test_empty_details_add_nothing(self):
        assert details_header({"event": {}}, {}) == []
