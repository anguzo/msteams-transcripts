---
name: teams-transcripts
description: List and download Microsoft Teams meeting transcripts (with optional attendees, speakers, shared files) through the user's signed-in browser session, using `uvx msteams-transcripts`. Use when asked to fetch, save, summarise, or search Teams meeting transcripts or recaps, or when a Teams recap link (teams.cloud.microsoft/l/meetingrecap) is given.
---

# Teams transcripts

`msteams-transcripts` downloads Teams meeting transcripts that the Teams user
interface will not let you download. It drives a dedicated browser profile over
the debugging protocol and replays the same calls the Recap tab makes, so it
only ever sees what the signed-in user can see.

## Preconditions

1. Decide how to invoke the tool, in this order of preference:
   - **`uvx msteams-transcripts ...`** whenever `uv` is available. This is the
     normal case. It needs no installation and always runs the current release.
   - `msteams-transcripts ...` if that command is already on PATH, or its alias
     `teams-transcripts`. Do not install it just to get the shorter form.
   - `python -m teams_transcripts ...` inside a checkout where the package is
     installed in the environment.

   Confirm the choice with `uvx msteams-transcripts --version` (or the variant
   you picked) before anything else. The first uvx run downloads dependencies
   and takes about a minute; later runs start in a second or two.

2. A dedicated browser must be running with the debugging port:
   `uvx msteams-transcripts browser`. This is idempotent and returns at once if
   the browser is already up. If Teams inside that window is not signed in, stop
   and ask the user to sign in there; you cannot do it for them.

## Commands

Every example below uses `uvx`. If you established a different invocation in
step 1, substitute it and leave the arguments unchanged.

```
uvx msteams-transcripts list  --from YYYY-MM-DD --to YYYY-MM-DD|today [--json]
uvx msteams-transcripts get   <row#|recap-url|19:meeting_...@thread.v2> [-d] [-f txt|json|vtt|all] [-o DIR]
uvx msteams-transcripts batch --from D --to D [-d] [-f FMT] [-o DIR] [--only 3,7,12]
```

- `list` prints one row per transcript; a recorded recurring meeting appears
  once per occurrence. With `--json` it prints an array of objects carrying
  `n`, `subject`, `start` (ISO, UTC), `organizer`, `threadId`, `driveId`,
  `driveItemId`, `transcriptId` and `host`. Prefer `--json` when picking rows
  programmatically. Rows are cached in `~/.teams_transcripts/last_list.json`,
  and the numbers used by `get N` always refer to the most recent `list`.
- `get` writes `<DIR>/<YYYY-MM-DD HHMM> - <subject>.txt`, with `./transcripts`
  as the default directory. `-f all` adds the raw `.json` transcript and a `.vtt`.
- `-d`/`--details` prepends organizer, location, invited attendees with response
  status, speakers with segment counts, files shared in the meeting chat and the
  invitation text, and writes a `.meta.json` sidecar holding the same data
  structured.
- A recap URL works without listing first and takes a few seconds. When the user
  pastes a link containing `driveId` and `driveItemId`, use `get "<url>"` directly.
- Progress and errors go to stderr; the `list` table and `--json` go to stdout.

## Output format to expect

```
# <subject>
# Date: 2026-09-16 09:02
# Source: <SharePoint recording URL>
# Organizer: Name <email>            (only with -d)
# Invited (13): ...                  (only with -d)
# Speakers in transcript (5): ...    (only with -d)

[00:00:03] Speaker Name:
    sentence
    the next sentence by the same speaker

[00:00:27] Other Speaker:
    ...
```

Timestamps are offsets from the start of the recording, not wall-clock time.
Speaker names come from Teams; meeting-room devices appear under the room name,
so a room can be the most talkative "speaker" in a hybrid meeting.

## Typical flows

- **"Summarise yesterday's meetings"**:
  `uvx msteams-transcripts list --from <yesterday> --to <yesterday> --json`, then
  `uvx msteams-transcripts batch --from ... --to ... -d -o <temp dir>`, then read
  the `.txt` files.
- **"Who attended X and what did they say?"**:
  `uvx msteams-transcripts list --from ... --to ...` over the date range, pick the
  row by subject, `uvx msteams-transcripts get N -d`, then read the header for
  invitees and speakers and the body for content.
- **"Save this recap link"**: `uvx msteams-transcripts get "<url>" -d -f all`.

## Failure modes and what to do

| Message | Meaning | Action |
|---|---|---|
| `The browser did not expose its debugging port` | a window already uses the profile without the port | ask the user to close it, then rerun the `browser` command |
| `Could not capture a Teams token` | the Teams tab is not signed in, or still loading | retry once after about 30 s; if it persists ask the user to sign in in that window |
| `Could not capture the meeting-content token` | the recap deep link did not resolve | retry; if it persists ask the user to open any meeting's Recap tab in that window, then rerun |
| `Downloading the transcript failed: HTTP 403` | the user has no access to that recording | report it; there is nothing to fix |
| `0 transcript(s)` | nothing recorded with a transcript in that range | widen the range, or confirm the meetings were transcribed |

Timing: the very first `uvx` call also downloads dependencies, which adds about
a minute once. The first command in a browser session takes 30 to 60 seconds
while tokens are captured; later ones take 5 to 25 seconds. A `batch` over a
month can take several minutes. Allow a tool timeout of five minutes or more.

## Constraints

- Transcripts are internal meeting content. Keep output in the user's workspace,
  never send it to an external service, and do not paste long excerpts back
  unless asked.
- Do not edit the browser profile directory or kill browser processes; the
  profile holds the user's session.
- Not available through this tool: the attendance report of who actually joined,
  collaborative Loop notes, and Copilot AI notes. Say so rather than guessing.
