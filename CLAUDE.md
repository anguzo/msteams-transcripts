# Project instructions

## Commits

- **Never mention Claude, Claude Code, or any AI assistant in commit messages,
  pull request descriptions, code comments, or documentation.** No
  `Co-Authored-By` trailer for an assistant, no "Generated with" footer, no
  attribution of any kind.
- Write commit subjects in the imperative mood, under 72 characters.
- Commit or push only when asked.

## Package

- Distribution name is `msteams-transcripts`; the import package is
  `teams_transcripts`. The name `teams-transcripts` was already taken on PyPI.
- The build backend is hatchling. **The version lives in one place only:**
  `__version__` in `src/teams_transcripts/__init__.py`, read by
  `[tool.hatch.version]`. Never hardcode a version anywhere else.
- Bump it with hatch rather than by hand:
  `hatch version patch` (or `minor`, `major`, or an explicit `1.2.3`).
- Releases are tag-driven. The tag must be `v` plus the version, for example
  `v0.2.0`; the publish workflow fails if the tag and `__version__` disagree.

## Layout

```
src/teams_transcripts/
  config.py    settings, browser discovery, paths
  browser.py   starting the dedicated browser, the debugging port
  model.py     pure functions: identifiers, dates, reference extraction
  formats.py   transcript and HTML to readable text
  session.py   the attached browser and every network call
  download.py  writing transcripts and the .meta.json sidecar
  cli.py       argument parsing and the four commands
```

Keep `model.py` and `formats.py` free of network and browser code. They hold
everything that the tests cover.

## Conventions

- Errors the user can act on raise `TranscriptError`; the command line prints
  the message without a traceback. Never call `sys.exit` inside library code.
- Progress messages go to stderr through `browser.log`, so that stdout stays
  parseable. The `list` table and `--json` output are the only things on stdout.
- Run `python -m pytest` before committing. Tests must not need a browser or
  the network.

## Data handling

Transcripts are meeting content belonging to the people in them. They are
ignored by git, and they must never be committed, published in an issue, or
sent to a third-party service.
