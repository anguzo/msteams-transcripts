"""Exception types raised by the package."""

from __future__ import annotations


class TranscriptError(Exception):
    """Something went wrong that the user can act on.

    The command line catches this and prints the message without a traceback,
    so messages should read as advice, not as internal state.
    """
