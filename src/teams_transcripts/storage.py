"""Small helpers for private directories and atomic, private file writes.

POSIX mode bits are enforced where the platform supports them. On Windows,
mode bits are best-effort only and do not provide ACL enforcement.
"""
from __future__ import annotations

import os
import tempfile
from pathlib import Path

from .errors import TranscriptError


def _absolute(path: Path) -> Path:
    path = Path(path).expanduser()
    return path if path.is_absolute() else Path.cwd() / path


def _reject_symlink_components(path: Path) -> Path:
    """Reject symlinks in every existing component of a lexical path."""
    path = _absolute(path)
    current = Path(path.anchor)
    for component in path.parts[1:]:
        current /= component
        if current.is_symlink():
            raise TranscriptError(f"Refusing to use a symlinked path component: {current}")
    return path


def _chmod_private(path: Path) -> None:
    try:
        os.chmod(path, 0o700)
    except OSError:
        if os.name != "nt":
            raise


def _create_missing_dirs(path: Path, *, private: bool) -> None:
    """Create missing components without changing existing directory modes."""
    path = _reject_symlink_components(path)
    missing: list[Path] = []
    current = path
    while not current.exists():
        missing.append(current)
        parent = current.parent
        if parent == current:
            break
        current = parent
    path.mkdir(parents=True, exist_ok=True, mode=0o700)
    _reject_symlink_components(path)
    if private:
        for created in missing:
            _chmod_private(created)


def ensure_private_dir(path: Path) -> Path:
    """Create an application-owned directory with private POSIX mode bits."""
    path = Path(path).expanduser()
    absolute = _absolute(path)
    if absolute.parent == absolute:
        raise TranscriptError("The private directory must not be the filesystem root.")
    _reject_symlink_components(absolute)
    try:
        _create_missing_dirs(absolute, private=True)
        if not absolute.is_dir():
            raise TranscriptError(f"The private path is not a directory: {path}")
        _chmod_private(absolute)
    except OSError as exc:
        raise TranscriptError(f"Could not create the private directory: {path}") from exc
    return path


def atomic_write_text(path: Path, text: str) -> None:
    """Replace *path* atomically with a private text file.

    POSIX files are owner-only; Windows mode bits are best-effort and do not
    enforce ACLs.

    ``os.replace`` replaces a destination symlink rather than following it, so
    an existing link cannot redirect the write to another file. The temporary
    file is created in the destination directory to keep the replacement on
    one filesystem.
    """
    path = _absolute(Path(path))
    try:
        # The destination itself may be a symlink because os.replace replaces
        # it. Its parent and every ancestor must not be a symlink.
        _create_missing_dirs(path.parent, private=False)
        fd, temporary = tempfile.mkstemp(prefix=".msteams-transcripts-", suffix=".tmp", dir=path.parent)
        temporary_path = Path(temporary)
        try:
            try:
                os.fchmod(fd, 0o600)
            except AttributeError:
                if os.name != "nt":
                    raise
            except OSError:
                if os.name != "nt":
                    raise
            with os.fdopen(fd, "w", encoding="utf-8", newline="") as handle:
                fd = -1
                handle.write(text)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary_path, path)
            try:
                directory_fd = os.open(path.parent, os.O_RDONLY)
            except (AttributeError, OSError):
                pass
            else:
                try:
                    os.fsync(directory_fd)
                except OSError:
                    pass
                finally:
                    os.close(directory_fd)
        finally:
            if fd != -1:
                os.close(fd)
            try:
                temporary_path.unlink()
            except FileNotFoundError:
                pass
    except OSError as exc:
        raise TranscriptError(f"Could not write {path}") from exc


def read_text_without_symlink(path: Path, encoding: str = "utf-8") -> str:
    """Read a file without following a destination symlink where supported."""
    path = _reject_symlink_components(Path(path))
    flags = os.O_RDONLY
    nofollow = getattr(os, "O_NOFOLLOW", 0)
    try:
        fd = os.open(path, flags | nofollow)
    except OSError as exc:
        raise TranscriptError(f"Could not read {path}") from exc
    try:
        with os.fdopen(fd, "r", encoding=encoding) as handle:
            fd = -1
            return handle.read()
    finally:
        if fd != -1:
            os.close(fd)
