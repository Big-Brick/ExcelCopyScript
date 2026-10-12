"""Persist initial directories for the two Excel open-file dialogs."""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path
from typing import Literal

FileKind = Literal["destination", "source"]
KINDS = ("destination", "source")


class RecentDirectories:
    """Remember the last successfully opened folder for each Excel file kind.

    An invalid or outdated preferences file never prevents launching the app.
    All writes are atomic; on failure the in-memory state is restored.
    """

    def __init__(self, path: Path | None = None):
        self.path = Path(path) if path is not None else (
            Path.home() / ".config" / "excel-copy-script" / "recent_directories.json"
        )
        self._directories: dict[str, str] = {}
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            if isinstance(data, dict) and data.get("version") == 1:
                stored = data.get("directories")
                if isinstance(stored, dict):
                    self._directories = {
                        kind: value
                        for kind, value in stored.items()
                        if kind in KINDS and isinstance(value, str) and value
                    }
        except (OSError, ValueError):
            # Optional UI preferences can be regenerated on the next successful selection.
            pass

    @staticmethod
    def _check_kind(kind: FileKind) -> None:
        if kind not in KINDS:
            raise ValueError(f"Невідомий тип файлу: {kind}")

    def initial_directory(self, kind: FileKind, fallback: Path | None = None) -> Path:
        """Use the last folder, then a caller-provided folder, then the home folder."""
        self._check_kind(kind)
        candidates = (self._directories.get(kind), fallback, Path.home())
        for candidate in candidates:
            if candidate is not None:
                directory = Path(candidate).expanduser()
                if directory.is_dir():
                    return directory
        # A missing home directory is unusual; let the file dialog decide.
        return Path.home()

    def remember_file(self, kind: FileKind, filename: str | Path) -> None:
        """Save the parent directory of a successfully opened file."""
        self._check_kind(kind)
        directory = Path(filename).expanduser().resolve().parent
        if not directory.is_dir():
            raise ValueError(f"Тека не існує: {directory}")
        value = str(directory)
        previous = self._directories.get(kind)
        if previous == value:
            return
        self._directories[kind] = value
        try:
            self._save()
        except Exception:
            if previous is None:
                del self._directories[kind]
            else:
                self._directories[kind] = previous
            raise

    def _save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="w", encoding="utf-8", dir=self.path.parent,
                prefix=".recent-directories-", suffix=".tmp", delete=False,
            ) as stream:
                tmp_path = Path(stream.name)
                json.dump(
                    {"version": 1, "directories": self._directories},
                    stream, ensure_ascii=False, indent=2,
                )
                stream.write("\n")
            os.replace(tmp_path, self.path)
        finally:
            if tmp_path is not None:
                tmp_path.unlink(missing_ok=True)
