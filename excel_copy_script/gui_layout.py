"""Persist manually adjusted Treeview widths and Panedwindow sashes.

All keys are stable UI identifiers (not workbook-derived), so a saved layout is
reused for every destination file. Only explicit drag gestures update settings:
geometry changes in hidden notebook tabs must never overwrite user choices.
"""

from __future__ import annotations

import json
import os
import tempfile
import tkinter as tk
from pathlib import Path
from tkinter import ttk
from typing import Callable


class LayoutPreferences:
    VERSION = 1
    MIN_COLUMN_WIDTH = 40
    MAX_COLUMN_WIDTH = 3000
    MIN_SASH_POSITION = 60
    MAX_SASH_POSITION = 10000

    def __init__(self, path: Path | None = None) -> None:
        self.path = Path(path) if path is not None else (
            Path.home() / ".config" / "excel-copy-script" / "layout.json"
        )
        self.columns: dict[str, dict[str, int]] = {}
        self.sashes: dict[str, int] = {}
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
            if not isinstance(raw, dict) or raw.get("version") != self.VERSION:
                return
            if isinstance(raw.get("columns"), dict):
                for key, widths in raw["columns"].items():
                    if not isinstance(key, str) or not isinstance(widths, dict):
                        continue
                    valid = {
                        name: width for name, width in widths.items()
                        if isinstance(name, str) and self._valid_column(width)
                    }
                    if valid:
                        self.columns[key] = valid
            if isinstance(raw.get("sashes"), dict):
                self.sashes = {
                    name: position for name, position in raw["sashes"].items()
                    if isinstance(name, str) and self._valid_sash(position)
                }
        except (OSError, ValueError, TypeError):
            # Missing or corrupt preferences must not prevent the GUI starting.
            pass

    @classmethod
    def _valid_column(cls, value: object) -> bool:
        return type(value) is int and cls.MIN_COLUMN_WIDTH <= value <= cls.MAX_COLUMN_WIDTH

    @classmethod
    def _valid_sash(cls, value: object) -> bool:
        return type(value) is int and cls.MIN_SASH_POSITION <= value <= cls.MAX_SASH_POSITION

    def set_columns(self, name: str, widths: dict[str, int]) -> None:
        validated = {key: int(width) for key, width in widths.items()
                     if self._valid_column(width)}
        if not validated or validated == self.columns.get(name):
            return
        previous = self.columns.get(name)
        self.columns[name] = validated
        try:
            self._save()
        except OSError:
            if previous is None:
                del self.columns[name]
            else:
                self.columns[name] = previous
            raise

    def set_sash(self, name: str, position: int) -> None:
        if not self._valid_sash(position) or self.sashes.get(name) == position:
            return
        previous = self.sashes.get(name)
        self.sashes[name] = position
        try:
            self._save()
        except OSError:
            if previous is None:
                del self.sashes[name]
            else:
                self.sashes[name] = previous
            raise

    def _save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="w", encoding="utf-8", dir=self.path.parent,
                prefix=".layout-", suffix=".tmp", delete=False,
            ) as stream:
                tmp = Path(stream.name)
                json.dump({
                    "version": self.VERSION, "columns": self.columns,
                    "sashes": self.sashes,
                }, stream, ensure_ascii=False, indent=2)
                stream.write("\n")
            os.replace(tmp, self.path)
        finally:
            if tmp is not None:
                tmp.unlink(missing_ok=True)


class LayoutPersistence:
    """Bind user resizes to atomic saves and restore sashes only when visible."""

    def __init__(
        self, settings: LayoutPreferences | None = None,
        on_error: Callable[[OSError], None] | None = None,
    ) -> None:
        self.settings = settings or LayoutPreferences()
        self.on_error = on_error
        self._resizing: set[str] = set()
        self._sashes_by_tab: dict[str, list[tuple[str, ttk.Panedwindow]]] = {}
        self._restored_tabs: set[str] = set()

    def _save(self, operation: Callable[[], None]) -> None:
        try:
            operation()
        except OSError as exc:
            if self.on_error is not None:
                self.on_error(exc)

    def register_columns(self, key: str, tree: ttk.Treeview) -> None:
        """Restore widths; watch header separators, not routine row clicks."""
        for column, width in self.settings.columns.get(key, {}).items():
            if column in tree["columns"]:
                tree.column(column, width=width)

        def pressed(event) -> None:
            if tree.identify_region(event.x, event.y) == "separator":
                self._resizing.add(key)

        def released(_event) -> None:
            if key not in self._resizing:
                return
            self._resizing.discard(key)
            # ttk.Treeview's class binding applies the drag after this binding.
            tree.after_idle(lambda: self.capture_columns(key, tree))

        tree.bind("<ButtonPress-1>", pressed, add="+")
        tree.bind("<ButtonRelease-1>", released, add="+")

    def capture_columns(self, key: str, tree: ttk.Treeview) -> None:
        """Store widths only for an explicit resize, after geometry settles."""
        try:
            widths = {column: int(tree.column(column, "width"))
                      for column in tree["columns"]}
        except tk.TclError:
            return
        self._save(lambda: self.settings.set_columns(key, widths))

    def register_sash(
        self, tab: str, key: str, widget: ttk.Panedwindow
    ) -> None:
        self._sashes_by_tab.setdefault(tab, []).append((key, widget))

        def released(_event) -> None:
            # Ignore releases in ordinary child widgets: the event target
            # must be the Panedwindow itself, as with an actual sash drag.
            widget.after_idle(lambda: self.capture_sash(key, widget))

        widget.bind("<ButtonRelease-1>", released, add="+")

    def capture_sash(self, key: str, widget: ttk.Panedwindow) -> None:
        try:
            if not widget.winfo_ismapped() or widget.winfo_width() <= 1:
                return
            position = int(widget.sashpos(0))
        except (tk.TclError, IndexError):
            return
        self._save(lambda: self.settings.set_sash(key, position))

    def restore_tab(self, tab: str) -> None:
        if tab in self._restored_tabs:
            return
        widgets = self._sashes_by_tab.get(tab, ())
        # ttk.Notebook does not allocate usable sizes to an inactive page.
        if not widgets or any(
            not widget.winfo_ismapped() or widget.winfo_width() <= 1
            for _key, widget in widgets
        ):
            return
        for key, widget in widgets:
            saved = self.settings.sashes.get(key)
            if saved is None:
                continue
            # Keep both panes reachable if the window has since shrunk.
            maximum = max(80, widget.winfo_width() - 100)
            widget.sashpos(0, min(saved, maximum))
        self._restored_tabs.add(tab)
