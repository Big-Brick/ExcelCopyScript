"""Reusable worksheet-filtered view of unmatched source fields."""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable, Sequence
from tkinter import ttk

from .config import SOURCE_SHEET
from .models import SourceField
from .utils import present


class SourceRowsPanel(ttk.Frame):
    """One sheet selector and one independent list of unmatched input rows."""

    def __init__(
        self, parent: tk.Misc, on_selection_changed: Callable[[], None] | None = None
    ) -> None:
        super().__init__(parent, padding=4)
        self._rows: list[SourceField] = []
        self._sources_by_item_id: dict[str, SourceField] = {}
        self._on_selection_changed = on_selection_changed
        self.selected_sheet = tk.StringVar(value="")

        ttk.Label(self, text="Без відповідника: вхідні рядки",
                  font=("", 10, "bold")).pack(anchor="w", pady=(0, 5))
        selector = ttk.Frame(self)
        selector.pack(fill="x", pady=(0, 6))
        ttk.Label(selector, text="Аркуш:").pack(side="left")
        self.sheet_combo = ttk.Combobox(
            selector, textvariable=self.selected_sheet, state="disabled", width=26,
        )
        self.sheet_combo.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_changed)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(
            container, columns=("row", "name", "value"), show="headings", selectmode="browse",
        )
        for column, title, width in (
            ("row", "Рядок", 70),
            ("name", "Назва", 240),
            ("value", "Значення", 130),
        ):
            self.tree.heading(column, text=title)
            self.tree.column(column, width=width, minwidth=60, stretch=True)
        scroll = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._notify_selection)

    def set_sheets(self, sheets: Sequence[str]) -> None:
        """Reset selector when opening a different input workbook."""
        sheets = list(sheets)
        self.sheet_combo.configure(values=sheets, state="readonly" if sheets else "disabled")
        self.selected_sheet.set(
            SOURCE_SHEET if SOURCE_SHEET in sheets else (sheets[0] if sheets else "")
        )
        self.refresh([])

    def refresh(self, rows: Sequence[SourceField] | None = None) -> None:
        """Update this view; retain the current worksheet even after manual copying."""
        if rows is not None:
            self._rows = list(rows)
        self._sources_by_item_id.clear()
        self.tree.delete(*self.tree.get_children())
        sheet = self.selected_sheet.get()
        for source in self._rows:
            if source.sheet != sheet:
                continue
            item = self.tree.insert(
                "", "end", values=(source.row, source.name, present(source.value)),
            )
            self._sources_by_item_id[item] = source
        self._notify_selection()

    def selected_source(self) -> SourceField | None:
        items = self.tree.selection()
        return self._sources_by_item_id.get(items[0]) if items else None

    def _on_sheet_changed(self, _event=None) -> None:
        self.refresh()

    def _notify_selection(self, _event=None) -> None:
        if self._on_selection_changed is not None:
            self._on_selection_changed()
