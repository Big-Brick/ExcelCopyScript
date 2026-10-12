"""Modal chooser for an ambiguous existing destination row."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from .models import RowSelectionRequest
from .workbook import RowSelectionCancelled


class RowSelectionDialog(tk.Toplevel):
    """Choose an existing row to overwrite, a new row, or cancel importing."""

    def __init__(self, parent: tk.Misc, request: RowSelectionRequest):
        super().__init__(parent)
        self.title("Вибір рядка для імпорту")
        self.geometry("970x460")
        self.minsize(750, 340)
        self.transient(parent)
        self.result: int | None = None
        self.accepted = False
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        content = ttk.Frame(self, padding=12)
        content.pack(fill="both", expand=True)
        ttk.Label(
            content,
            text=f"Файл «{request.filename}»: знайдено можливі існуючі рядки",
            font=("", 11, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            content,
            text="Виберіть рядок для перезапису значень із аліасів, "
                 "або запишіть їх у новий рядок. Інші колонки не зміняться.",
            wraplength=920,
        ).pack(anchor="w", pady=(6, 10))

        area = ttk.Frame(content)
        area.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(
            area, columns=("row", "matched", "differences"),
            show="headings", selectmode="browse",
        )
        for column, label, width in (
            ("row", "Рядок", 85),
            ("matched", "Збіги", 265),
            ("differences", "Відмінності (поточне → нове)", 575),
        ):
            self.tree.heading(column, text=label)
            self.tree.column(column, width=width, minwidth=70, stretch=column != "row")
        self._rows: dict[str, int] = {}
        for candidate in sorted(
            request.candidates,
            key=lambda item: (-len(item.matching), item.row),
        ):
            matched = f"{len(candidate.matching)}/{candidate.total}: " + ", ".join(candidate.matching)
            if candidate.complete:
                matched += " (повний збіг)"
            iid = self.tree.insert(
                "", "end",
                values=(
                    candidate.row, matched,
                    "; ".join(candidate.differing) if candidate.differing else "—",
                ),
            )
            self._rows[iid] = candidate.row
        scrollbar = ttk.Scrollbar(area, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        buttons = ttk.Frame(content)
        buttons.pack(fill="x", pady=(12, 0))
        self.overwrite_button = ttk.Button(
            buttons, text="Перезаписати вибраний рядок",
            command=self._overwrite, state="disabled",
        )
        self.overwrite_button.pack(side="left")
        ttk.Button(
            buttons, text="Записати в новий рядок", command=self._new_row,
        ).pack(side="left", padx=12)
        ttk.Button(
            buttons, text="Скасувати імпорт", command=self.destroy,
        ).pack(side="right")
        self.tree.bind(
            "<<TreeviewSelect>>",
            lambda _event: self.overwrite_button.configure(
                state="normal" if self.tree.selection() else "disabled",
            ),
        )
        self.tree.bind("<Double-1>", lambda _event: self._overwrite())
        self.grab_set()
        self.focus_set()

    def _overwrite(self) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        self.result = self._rows[selected[0]]
        self.accepted = True
        self.destroy()

    def _new_row(self) -> None:
        self.result = None
        self.accepted = True
        self.destroy()

    @classmethod
    def choose(cls, parent: tk.Misc, request: RowSelectionRequest) -> int | None:
        dialog = cls(parent, request)
        parent.wait_window(dialog)
        if not dialog.accepted:
            raise RowSelectionCancelled()
        return dialog.result
