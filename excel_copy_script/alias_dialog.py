"""Manage saved aliases for the currently opened summary workbook."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from .aliases import AliasStore
from .models import AliasEntry, DestinationField


class AliasManagerDialog(tk.Toplevel):
    """A modal view of aliases for the selected destination template."""

    def __init__(
        self,
        parent: tk.Misc,
        aliases: AliasStore,
        template_id: str,
        fields: list[DestinationField],
    ) -> None:
        super().__init__(parent)
        self.aliases = aliases
        self.template_id = template_id
        self.fields = fields

        self.title("Керування аліасами")
        self.geometry("830x460")
        self.minsize(630, 330)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._build_ui()
        self._refresh()
        self.grab_set()
        self.focus_set()

    def _build_ui(self) -> None:
        content = ttk.Frame(self, padding=12)
        content.pack(fill="both", expand=True)
        ttk.Label(
            content,
            text="Аліаси поточної підсумкової таблиці",
            font=("", 11, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            content,
            text="Видалення аліаса вплине лише на наступні імпорти. "
                 "Уже скопійовані значення не зміняться.",
            wraplength=785,
        ).pack(anchor="w", pady=(5, 10))

        table_frame = ttk.Frame(content)
        table_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(
            table_frame,
            columns=("source", "destination", "column"),
            show="headings",
            selectmode="browse",
        )
        for column, title, width in (
            ("source", "Назва у вхідному файлі", 320),
            ("destination", "Назва у підсумковому файлі", 340),
            ("column", "Стовпчик", 85),
        ):
            self.tree.heading(column, text=title)
            self.tree.column(column, width=width, minwidth=70, stretch=column != "column")
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")

        buttons = ttk.Frame(content)
        buttons.pack(fill="x", pady=(10, 0))
        self.delete_button = ttk.Button(
            buttons, text="Видалити вибраний аліас", command=self._delete_selected,
            state="disabled",
        )
        self.delete_button.pack(side="left")
        ttk.Button(buttons, text="Закрити", command=self.destroy).pack(side="right")
        self.count_label = ttk.Label(content)
        self.count_label.pack(anchor="w", pady=(8, 0))
        self.tree.bind("<<TreeviewSelect>>", self._on_selection)
        self.tree.bind("<Delete>", self._on_delete_key)

    def _target_display(self, entry: AliasEntry) -> tuple[str, str]:
        matching = [field for field in self.fields if field.key == entry.destination_key]
        if len(matching) == 1:
            return matching[0].name, matching[0].letter
        matching_column = next(
            (field for field in matching if field.column == entry.destination_column), None
        )
        if matching_column is not None:
            return matching_column.name, matching_column.letter
        column = str(entry.destination_column) if entry.destination_column is not None else "—"
        return f"{entry.destination_key} (не знайдено)", column

    def _refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        entries = self.aliases.list_for_template(self.template_id)
        for entry in entries:
            target_name, column = self._target_display(entry)
            self.tree.insert(
                "", "end", iid=entry.source_key,
                values=(entry.source_name, target_name, column),
            )
        self.count_label.configure(text=f"Збережено аліасів: {len(entries)}")
        self._on_selection()

    def _on_selection(self, _event=None) -> None:
        self.delete_button.configure(
            state="normal" if self.tree.selection() else "disabled",
        )

    def _on_delete_key(self, _event=None) -> None:
        if self.tree.selection():
            self._delete_selected()

    def _delete_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        key = selection[0]
        label = self.tree.item(key, "values")
        source_name = label[0] if label else key
        destination_name = label[1] if label else ""
        if not messagebox.askyesno(
            "Підтвердження видалення",
            f"Видалити аліас «{source_name}» → «{destination_name}»?\n\n"
            "Подальші імпорти більше не використовуватимуть цю відповідність.",
            parent=self,
        ):
            return
        try:
            self.aliases.delete(self.template_id, key)
        except (OSError, ValueError) as exc:
            messagebox.showerror("Помилка видалення аліаса", str(exc), parent=self)
            return
        self._refresh()
