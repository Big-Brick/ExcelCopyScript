"""Tkinter application shell: reviewing imported values and creating aliases."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .alias_dialog import AliasManagerDialog
from .aliases import AliasStore
from .config import FILE_TYPES
from .models import CopyRecord
from .recent_directories import RecentDirectories
from .source_rows_panel import SourceRowsPanel
from .utils import present
from .workbook import WorkbookSession


class ExcelCopyApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Перенесення даних Excel")
        self.geometry("1260x720")
        self.minsize(900, 510)
        self.session: WorkbookSession | None = None
        self.aliases: AliasStore | None = None
        self.recent_directories = RecentDirectories()
        self.destination_var = tk.StringVar(value="Підсумковий файл не вибрано")
        self.source_var = tk.StringVar(value="Вхідний файл не вибрано")
        self.status_var = tk.StringVar(value="Виберіть підсумковий файл, потім вхідний.")
        self._build_ui()

    def _build_ui(self) -> None:
        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")
        ttk.Button(
            top, text="1. Відкрити підсумковий файл", command=self.open_destination,
        ).pack(side="left")
        self.source_button = ttk.Button(
            top, text="2. Відкрити вхідний / наступний файл",
            command=self.open_source, state="disabled",
        )
        self.source_button.pack(side="left", padx=(12, 0))
        ttk.Label(self, textvariable=self.destination_var, padding=(12, 0)).pack(anchor="w")
        ttk.Label(self, textvariable=self.source_var, padding=(12, 2)).pack(anchor="w")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=(8, 4))
        review_tab = ttk.Frame(self.notebook)
        create_tab = ttk.Frame(self.notebook)
        self.notebook.add(review_tab, text="Аналіз і видалення аліасів")
        self.notebook.add(create_tab, text="Створення аліасів")

        # Review: copied rows + unmatched source rows (independent sheet selector).
        review_panes = ttk.Panedwindow(review_tab, orient="horizontal")
        review_panes.pack(fill="both", expand=True)
        copied_frame, self.copied_tree = self._make_table(
            review_panes, "Скопійовані значення",
            ("Звідки", "Куди", "Значення", "Спосіб"), (180, 190, 140, 130),
        )
        self.review_sources = SourceRowsPanel(review_panes)
        review_panes.add(copied_frame, weight=1)
        review_panes.add(self.review_sources, weight=1)

        review_actions = ttk.Frame(review_tab, padding=(4, 8, 4, 6))
        review_actions.pack(fill="x")
        self.delete_alias_button = ttk.Button(
            review_actions, text="Видалити аліас вибраного запису",
            command=self.delete_selected_alias, state="disabled",
        )
        self.delete_alias_button.pack(side="left")
        self.alias_button = ttk.Button(
            review_actions, text="Керувати всіма аліасами",
            command=self.open_alias_manager, state="disabled",
        )
        self.alias_button.pack(side="left", padx=(12, 0))
        ttk.Label(
            review_tab,
            text="Видалення аліаса не змінює вже записані значення Excel. "
                 "Для створення нового аліаса перейдіть на другу вкладку.",
            padding=(4, 0, 4, 8),
        ).pack(anchor="w")
        self.copied_tree.bind("<<TreeviewSelect>>", self._on_copied_selection)

        # Creation: destination fields + unmatched source rows.
        create_panes = ttk.Panedwindow(create_tab, orient="horizontal")
        create_panes.pack(fill="both", expand=True)
        destination_frame, self.destination_tree = self._make_table(
            create_panes, "Без відповідника: вихідні стовпчики",
            ("Колонка", "Назва"), (80, 260),
        )
        self.create_sources = SourceRowsPanel(
            create_panes, on_selection_changed=self._on_selection,
        )
        create_panes.add(destination_frame, weight=1)
        create_panes.add(self.create_sources, weight=1)
        self.destination_tree.bind("<<TreeviewSelect>>", self._on_selection)

        create_actions = ttk.Frame(create_tab, padding=(4, 8, 4, 6))
        create_actions.pack(fill="x")
        self.copy_button = ttk.Button(
            create_actions, text="Копіювати вибрану пару + запам'ятати аліас",
            command=self.copy_selected, state="disabled",
        )
        self.copy_button.pack(anchor="w")
        ttk.Label(
            create_tab,
            text="Виберіть вихідний стовпчик та вхідний рядок на потрібному аркуші.",
            padding=(4, 0, 4, 8),
        ).pack(anchor="w")

        status = ttk.Frame(self, padding=(12, 4, 12, 12))
        status.pack(fill="x")
        ttk.Separator(status).pack(fill="x", pady=(0, 8))
        ttk.Label(status, textvariable=self.status_var, wraplength=1150).pack(anchor="w")

    @staticmethod
    def _make_table(parent, title, headings, widths):
        frame = ttk.Frame(parent, padding=4)
        ttk.Label(frame, text=title, font=("", 10, "bold")).pack(anchor="w", pady=(0, 5))
        container = ttk.Frame(frame)
        container.pack(fill="both", expand=True)
        tree = ttk.Treeview(container, columns=headings, show="headings", selectmode="browse")
        for heading, width in zip(headings, widths):
            tree.heading(heading, text=heading)
            tree.column(heading, width=width, minwidth=60, stretch=True)
        scrollbar_y = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar_y.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        return frame, tree

    def _on_selection(self, _event=None) -> None:
        can_copy = bool(
            self.session and self.session.current
            and self.destination_tree.selection()
            and self.create_sources.selected_source() is not None
        )
        self.copy_button.configure(state="normal" if can_copy else "disabled")

    def _selected_copy_record(self) -> CopyRecord | None:
        selected = self.copied_tree.selection()
        if not selected or self.session is None or self.session.current is None:
            return None
        try:
            return self.session.current.copied[int(selected[0])]
        except (IndexError, ValueError):
            return None

    def _record_has_alias(self, record: CopyRecord) -> bool:
        """Only allow deleting an alias that resolves to this record's destination."""
        if self.aliases is None or self.session is None:
            return False
        alias = self.aliases.get(
            self.session.template_id, record.source.name, record.source.sheet,
        )
        if not alias or alias.get("header") != record.destination.key:
            return False
        duplicates = sum(field.key == record.destination.key for field in self.session.fields)
        return duplicates == 1 or alias.get("column") == record.destination.column

    def _on_copied_selection(self, _event=None) -> None:
        record = self._selected_copy_record()
        enabled = record is not None and self._record_has_alias(record)
        self.delete_alias_button.configure(state="normal" if enabled else "disabled")

    def delete_selected_alias(self) -> None:
        """Remove mapping for future imports; never modify the existing workbook."""
        record = self._selected_copy_record()
        if record is None or not self._record_has_alias(record):
            return
        if not messagebox.askyesno(
            "Підтвердження видалення",
            f"Видалити аліас «{record.source.sheet}!{record.source.name}» → "
            f"«{record.destination.name}»?\n\n"
            "Уже скопійоване значення залишиться у підсумковому файлі.",
            parent=self,
        ):
            return
        try:
            self.aliases.delete(
                self.session.template_id, record.source.name, sheet=record.source.sheet,
            )
        except (OSError, ValueError) as exc:
            messagebox.showerror("Помилка видалення аліаса", str(exc), parent=self)
            return
        self._refresh_copied()

    def _remember_directory(self, kind: str, filename: str) -> None:
        """Do not fail a successful file open because UI preferences cannot be saved."""
        try:
            self.recent_directories.remember_file(kind, filename)
        except (OSError, ValueError) as exc:
            messagebox.showwarning(
                "Не вдалося запам'ятати теку",
                f"Файл відкрито, але теку не вдалося зберегти:\n{exc}",
            )

    def open_destination(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Підсумковий Excel-файл", filetypes=FILE_TYPES,
            initialdir=str(self.recent_directories.initial_directory("destination")),
        )
        if not chosen:
            return
        try:
            aliases = AliasStore()
            session = WorkbookSession(Path(chosen), aliases)
        except Exception as exc:
            messagebox.showerror("Помилка відкриття", str(exc))
            return
        self.aliases, self.session = aliases, session
        self._remember_directory("destination", chosen)
        self.destination_var.set(f"Підсумковий: {session.path}")
        self.source_var.set("Вхідний файл не вибрано")
        self.source_button.configure(state="normal")
        self.alias_button.configure(state="normal")
        self._set_source_sheets([])
        self._clear_tables()
        self.status_var.set(
            f"Знайдено {len(session.fields)} стовпчиків. "
            f"Наступний рядок: {session.next_empty_row()}."
        )

    def open_alias_manager(self) -> None:
        if self.session is None or self.aliases is None:
            return
        dialog = AliasManagerDialog(
            self, self.aliases, self.session.template_id, self.session.fields,
        )
        self.wait_window(dialog)
        self._refresh_copied()

    def open_source(self) -> None:
        if self.session is None:
            return
        chosen = filedialog.askopenfilename(
            title="Вхідний Excel-файл", filetypes=FILE_TYPES,
            initialdir=str(self.recent_directories.initial_directory(
                "source", fallback=self.session.path.parent,
            )),
        )
        if not chosen:
            return
        try:
            result = self.session.open_input(Path(chosen))
        except Exception as exc:
            messagebox.showerror("Помилка імпорту", str(exc))
            return
        self._remember_directory("source", chosen)
        self.source_var.set(f"Вхідний: {chosen} → рядок {result.row}")
        self._set_source_sheets(result.sheets)
        self._refresh_tables()

    def _set_source_sheets(self, sheets: list[str]) -> None:
        for panel in (self.review_sources, self.create_sources):
            panel.set_sheets(sheets)

    def _clear_tables(self) -> None:
        for tree in (self.copied_tree, self.destination_tree):
            tree.delete(*tree.get_children())
        for panel in (self.review_sources, self.create_sources):
            panel.refresh([])
        self._on_selection()
        self._on_copied_selection()

    def _refresh_copied(self) -> None:
        self.copied_tree.delete(*self.copied_tree.get_children())
        result = self.session.current if self.session is not None else None
        if result is not None:
            for index, record in enumerate(result.copied):
                method = record.method
                if method in ("Аліас", "Вручну") and not self._record_has_alias(record):
                    method = "Аліас видалено"
                self.copied_tree.insert(
                    "", "end", iid=str(index),
                    values=(
                        f"{record.source.sheet}!B{record.source.row}: {record.source.name}",
                        f"{record.destination.letter}: {record.destination.name}",
                        present(record.source.value), method,
                    ),
                )
        self._on_copied_selection()

    def _refresh_tables(self) -> None:
        self._clear_tables()
        if self.session is None or self.session.current is None:
            return
        result = self.session.current
        self._refresh_copied()
        for dest in result.unmatched_destinations:
            self.destination_tree.insert(
                "", "end", iid=str(dest.column), values=(dest.letter, dest.name),
            )
        for panel in (self.review_sources, self.create_sources):
            panel.refresh(result.unmatched_sources)
        self._on_selection()
        status = (
            f"{result.filename} → рядок {result.row}. Скопійовано: {len(result.copied)}; "
            f"не зіставлено вихідних стовпчиків: {len(result.unmatched_destinations)}; "
            f"вхідних рядків: {len(result.unmatched_sources)}. Зміни автоматично збережено."
        )
        if result.warnings:
            status += " Увага: " + "; ".join(result.warnings[:4])
        self.status_var.set(status)

    def copy_selected(self) -> None:
        if self.session is None:
            return
        source = self.create_sources.selected_source()
        destinations = self.destination_tree.selection()
        if source is None or not destinations:
            return
        try:
            warning = self.session.copy_manually(
                source.row, int(destinations[0]), sheet=source.sheet,
            )
        except Exception as exc:
            messagebox.showerror("Помилка копіювання", str(exc))
            return
        self._refresh_tables()
        if warning:
            messagebox.showwarning("Аліас не збережено", warning)


def main() -> None:
    ExcelCopyApp().mainloop()
