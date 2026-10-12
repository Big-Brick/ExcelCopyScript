"""ExcelCopyScript: interactive transfer from ВИСНОВКИ to Дані_підрозділів."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .alias_dialog import AliasManagerDialog
from .aliases import AliasStore
from .config import FILE_TYPES
from .recent_directories import RecentDirectories
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
        self._unmatched_source_ids: dict[str, tuple[str, int]] = {}
        self.recent_directories = RecentDirectories()
        self.destination_var = tk.StringVar(value="Підсумковий файл не вибрано")
        self.source_var = tk.StringVar(value="Вхідний файл не вибрано")
        self.status_var = tk.StringVar(value="Виберіть підсумковий файл, потім вхідний.")
        self._build_ui()

    def _build_ui(self) -> None:
        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")
        ttk.Button(top, text="1. Відкрити підсумковий файл", command=self.open_destination).pack(side="left")
        self.source_button = ttk.Button(top, text="2. Відкрити вхідний / наступний файл",
                                        command=self.open_source, state="disabled")
        self.source_button.pack(side="left", padx=(12, 0))
        self.alias_button = ttk.Button(
            top, text="Керувати аліасами", command=self.open_alias_manager,
            state="disabled",
        )
        self.alias_button.pack(side="left", padx=(12, 0))
        ttk.Label(self, textvariable=self.destination_var, padding=(12, 0)).pack(anchor="w")
        ttk.Label(self, textvariable=self.source_var, padding=(12, 2)).pack(anchor="w")

        panes = ttk.Panedwindow(self, orient="horizontal")
        panes.pack(fill="both", expand=True, padx=12, pady=(8, 4))
        copied_frame, self.copied_tree = self._make_table(
            panes, "Скопійовані значення", ("Звідки", "Куди", "Значення", "Спосіб"), (180, 190, 140, 130))
        destination_frame, self.destination_tree = self._make_table(
            panes, "Без відповідника: вихідні стовпчики", ("Колонка", "Назва"), (80, 260))
        source_frame, self.source_tree = self._make_table(
            panes, "Без відповідника: вхідні рядки",
            ("Аркуш", "Рядок", "Назва", "Значення"), (130, 70, 200, 110))
        for frame in (copied_frame, destination_frame, source_frame):
            panes.add(frame, weight=1)

        bottom = ttk.Frame(self, padding=(12, 8, 12, 12))
        bottom.pack(fill="x")
        self.copy_button = ttk.Button(bottom, text="Копіювати вибрану пару + запам'ятати аліас",
                                      command=self.copy_selected, state="disabled")
        self.copy_button.pack(anchor="w")
        ttk.Label(bottom, text="Виберіть один рядок у кожному зі списків «Без відповідника».").pack(anchor="w", pady=(5, 3))
        ttk.Separator(bottom).pack(fill="x", pady=(3, 8))
        ttk.Label(bottom, textvariable=self.status_var, wraplength=1150).pack(anchor="w")
        self.destination_tree.bind("<<TreeviewSelect>>", self._on_selection)
        self.source_tree.bind("<<TreeviewSelect>>", self._on_selection)

    def _make_table(self, parent, title, headings, widths):
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

    def _on_selection(self, _event=None):
        can_copy = bool(self.session and self.session.current and
                        self.destination_tree.selection() and self.source_tree.selection())
        self.copy_button.configure(state="normal" if can_copy else "disabled")

    def _remember_directory(self, kind: str, filename: str) -> None:
        """Do not fail a successful file open just because UI preferences cannot be saved."""
        try:
            self.recent_directories.remember_file(kind, filename)
        except (OSError, ValueError) as exc:
            messagebox.showwarning(
                "Не вдалося запам'ятати теку",
                f"Файл відкрито, але теку не вдалося зберегти:\n{exc}",
            )

    def open_destination(self):
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
        self._clear_tables()
        self.status_var.set(f"Знайдено {len(session.fields)} стовпчиків. Наступний рядок: {session.next_empty_row()}.")

    def open_alias_manager(self):
        if self.session is None or self.aliases is None:
            return
        AliasManagerDialog(
            self, self.aliases, self.session.template_id, self.session.fields,
        )

    def open_source(self):
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
        self._refresh_tables()

    def _clear_tables(self):
        self._unmatched_source_ids.clear()
        for tree in (self.copied_tree, self.destination_tree, self.source_tree):
            tree.delete(*tree.get_children())
        self._on_selection()

    def _refresh_tables(self):
        self._clear_tables()
        if self.session is None or self.session.current is None:
            return
        result = self.session.current
        for record in result.copied:
            self.copied_tree.insert("", "end", values=(f"{record.source.sheet}!B{record.source.row}: {record.source.name}",
                f"{record.destination.letter}: {record.destination.name}", present(record.source.value), record.method))
        for dest in result.unmatched_destinations:
            self.destination_tree.insert("", "end", iid=str(dest.column), values=(dest.letter, dest.name))
        for source in result.unmatched_sources:
            item_id = self.source_tree.insert(
                "", "end",
                values=(source.sheet, source.row, source.name, present(source.value)),
            )
            self._unmatched_source_ids[item_id] = (source.sheet, source.row)
        status = (f"{result.filename} → рядок {result.row}. Скопійовано: {len(result.copied)}; "
                  f"не зіставлено вихідних стовпчиків: {len(result.unmatched_destinations)}; "
                  f"вхідних рядків: {len(result.unmatched_sources)}. Зміни автоматично збережено.")
        if result.warnings:
            status += " Увага: " + "; ".join(result.warnings[:4])
        self.status_var.set(status)

    def copy_selected(self):
        if self.session is None:
            return
        sources, destinations = self.source_tree.selection(), self.destination_tree.selection()
        if not sources or not destinations:
            return
        try:
            sheet, row = self._unmatched_source_ids[sources[0]]
            warning = self.session.copy_manually(row, int(destinations[0]), sheet=sheet)
        except Exception as exc:
            messagebox.showerror("Помилка копіювання", str(exc))
            return
        self._refresh_tables()
        if warning:
            messagebox.showwarning("Аліас не збережено", warning)


def main() -> None:
    ExcelCopyApp().mainloop()
