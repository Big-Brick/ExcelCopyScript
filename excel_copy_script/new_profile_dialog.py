"""Modal GUI for creating a completely empty, independent Excel profile."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .profiles import Profile
from .workbook_layout import WorkbookLayout


# Each field is configured independently; the dialog never copies active aliases.
FIELDS = (
    ("destination_sheet", "Вихідний аркуш", False),
    ("header_row", "Рядок заголовків", True),
    ("first_data_row", "Перший рядок даних", True),
    ("first_destination_column", "Перша вихідна колонка", True),
    ("source_name_column", "Колонка назви параметра", True),
    ("first_source_row", "Перший вхідний рядок", True),
    ("source_value_first_column", "Перша колонка значення", True),
    ("source_value_last_column", "Остання колонка значення", True),
    ("preferred_source_sheet", "Типовий вхідний аркуш", False),
)


def suggested_profile_path(folder: Path) -> Path:
    """Suggest an unused filename without creating files or directories."""
    folder = Path(folder)
    first = folder / "new-profile.json"
    if not first.exists():
        return first
    number = 2
    while (candidate := folder / f"new-profile-{number}.json").exists():
        number += 1
    return candidate


class NewProfileDialog(tk.Toplevel):
    """A validated dialog; cancel never creates a file or switches profiles."""

    def __init__(self, parent: tk.Misc, profiles_directory: Path):
        super().__init__(parent)
        self.title("Новий профіль")
        self.geometry("760x600")
        self.minsize(680, 570)
        self.transient(parent)
        self.profile: Profile | None = None
        self.path_var = tk.StringVar(value=str(suggested_profile_path(profiles_directory)))
        defaults = WorkbookLayout().to_dict()
        self.fields = {key: tk.StringVar(value=str(defaults[key])) for key, _, _ in FIELDS}

        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(
            body, text="Створити порожній профіль", font=("", 12, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            body, text="Профіль почнеться без аліасів. Налаштування іншого "
                       "профілю не копіюються.",
        ).pack(anchor="w", pady=(3, 13))

        location = ttk.LabelFrame(body, text="Новий файл профілю", padding=8)
        location.pack(fill="x")
        ttk.Entry(location, textvariable=self.path_var).pack(
            side="left", fill="x", expand=True,
        )
        ttk.Button(
            location, text="Огляд…", command=self._browse_path,
        ).pack(side="left", padx=(8, 0))

        settings = ttk.LabelFrame(body, text="Структура Excel", padding=10)
        settings.pack(fill="both", expand=True, pady=(12, 0))
        for index, (key, label, is_number) in enumerate(FIELDS):
            ttk.Label(settings, text=label).grid(
                row=index, column=0, sticky="w", pady=3,
            )
            field = ttk.Entry(
                settings, textvariable=self.fields[key],
                width=11 if is_number else 38,
            )
            field.grid(row=index, column=1, sticky="ew", padx=(12, 0), pady=3)
        settings.columnconfigure(1, weight=1)
        ttk.Label(
            settings,
            text="Колонки вказуються числами: A = 1, B = 2, C = 3 тощо.",
        ).grid(row=len(FIELDS), column=0, columnspan=2, sticky="w", pady=(10, 0))

        buttons = ttk.Frame(body)
        buttons.pack(fill="x", pady=(14, 0))
        self.create_button = ttk.Button(
            buttons, text="Створити профіль", command=self._create,
        )
        self.create_button.pack(side="right")
        ttk.Button(
            buttons, text="Скасувати", command=self.destroy,
        ).pack(side="right", padx=(0, 12))
        self.bind("<Escape>", lambda _event: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.grab_set()
        self.focus_set()

    def _browse_path(self) -> None:
        path = Path(self.path_var.get().strip()).expanduser()
        chosen = filedialog.asksaveasfilename(
            parent=self, title="Куди зберегти новий профіль",
            filetypes=[("Профіль JSON", "*.json")],
            defaultextension=".json",
            initialdir=str(path.parent),
            initialfile=path.name,
        )
        if chosen:
            self.path_var.set(chosen)

    def _create(self) -> None:
        raw_path = self.path_var.get().strip()
        if not raw_path or Path(raw_path).suffix.lower() != ".json":
            messagebox.showerror(
                "Некоректний файл", "Вкажіть шлях до нового файлу з розширенням .json.",
                parent=self,
            )
            return

        values = {}
        try:
            for key, label, is_number in FIELDS:
                text = self.fields[key].get().strip()
                if is_number:
                    try:
                        values[key] = int(text)
                    except (TypeError, ValueError) as exc:
                        raise ValueError(f"«{label}»: вкажіть ціле число") from exc
                else:
                    values[key] = text
            layout = WorkbookLayout.from_dict(values)
            # FileExistsError is intentional: NEVER overwrite a saved profile.
            profile = Profile.create_empty(Path(raw_path), layout)
        except (OSError, ValueError) as exc:
            messagebox.showerror("Не вдалося створити профіль", str(exc), parent=self)
            return
        self.profile = profile
        self.destroy()
