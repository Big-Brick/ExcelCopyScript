"""Migrations for persistent aliases.

v1: template -> source name -> destination.
v2: template -> sheet -> source name -> destination.
v3: v2 plus source value column (C by default).
v4: a parameter may have a general alias and row-specific overrides:
    template -> sheet -> source name -> {default: record | null, rows: {str(row): record}}.
"""

from __future__ import annotations

from .config import SOURCE_SHEET
from .utils import normalize

CURRENT_VERSION = 4


def migrate_alias_data(raw: dict) -> tuple[dict, bool]:
    """Return v4 alias data, retaining all pre-existing general mappings."""
    if not isinstance(raw, dict) or not isinstance(raw.get("templates"), dict):
        raise ValueError("Непідтримуваний формат файлу аліасів")
    version = raw.get("version")
    if version == CURRENT_VERSION:
        return raw, False
    if version not in (1, 2, 3):
        raise ValueError(f"Непідтримувана версія файлу аліасів: {version}")

    data = {}
    for template_id, templates in raw["templates"].items():
        if not isinstance(templates, dict):
            raise ValueError("Пошкоджений шаблон аліасів")
        if version == 1:
            converted = {}
            for key, value in templates.items():
                if not isinstance(value, dict):
                    raise ValueError("Пошкоджений запис аліаса")
                converted[key] = {**value, "sheet_name": SOURCE_SHEET}
            templates = {normalize(SOURCE_SHEET): converted}
        migrated_sheets = {}
        for sheet_key, records in templates.items():
            if not isinstance(records, dict):
                raise ValueError("Пошкоджений запис аркуша аліасів")
            migrated_sheets[sheet_key] = {}
            for key, entry in records.items():
                if not isinstance(entry, dict):
                    raise ValueError("Пошкоджений запис аліаса")
                if version < 3:
                    entry = {**entry, "source_column": 3}
                migrated_sheets[sheet_key][key] = {
                    "default": dict(entry),
                    "rows": {},
                }
        data[template_id] = migrated_sheets
    return {"version": CURRENT_VERSION, "templates": data}, True
