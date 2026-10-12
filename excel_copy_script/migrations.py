"""Migrations for persistent aliases.

v1: template -> source name -> destination.
v2: template -> sheet -> source name -> destination.
v3: v2 plus the source value column (C by default for legacy aliases).
"""

from __future__ import annotations

from .config import SOURCE_SHEET
from .utils import normalize

CURRENT_VERSION = 3


def migrate_alias_data(raw: dict) -> tuple[dict, bool]:
    """Return version-3 data and whether a migration occurred; preserve input."""
    if not isinstance(raw, dict) or not isinstance(raw.get("templates"), dict):
        raise ValueError("Непідтримуваний формат файлу аліасів")
    version = raw.get("version")
    if version == CURRENT_VERSION:
        return raw, False
    if version not in (1, 2):
        raise ValueError(f"Непідтримувана версія файлу аліасів: {version}")

    data = {}
    for template_id, templates in raw["templates"].items():
        if not isinstance(templates, dict):
            raise ValueError("Пошкоджений шаблон аліасів")
        if version == 1:
            templates = {normalize(SOURCE_SHEET): {
                key: {**value, "sheet_name": SOURCE_SHEET}
                for key, value in templates.items()
                if isinstance(value, dict)
            }}
        migrated_sheets = {}
        for sheet_key, records in templates.items():
            if not isinstance(records, dict):
                raise ValueError("Пошкоджений запис аркуша аліасів")
            migrated_sheets[sheet_key] = {}
            for key, entry in records.items():
                if not isinstance(entry, dict):
                    raise ValueError("Пошкоджений запис аліаса")
                migrated_sheets[sheet_key][key] = {**entry, "source_column": 3}
        data[template_id] = migrated_sheets
    return {"version": CURRENT_VERSION, "templates": data}, True
