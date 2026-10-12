"""Schema migrations for persistent aliases.

Version 1 was indexed only by parameter name. Version 2 adds the input-sheet
dimension, with all old aliases belonging to ВИСНОВКИ.
"""

from __future__ import annotations

from .config import SOURCE_SHEET
from .utils import normalize

CURRENT_VERSION = 2


def migrate_alias_data(raw: dict) -> tuple[dict, bool]:
    """Return (data, migrated) without mutating the caller's object."""
    if not isinstance(raw, dict) or not isinstance(raw.get("templates"), dict):
        raise ValueError("Непідтримуваний формат файлу аліасів")

    version = raw.get("version")
    if version == CURRENT_VERSION:
        return raw, False
    if version != 1:
        raise ValueError(f"Непідтримувана версія файлу аліасів: {version}")

    summary_key = normalize(SOURCE_SHEET)
    templates = {}
    for template_id, entries in raw["templates"].items():
        if not isinstance(entries, dict):
            raise ValueError("Пошкоджений запис шаблону аліасів")
        migrated_entries = {}
        for source_key, record in entries.items():
            if not isinstance(record, dict):
                raise ValueError("Пошкоджений запис аліаса")
            migrated_entries[source_key] = {
                **record,
                "sheet_name": SOURCE_SHEET,
            }
        templates[template_id] = {summary_key: migrated_entries}

    return {"version": CURRENT_VERSION, "templates": templates}, True
