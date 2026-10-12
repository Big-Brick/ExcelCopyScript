"""Non-destructive conversion of the legacy aliases.json into a profile."""

from __future__ import annotations

import json
from pathlib import Path

from .migrations import CURRENT_VERSION, migrate_alias_data
from .workbook_layout import WorkbookLayout

PROFILE_VERSION = 1


def build_profile_from_legacy(legacy_path: Path, layout: WorkbookLayout | None = None) -> dict:
    """Read but NEVER modify or rename the original aliases.json."""
    data = {"version": CURRENT_VERSION, "templates": {}}
    if legacy_path.is_file():
        try:
            raw = json.loads(legacy_path.read_text(encoding="utf-8"))
            data, _ = migrate_alias_data(raw)
        except (OSError, ValueError, TypeError) as exc:
            raise ValueError(
                f"Не вдалося перенести аліаси з «{legacy_path}»: {exc}. "
                "Старий файл не змінено."
            ) from exc
    return {
        "profile_version": PROFILE_VERSION,
        "layout": (layout or WorkbookLayout()).to_dict(),
        "aliases": data,
    }
