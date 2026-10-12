"""Read, write and migrate sheet-scoped aliases for a destination template."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
from pathlib import Path

from .config import SOURCE_SHEET
from .migrations import CURRENT_VERSION, migrate_alias_data
from .models import AliasEntry, DestinationField, SourceField
from .utils import normalize


class AliasStore:
    """Aliases indexed by destination template, source sheet and parameter."""

    def __init__(self, path: Path | None = None):
        self.path = path or (Path.home() / ".config" / "excel-copy-script" / "aliases.json")
        if self.path.exists():
            try:
                raw = json.loads(self.path.read_text(encoding="utf-8"))
            except (OSError, ValueError) as exc:
                raise ValueError(f"Не вдалося прочитати файл аліасів {self.path}: {exc}") from exc
            try:
                self.data, migrated = migrate_alias_data(raw)
            except ValueError as exc:
                raise ValueError(f"Непідтримуваний формат файлу аліасів {self.path}: {exc}") from exc
            if migrated:
                # Keep the exact pre-migration JSON as a one-time backup.
                backup = self.path.with_name(self.path.name + ".v1.bak")
                if not backup.exists():
                    shutil.copy2(self.path, backup)
                self._save()
        else:
            self.data = {"version": CURRENT_VERSION, "templates": {}}

    @staticmethod
    def template_id(fields: list[DestinationField]) -> str:
        headers = sorted(field.key for field in fields)
        return hashlib.sha256(json.dumps(headers, ensure_ascii=False).encode("utf-8")).hexdigest()

    def get(self, template_id: str, source_name: str, sheet: str = SOURCE_SHEET) -> dict | None:
        entry = (self.data["templates"].get(template_id, {})
                 .get(normalize(sheet), {}).get(normalize(source_name)))
        return entry if isinstance(entry, dict) else None

    def list_for_template(self, template_id: str) -> list[AliasEntry]:
        """List alias mappings from every sheet for the current destination template."""
        entries = []
        for sheet_key, mappings in self.data["templates"].get(template_id, {}).items():
            for source_key, record in mappings.items():
                if not isinstance(record, dict):
                    continue
                entries.append(AliasEntry(
                    source_key=source_key,
                    source_name=record.get("source_name") or source_key,
                    destination_key=record.get("header", ""),
                    destination_column=record.get("column"),
                    sheet=record.get("sheet_name") or sheet_key,
                ))
        return sorted(entries, key=lambda e: (e.sheet.casefold(), e.source_name.casefold()))

    def migrate_legacy_template(
        self, legacy_id: str, current_id: str, fields: list[DestinationField]
    ) -> None:
        """Move aliases from a template that included ordinal column A.

        Keeping the old template would resurrect deleted aliases on next launch,
        so remove the obsolete key after migrating usable mappings.
        """
        if legacy_id == current_id:
            return
        templates = self.data["templates"]
        old = templates.get(legacy_id)
        if not isinstance(old, dict) or not old:
            return

        valid: dict[str, dict] = {}
        for sheet_key, entries in old.items():
            valid_entries = {}
            for source_key, record in entries.items():
                if not isinstance(record, dict) or record.get("column") == 1:
                    continue
                candidates = [field for field in fields if field.key == record.get("header")]
                if len(candidates) == 1 or (
                    len(candidates) > 1
                    and any(field.column == record.get("column") for field in candidates)
                ):
                    valid_entries[source_key] = record
            if valid_entries:
                valid[sheet_key] = valid_entries

        original_current = templates.get(current_id)
        merged = {sheet: dict(entries) for sheet, entries in valid.items()}
        for sheet, entries in (original_current or {}).items():
            merged.setdefault(sheet, {}).update(entries)

        templates.pop(legacy_id)
        if merged:
            templates[current_id] = merged
        try:
            self._save()
        except Exception:
            templates[legacy_id] = old
            if original_current is None:
                templates.pop(current_id, None)
            else:
                templates[current_id] = original_current
            raise

    def delete(self, template_id: str, source_name: str, sheet: str = SOURCE_SHEET) -> bool:
        """Delete only the selected (sheet, parameter) alias."""
        templates = self.data["templates"]
        template = templates.get(template_id)
        sheet_key, source_key = normalize(sheet), normalize(source_name)
        if not template or source_key not in template.get(sheet_key, {}):
            return False
        mappings = template[sheet_key]
        old = mappings.pop(source_key)
        if not mappings:
            del template[sheet_key]
        if not template:
            del templates[template_id]
        try:
            self._save()
        except Exception:
            templates.setdefault(template_id, {}).setdefault(sheet_key, {})[source_key] = old
            raise
        return True

    def set(self, template_id: str, source: SourceField, destination: DestinationField) -> None:
        sheet_key, source_key = normalize(source.sheet), source.key
        template = self.data["templates"].setdefault(template_id, {})
        mappings = template.setdefault(sheet_key, {})
        old = mappings.get(source_key)
        mappings[source_key] = {
            "header": destination.key, "column": destination.column,
            "source_name": source.name, "sheet_name": source.sheet,
        }
        try:
            self._save()
        except Exception:
            if old is None:
                del mappings[source_key]
                if not mappings:
                    del template[sheet_key]
                if not template:
                    del self.data["templates"][template_id]
            else:
                mappings[source_key] = old
            raise

    def _save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="w", encoding="utf-8", dir=self.path.parent,
                prefix=".aliases-", suffix=".tmp", delete=False
            ) as stream:
                tmp_path = Path(stream.name)
                json.dump(self.data, stream, ensure_ascii=False, indent=2)
                stream.write("\n")
            os.replace(tmp_path, self.path)
        finally:
            if tmp_path is not None:
                tmp_path.unlink(missing_ok=True)
