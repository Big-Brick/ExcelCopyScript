"""Read, write and migrate sheet- and optional row-scoped aliases."""

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
    """Aliases indexed by template, sheet, name and optional source row."""

    def __init__(self, path: Path | None = None, *, profile=None):
        # Legacy standalone AliasStore remains readable for old integrations.
        # The GUI uses the profile-backed store exclusively.
        self.profile = profile
        if profile is not None:
            self.path = profile.path
            self.data = profile.aliases_data
            return
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
                backup = self.path.with_name(self.path.name + f".v{raw['version']}.bak")
                if not backup.exists():
                    shutil.copy2(self.path, backup)
                self._save()
        else:
            self.data = {"version": CURRENT_VERSION, "templates": {}}

    @staticmethod
    def template_id(fields: list[DestinationField]) -> str:
        headers = sorted(field.key for field in fields)
        return hashlib.sha256(json.dumps(headers, ensure_ascii=False).encode("utf-8")).hexdigest()

    @staticmethod
    def _check_row(row: int | None) -> None:
        if row is not None and (type(row) is not int or row < 1):
            raise ValueError("Номер рядка аліаса має бути додатним цілим числом")

    def _group(self, template_id: str, source_name: str, sheet: str) -> dict | None:
        group = (self.data["templates"].get(template_id, {})
                 .get(normalize(sheet), {}).get(normalize(source_name)))
        return group if isinstance(group, dict) else None

    def get_exact(
        self, template_id: str, source_name: str,
        sheet: str = SOURCE_SHEET, row: int | None = None,
    ) -> dict | None:
        """Retrieve precisely the requested general or row-scoped alias."""
        self._check_row(row)
        group = self._group(template_id, source_name, sheet)
        if group is None:
            return None
        record = group.get("default") if row is None else group.get("rows", {}).get(str(row))
        return record if isinstance(record, dict) else None

    def get(
        self, template_id: str, source_name: str,
        sheet: str = SOURCE_SHEET, row: int | None = None,
    ) -> dict | None:
        """Match row-specific first, then fall back to the general alias."""
        if row is not None:
            scoped = self.get_exact(template_id, source_name, sheet, row)
            if scoped is not None:
                return scoped
        return self.get_exact(template_id, source_name, sheet)

    def list_for_template(self, template_id: str) -> list[AliasEntry]:
        """Include both general and row-specific mappings from all sheets."""
        entries = []
        for sheet_key, mappings in self.data["templates"].get(template_id, {}).items():
            for source_key, group in mappings.items():
                if not isinstance(group, dict):
                    continue
                records = [(None, group.get("default"))]
                records.extend(
                    (int(row), entry)
                    for row, entry in group.get("rows", {}).items()
                )
                for row, record in records:
                    if not isinstance(record, dict):
                        continue
                    entries.append(AliasEntry(
                        source_key=source_key,
                        source_name=record.get("source_name") or source_key,
                        destination_key=record.get("header", ""),
                        destination_column=record.get("column"),
                        sheet=record.get("sheet_name") or sheet_key,
                        source_column=record.get("source_column", 3),
                        source_row=row,
                    ))
        return sorted(entries, key=lambda e: (
            e.sheet.casefold(), e.source_name.casefold(),
            e.source_row is not None, e.source_row or 0,
        ))

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
            for source_key, group in entries.items():
                if not isinstance(group, dict):
                    continue
                def valid_record(record):
                    if not isinstance(record, dict) or record.get("column") == 1:
                        return False
                    candidates = [field for field in fields if field.key == record.get("header")]
                    return len(candidates) == 1 or (
                        len(candidates) > 1
                        and any(field.column == record.get("column") for field in candidates)
                    )
                general = group.get("default")
                rows = group.get("rows", {})
                filtered_rows = {row: record for row, record in rows.items()
                                 if valid_record(record)}
                if valid_record(general) or filtered_rows:
                    valid_entries[source_key] = {
                        "default": general if valid_record(general) else None,
                        "rows": filtered_rows,
                    }
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

    def delete(
        self, template_id: str, source_name: str,
        sheet: str = SOURCE_SHEET, row: int | None = None,
    ) -> bool:
        """Delete only one exact alias; never erase its sibling variant."""
        self._check_row(row)
        templates = self.data["templates"]
        sheet_key, source_key = normalize(sheet), normalize(source_name)
        group = self._group(template_id, source_name, sheet)
        if group is None or self.get_exact(template_id, source_name, sheet, row) is None:
            return False
        # Copy containers for rollback if saving fails.
        import copy
        snapshot = copy.deepcopy(templates.get(template_id))
        if row is None:
            group["default"] = None
        else:
            group["rows"].pop(str(row))
        if group.get("default") is None and not group["rows"]:
            del templates[template_id][sheet_key][source_key]
            if not templates[template_id][sheet_key]:
                del templates[template_id][sheet_key]
            if not templates[template_id]:
                del templates[template_id]
        try:
            self._save()
        except Exception:
            templates[template_id] = snapshot
            raise
        return True

    def set(
        self, template_id: str, source: SourceField,
        destination: DestinationField, row: int | None = None,
    ) -> None:
        """Save a general alias or one restricted to a specific row."""
        self._check_row(row)
        sheet_key, source_key = normalize(source.sheet), source.key
        templates = self.data["templates"]
        import copy
        snapshot = copy.deepcopy(templates.get(template_id))
        mappings = templates.setdefault(template_id, {}).setdefault(sheet_key, {})
        group = mappings.setdefault(source_key, {"default": None, "rows": {}})
        record = {
            "header": destination.key, "column": destination.column,
            "source_name": source.name, "sheet_name": source.sheet,
            "source_column": source.value_column,
        }
        if row is None:
            group["default"] = record
        else:
            group["rows"][str(row)] = record
        try:
            self._save()
        except Exception:
            if snapshot is None:
                templates.pop(template_id, None)
            else:
                templates[template_id] = snapshot
            raise

    @classmethod
    def from_profile(cls, profile) -> "AliasStore":
        return cls(profile=profile)

    def reload(self) -> "AliasStore":
        if self.profile is not None:
            from .profiles import Profile
            return type(self).from_profile(Profile.load(self.profile.path))
        return type(self)(self.path)

    def _save(self) -> None:
        if self.profile is not None:
            self.profile.save()
            return
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
