"""ExcelCopyScript: interactive transfer from ВИСНОВКИ to Дані_підрозділів."""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
from pathlib import Path

from .models import AliasEntry, DestinationField, SourceField
from .utils import normalize

class AliasStore:
    """Persistent input-name -> destination-header mapping, scoped to a header set."""

    def __init__(self, path: Path | None = None):
        self.path = path or (Path.home() / ".config" / "excel-copy-script" / "aliases.json")
        if self.path.exists():
            try:
                data = json.loads(self.path.read_text(encoding="utf-8"))
            except (OSError, ValueError) as exc:
                raise ValueError(f"Не вдалося прочитати файл аліасів {self.path}: {exc}") from exc
            if not isinstance(data, dict) or data.get("version") != 1 or not isinstance(data.get("templates"), dict):
                raise ValueError(f"Непідтримуваний формат файлу аліасів: {self.path}")
            self.data = data
        else:
            self.data = {"version": 1, "templates": {}}

    @staticmethod
    def template_id(fields: list[DestinationField]) -> str:
        headers = sorted(field.key for field in fields)
        return hashlib.sha256(json.dumps(headers, ensure_ascii=False).encode("utf-8")).hexdigest()

    def get(self, template_id: str, source_name: str) -> dict | None:
        entry = self.data["templates"].get(template_id, {}).get(normalize(source_name))
        return entry if isinstance(entry, dict) else None

    def list_for_template(self, template_id: str) -> list[AliasEntry]:
        """Return stored aliases for a destination template, including legacy records."""
        template = self.data["templates"].get(template_id, {})
        entries = []
        for key, value in template.items():
            if not isinstance(value, dict):
                continue
            entries.append(AliasEntry(
                source_key=key,
                source_name=value.get("source_name") or key,
                destination_key=value.get("header", ""),
                destination_column=value.get("column"),
            ))
        return sorted(entries, key=lambda entry: entry.source_name.casefold())

    def delete(self, template_id: str, source_name: str) -> bool:
        """Delete one alias and persist it; restore memory state on write failure.

        Returns False when the alias doesn't exist (no file write).
        """
        key = normalize(source_name)
        template = self.data["templates"].get(template_id)
        if template is None or key not in template:
            return False

        removed = template.pop(key)
        empty = not template
        if empty:
            del self.data["templates"][template_id]
        try:
            self._save()
        except Exception:
            if empty:
                self.data["templates"][template_id] = template
            template[key] = removed
            raise
        return True

    def set(self, template_id: str, source: SourceField, destination: DestinationField) -> None:
        key = source.key
        template = self.data["templates"].setdefault(template_id, {})
        previous = template.get(key)
        template[key] = {"header": destination.key, "column": destination.column,
                         "source_name": source.name if hasattr(source, "name") else key}
        try:
            self._save()
        except Exception:
            if previous is None:
                del template[key]
            else:
                template[key] = previous
            raise

    def _save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=self.path.parent,
                                             prefix=".aliases-", suffix=".tmp", delete=False) as stream:
                tmp_path = Path(stream.name)
                json.dump(self.data, stream, ensure_ascii=False, indent=2)
                stream.write("\n")
            os.replace(tmp_path, self.path)
        finally:
            if tmp_path is not None:
                tmp_path.unlink(missing_ok=True)
