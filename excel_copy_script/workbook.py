"""Import spreadsheet values using persisted sheet-scoped aliases only."""

from __future__ import annotations

import os
import shutil
import tempfile
from collections import defaultdict
from pathlib import Path
from typing import Sequence

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .aliases import AliasStore
from .config import DESTINATION_SHEET, FIRST_DATA_ROW, FIRST_DESTINATION_COLUMN, HEADER_ROW, SOURCE_SHEET
from .models import AliasEntry, CopyRecord, DestinationField, ImportResult, SourceField
from .utils import nonempty, text_as_excel_value

class WorkbookSession:
    def __init__(self, path: Path, aliases: AliasStore):
        self.path = Path(path)
        if self.path.suffix.lower() != ".xlsx":
            raise ValueError("Підтримуються лише файли .xlsx")
        self.workbook = load_workbook(self.path)
        if DESTINATION_SHEET not in self.workbook:
            raise ValueError(f"У підсумковому файлі немає аркуша «{DESTINATION_SHEET}»")
        self.sheet = self.workbook[DESTINATION_SHEET]
        header_fields = [DestinationField(cell.column, cell.value.strip())
                         for cell in self.sheet[HEADER_ROW]
                         if isinstance(cell.value, str) and cell.value.strip()]
        # Column A contains a row number, not a field to import or match.
        self.fields = [field for field in header_fields
                       if field.column >= FIRST_DESTINATION_COLUMN]
        if not self.fields:
            raise ValueError(
                f"У рядку {HEADER_ROW}, починаючи зі стовпчика B, немає текстових заголовків"
            )
        self.aliases = aliases
        self.template_id = aliases.template_id(self.fields)
        # Older releases included column A in the alias template fingerprint.
        # Migrate their saved mappings without ever restoring mappings to A.
        legacy_template_id = aliases.template_id(header_fields)
        if legacy_template_id != self.template_id:
            aliases.migrate_legacy_template(
                legacy_template_id, self.template_id, self.fields
            )
        self.current: ImportResult | None = None

    def restart_with_inputs(self, inputs: Sequence[Path]) -> "WorkbookSession":
        """Rebuild the whole import from disk with current persisted aliases.

        Work on a disposable workbook in the destination's directory, then
        atomically replace the real file only after *all* inputs succeed. This
        leaves both the original session and file untouched on any failure.
        It creates no lasting backup file.
        """
        staged_path = None
        try:
            with tempfile.NamedTemporaryFile(
                dir=self.path.parent, prefix=".excel-copy-restart-",
                suffix=".xlsx", delete=False,
            ) as stream:
                staged_path = Path(stream.name)
            shutil.copy2(self.path, staged_path)
            fresh_aliases = AliasStore(self.aliases.path)
            fresh = WorkbookSession(staged_path, fresh_aliases)
            fresh.clear_previous_imports()
            for index, filename in enumerate(inputs, start=1):
                filename = Path(filename)
                try:
                    fresh.open_input(filename)
                except Exception as exc:
                    raise ValueError(
                        f"Вхідний файл {index}/{len(inputs)} "
                        f"«{filename.name}»: {exc}"
                    ) from exc
            os.replace(staged_path, self.path)
            fresh.path = self.path
            return fresh
        finally:
            if staged_path is not None:
                staged_path.unlink(missing_ok=True)

    def clear_previous_imports(self) -> int:
        """Erase the writable output area when an existing destination is opened.

        Only columns represented by importable headers are writable, starting
        from FIRST_DATA_ROW. Leave ordinal column A, headers, styles, other
        sheets and unrelated columns alone. Persist immediately if needed.
        """
        modified = []
        for row in range(FIRST_DATA_ROW, self.sheet.max_row + 1):
            for field in self.fields:
                cell = self.sheet.cell(row, field.column)
                if cell.value is not None:
                    modified.append((cell, cell.value))
                    cell.value = None
        if modified:
            try:
                self._save()
            except Exception:
                # The on-disk file is unchanged if atomic save failed.
                for cell, previous_value in modified:
                    cell.value = previous_value
                raise
        self.current = None
        return len(modified)

    def next_empty_row(self) -> int:
        # An ordinal number in A does not make a row occupied. Still inspect
        # every other column to avoid overwriting partially populated records.
        for row in range(FIRST_DATA_ROW, self.sheet.max_row + 1):
            if all(not nonempty(self.sheet.cell(row, col).value)
                   for col in range(FIRST_DESTINATION_COLUMN, self.sheet.max_column + 1)):
                return row
        return max(FIRST_DATA_ROW, self.sheet.max_row + 1)

    def _read_source(self, path: Path) -> tuple[list[SourceField], list[str], list[str]]:
        if path.suffix.lower() != ".xlsx":
            raise ValueError("Підтримуються лише файли .xlsx")
        if path.resolve() == self.path.resolve():
            raise ValueError("Вхідний і підсумковий файли не можуть збігатися")
        cached = load_workbook(path, read_only=True, data_only=True)
        formulas = None
        try:
            # Read formula metadata separately: openpyxl does not calculate formulas.
            formulas = load_workbook(path, read_only=True, data_only=False)
            entries: list[SourceField] = []
            warnings: list[str] = []
            sheet_names = list(cached.sheetnames)
            for sheet_name in sheet_names:
                cached_ws = cached[sheet_name]
                formula_ws = formulas[sheet_name]
                cached_rows = cached_ws.iter_rows(
                    min_col=2, max_col=7, values_only=True,
                )
                formula_rows = formula_ws.iter_rows(min_col=2, max_col=7)
                for row, (cells, formula_cells) in enumerate(
                    zip(cached_rows, formula_rows), start=1
                ):
                    raw_name = cells[0]
                    name_valid = isinstance(raw_name, str) and bool(raw_name.strip())
                    name = raw_name.strip() if name_valid else (
                        "" if raw_name is None else str(raw_name)
                    )
                    values = tuple(cells[1:])
                    for col, (value, formula_cell) in enumerate(
                        zip(values, formula_cells[1:]), start=3
                    ):
                        if not nonempty(value) and formula_cell.data_type == "f":
                            warnings.append(
                                f"{sheet_name}!{get_column_letter(col)}{row}: "
                                "формула без збереженого результату — пропущено"
                            )
                    entries.append(SourceField(
                        row, name, values[0], sheet=sheet_name,
                        values=values, name_valid=name_valid,
                    ))
            return entries, warnings, sheet_names
        finally:
            cached.close()
            if formulas is not None:
                formulas.close()

    def _resolve_alias_destination(
        self, header: str | None, column: int | None
    ) -> DestinationField | None:
        """Resolve a saved target exactly as automatic importing does."""
        candidates = [field for field in self.fields if field.key == header]
        if len(candidates) == 1:
            return candidates[0]
        if len(candidates) > 1:
            return next((field for field in candidates if field.column == column), None)
        return None

    def _alias_target_for(self, source: SourceField) -> tuple[DestinationField | None, int]:
        """Resolve only an explicitly saved alias, never an incidental name match."""
        alias = self.aliases.get(self.template_id, source.name, source.sheet, row=source.row)
        if alias is None:
            return None, 3
        return (
            self._resolve_alias_destination(alias.get("header"), alias.get("column")),
            alias.get("source_column", 3),
        )

    def aliases_for_destination(self, column: int) -> list[AliasEntry]:
        """All persisted mappings to this output column, even if not copied now."""
        destination = next((field for field in self.fields if field.column == column), None)
        if destination is None:
            return []
        return [
            entry for entry in self.aliases.list_for_template(self.template_id)
            if self._resolve_alias_destination(
                entry.destination_key, entry.destination_column
            ) == destination
        ]

    def source_alias_targets(self) -> dict[tuple[str, int], str]:
        """Display each matching alias's target next to its source row."""
        targets = {}
        if self.current is None:
            return targets
        for source in self.current.all_sources:
            if not source.name_valid:
                continue
            alias = self.aliases.get(self.template_id, source.name, source.sheet, row=source.row)
            if alias is None:
                continue
            destination, _ = self._alias_target_for(source)
            targets[source.location] = (
                f"{destination.letter}: {destination.name}"
                if destination is not None else "Цільовий стовпчик не знайдено"
            )
        return targets

    def _write(self, row: int, records: list[CopyRecord]) -> None:
        if not records:
            return
        before = {}
        for record in records:
            column = record.destination.column
            if column < FIRST_DESTINATION_COLUMN:
                raise ValueError("Стовпчик A містить порядкові номери і не призначений для імпорту")
            cell = self.sheet.cell(row, column)
            if nonempty(cell.value):
                raise ValueError(f"Комірка {cell.coordinate} вже заповнена")
            before[column] = cell.value
            cell.value = text_as_excel_value(record.source.value)
        try:
            self._save()
        except Exception:
            for column, old_value in before.items():
                self.sheet.cell(row, column).value = old_value
            raise

    def _save(self) -> None:
        # The selected destination is written atomically in place.
        # Automatic backup copies are intentionally not created.
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(dir=self.path.parent, prefix=".excel-copy-",
                                             suffix=".xlsx", delete=False) as file:
                tmp_path = Path(file.name)
            self.workbook.save(tmp_path)
            os.chmod(tmp_path, self.path.stat().st_mode)
            os.replace(tmp_path, self.path)
        finally:
            if tmp_path is not None:
                tmp_path.unlink(missing_ok=True)

    def open_input(self, path: Path) -> ImportResult:
        source_fields, warnings, sheet_names = self._read_source(Path(path))
        target_row = self.next_empty_row()
        eligible_sources = [field for field in source_fields if field.eligible]
        # Only saved mappings are eligible. A row-scoped alias overrides the
        # general alias for exactly its row (while still checking its name).
        alias_issues: dict[tuple[str, int], str] = {}
        proposals: dict[int, list[CopyRecord]] = defaultdict(list)
        for source in source_fields:
            if not source.name_valid:
                continue
            alias = self.aliases.get(
                self.template_id, source.name, source.sheet, row=source.row,
            )
            if alias is None:
                continue
            destination, column = self._alias_target_for(source)
            if destination is None:
                alias_issues[source.location] = "Цільовий стовпчик аліаса не знайдено"
            elif type(column) is not int or not 3 <= column <= 7:
                alias_issues[source.location] = f"Недопустима колонка значення: {column}"
            elif not nonempty(source.value_at(column)):
                alias_issues[source.location] = (
                    f"У колонці {get_column_letter(column)} немає значення для копіювання"
                )
            else:
                specific = self.aliases.get_exact(
                    self.template_id, source.name, source.sheet, row=source.row,
                ) is not None
                proposals[destination.column].append(CopyRecord(
                    source.using_column(column), destination, "Аліас",
                    alias_row=source.row if specific else None,
                ))

        safe_matches = []
        had_duplicates = False
        had_conflicts = False
        for matches in proposals.values():
            if len(matches) == 1:
                safe_matches.append(matches[0])
                continue
            # One row-specific alias wins over any number of general aliases
            # with the same normalized source name targeting this column.
            scoped = [record for record in matches if record.alias_row is not None]
            if (len(scoped) == 1
                    and all(record.source.key == scoped[0].source.key
                            for record in matches)):
                winner = scoped[0]
                safe_matches.append(winner)
                for record in matches:
                    if record is not winner:
                        alias_issues[record.source.location] = (
                            f"Пріоритет має аліас для рядка {winner.source.row}"
                        )
                continue
            # A general alias pointing to several identically named rows
            # is ambiguous unless the single explicit-row alias above resolves it.
            same_sheet_and_name = len({
                (record.source.sheet.casefold(), record.source.key)
                for record in matches
            }) == 1
            had_duplicates |= same_sheet_and_name
            had_conflicts |= not same_sheet_and_name
            reason = (
                "Декілька рядків із такою назвою на аркуші мають значення "
                "у колонці аліаса"
                if same_sheet_and_name else
                "Кілька параметрів із аліасами претендують на цей вихідний стовпчик"
            )
            for record in matches:
                alias_issues[record.source.location] = reason

        safe_matches.sort(key=lambda rec: rec.destination.column)
        self._write(target_row, safe_matches)
        result = ImportResult(
            filename=Path(path).name, row=target_row,
            copied=safe_matches, warnings=warnings,
            sheets=sheet_names, all_sources=source_fields,
            alias_issues=alias_issues,
        )
        if had_duplicates:
            result.warnings.append(
                "Повторювані назви в межах одного аркуша потребують ручного вибору"
            )
        if had_conflicts:
            result.warnings.append(
                "Кілька вхідних параметрів претендують на один стовпчик — пропущено"
            )
        for source in source_fields:
            reason = alias_issues.get(source.location)
            if reason:
                result.warnings.append(
                    f"{source.sheet}!B{source.row} «{source.name}»: {reason}"
                )
        used_rows = {record.source.location for record in safe_matches}
        used_columns = {record.destination.column for record in safe_matches}
        result.unmatched_sources = [source for source in eligible_sources
                                    if source.location not in used_rows]
        result.unmatched_destinations = [dest for dest in self.fields if dest.column not in used_columns]
        self.current = result
        return result

    def copy_manually(
        self, source_row: int, destination_col: int, sheet: str = SOURCE_SHEET,
        source_column: int = 3, alias_row: int | None = None,
    ) -> str | None:
        result = self.current
        if result is None:
            raise ValueError("Спочатку відкрийте вхідний файл")
        source = next((s for s in result.unmatched_sources
                       if s.row == source_row and s.sheet.casefold() == sheet.casefold()), None)
        destination = next((d for d in result.unmatched_destinations if d.column == destination_col), None)
        if source is None or destination is None:
            raise ValueError("Вибраний параметр уже був скопійований")
        if not nonempty(source.value_at(source_column)):
            raise ValueError(
                f"У комірці {sheet}!{get_column_letter(source_column)}{source_row} немає значення"
            )
        if alias_row is not None and alias_row != source.row:
            raise ValueError("Номер рядка аліаса має збігатися з вибраним рядком")
        record = CopyRecord(
            source.using_column(source_column), destination, "Вручну",
            alias_row=alias_row,
        )
        self._write(result.row, [record])
        result.copied.append(record)
        result.unmatched_sources.remove(source)
        result.unmatched_destinations.remove(destination)
        try:
            self.aliases.set(self.template_id, record.source, destination, row=alias_row)
        except (OSError, ValueError) as exc:
            return f"Значення скопійовано, але аліас не збережено: {exc}"
        return None
