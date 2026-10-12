"""ExcelCopyScript: interactive transfer from ВИСНОВКИ to Дані_підрозділів."""

from __future__ import annotations

from dataclasses import dataclass, field, replace

from openpyxl.utils import get_column_letter

from .config import SOURCE_SHEET
from .utils import normalize

@dataclass(frozen=True)
class DestinationField:
    column: int
    name: str

    @property
    def key(self) -> str:
        return normalize(self.name)

    @property
    def letter(self) -> str:
        return get_column_letter(self.column)


@dataclass(frozen=True)
class SourceField:
    row: int
    name: str
    value: object
    sheet: str = SOURCE_SHEET
    values: tuple[object, ...] = ()
    name_valid: bool = True
    value_column: int = 3
    value_first_column: int = 3

    def value_at(self, column: int) -> object:
        if self.values:
            offset = column - self.value_first_column
            if not 0 <= offset < len(self.values):
                raise ValueError("Колонка значення поза діапазоном джерела")
            return self.values[offset]
        # Preserve older programmatic uses without an explicit values tuple.
        if column != self.value_first_column:
            return None
        return self.value

    def using_column(self, column: int) -> "SourceField":
        return replace(self, value=self.value_at(column), value_column=column)

    @property
    def eligible(self) -> bool:
        from .utils import nonempty
        return self.name_valid and bool(self.name.strip()) and any(
            nonempty(value) for value in (self.values or (self.value,))
        )

    @property
    def key(self) -> str:
        return normalize(self.name)

    @property
    def location(self) -> tuple[str, int]:
        return normalize(self.sheet), self.row



@dataclass(frozen=True)
class AliasEntry:
    """Displayable alias mapping, including aliases created by older versions."""

    source_key: str
    source_name: str
    destination_key: str
    destination_column: int | None
    sheet: str = SOURCE_SHEET
    source_column: int = 3
    source_row: int | None = None

@dataclass(frozen=True)
class CopyRecord:
    source: SourceField
    destination: DestinationField
    method: str
    alias_row: int | None = None


@dataclass
class ImportResult:
    filename: str
    row: int
    copied: list[CopyRecord] = field(default_factory=list)
    unmatched_destinations: list[DestinationField] = field(default_factory=list)
    unmatched_sources: list[SourceField] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    sheets: list[str] = field(default_factory=list)
    all_sources: list[SourceField] = field(default_factory=list)
    alias_issues: dict[tuple[str, int], str] = field(default_factory=dict)
    overwrite_existing: bool = False


@dataclass(frozen=True)
class RowMatch:
    """An existing destination row sharing one or more imported values."""

    row: int
    matching: tuple[str, ...]
    differing: tuple[str, ...]
    total: int

    @property
    def complete(self) -> bool:
        return len(self.matching) == self.total


@dataclass(frozen=True)
class RowSelectionRequest:
    """Ambiguous destination-row matches requiring an explicit user choice."""

    filename: str
    candidates: tuple[RowMatch, ...]
