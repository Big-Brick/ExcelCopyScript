"""ExcelCopyScript: interactive transfer from ВИСНОВКИ to Дані_підрозділів."""

from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl.utils import get_column_letter

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

    @property
    def key(self) -> str:
        return normalize(self.name)



@dataclass(frozen=True)
class AliasEntry:
    """Displayable alias mapping, including aliases created by older versions."""

    source_key: str
    source_name: str
    destination_key: str
    destination_column: int | None

@dataclass(frozen=True)
class CopyRecord:
    source: SourceField
    destination: DestinationField
    method: str


@dataclass
class ImportResult:
    filename: str
    row: int
    copied: list[CopyRecord] = field(default_factory=list)
    unmatched_destinations: list[DestinationField] = field(default_factory=list)
    unmatched_sources: list[SourceField] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
