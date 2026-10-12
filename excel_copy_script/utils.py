"""ExcelCopyScript: interactive transfer from ВИСНОВКИ to Дані_підрозділів."""

from __future__ import annotations

import unicodedata

def normalize(value: str) -> str:
    """Ignore casing, Unicode presentation variants and repeated whitespace."""
    return " ".join(unicodedata.normalize("NFKC", value).casefold().split())


def present(value: object) -> str:
    return "" if value is None else str(value).replace("\n", " ")


def nonempty(value: object) -> bool:
    return value is not None and (not isinstance(value, str) or bool(value.strip()))


def text_as_excel_value(value: object) -> object:
    """Never turn text copied from a source cell into a new Excel formula."""
    if isinstance(value, str) and value.startswith("="):
        return "'" + value
    return value
