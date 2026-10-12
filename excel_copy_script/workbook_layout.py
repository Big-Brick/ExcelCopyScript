"""Validated, editable workbook structure stored inside each profile."""

from __future__ import annotations

from dataclasses import asdict, dataclass

from openpyxl.utils import get_column_letter


@dataclass(frozen=True)
class WorkbookLayout:
    destination_sheet: str = "Дані_підрозділів"
    header_row: int = 3
    first_data_row: int = 4
    first_destination_column: int = 2
    source_name_column: int = 2
    first_source_row: int = 1
    source_value_first_column: int = 3
    source_value_last_column: int = 7
    preferred_source_sheet: str = "ВИСНОВКИ"

    def __post_init__(self):
        if not isinstance(self.destination_sheet, str) or not self.destination_sheet.strip():
            raise ValueError("Назва вихідного аркуша не може бути порожньою")
        if not isinstance(self.preferred_source_sheet, str) or not self.preferred_source_sheet.strip():
            raise ValueError("Назва типового вхідного аркуша не може бути порожньою")
        for key, value in asdict(self).items():
            if key in ("destination_sheet", "preferred_source_sheet"):
                continue
            if type(value) is not int or value < 1:
                raise ValueError(f"Налаштування {key} має бути додатним цілим числом")
        if self.first_data_row <= self.header_row:
            raise ValueError("Перший рядок даних має бути після рядка заголовків")
        if self.source_value_first_column > self.source_value_last_column:
            raise ValueError("Неправильний діапазон вхідних стовпчиків значень")
        if self.source_value_first_column <= self.source_name_column <= self.source_value_last_column:
            raise ValueError("Колонка назви не може одночасно містити значення")
        if self.source_value_last_column > 16384 or self.first_destination_column > 16384:
            raise ValueError("Номер стовпчика перевищує ліміт Excel")

    @property
    def source_columns(self) -> tuple[int, ...]:
        return tuple(range(self.source_value_first_column, self.source_value_last_column + 1))

    @property
    def source_letters(self) -> tuple[str, ...]:
        return tuple(get_column_letter(column) for column in self.source_columns)

    @classmethod
    def from_dict(cls, data: dict) -> "WorkbookLayout":
        if not isinstance(data, dict):
            raise ValueError("Секція layout профілю має бути об’єктом")
        expected = set(cls.__dataclass_fields__)
        if set(data) != expected:
            missing = sorted(expected - set(data))
            extras = sorted(set(data) - expected)
            raise ValueError(f"Поля layout: відсутні {missing}; невідомі {extras}")
        return cls(**data)

    def to_dict(self) -> dict:
        return asdict(self)
