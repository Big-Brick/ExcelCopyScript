"""Excel workbook layout and file selection configuration."""

DESTINATION_SHEET = "Дані_підрозділів"
SOURCE_SHEET = "ВИСНОВКИ"
HEADER_ROW = 3
FIRST_DATA_ROW = 4
FIRST_DESTINATION_COLUMN = 2  # Column A is an ordinal number; never import into it.
FILE_TYPES = [("Книги Excel", "*.xlsx")]
