"""Multi-sheet import, alias scoping and v1-to-v2 migration regression tests."""

import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

from openpyxl import Workbook, load_workbook

from excel_copy_script.aliases import AliasStore
from excel_copy_script.config import SOURCE_SHEET
from excel_copy_script.gui import ExcelCopyApp
from excel_copy_script.models import DestinationField, SourceField
from excel_copy_script.workbook import WorkbookSession


class MultiSheetTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.summary = self.root / "summary.xlsx"
        self.alias_file = self.root / "config" / "aliases.json"
        book = Workbook()
        ws = book.active
        ws.title = "Дані_підрозділів"
        ws["A3"], ws["B3"], ws["C3"], ws["D3"] = "№", "Кількість", "Техніка", "Резерв"
        book.save(self.summary)
        self.aliases = AliasStore(self.alias_file)
        self.session = WorkbookSession(self.summary, self.aliases)

    def make_input(self, filename, sheets):
        book = Workbook()
        primary = book.active
        primary.title = SOURCE_SHEET
        for sheet_name, records in sheets.items():
            ws = primary if sheet_name == SOURCE_SHEET else book.create_sheet(sheet_name)
            for row, name, value in records:
                ws.cell(row, 2, name)
                ws.cell(row, 3, value)
        path = self.root / filename
        book.save(path)
        return path

    def test_source_sheet_list_includes_empty_worksheets(self):
        path = self.make_input("empty_sheet.xlsx", {
            SOURCE_SHEET: [(1, "Техніка", 5)],
            "Без даних": [],
            "Склад": [(2, "Неоднозначне", 12)],
        })
        result = self.session.open_input(path)
        self.assertEqual(result.sheets, [SOURCE_SHEET, "Без даних", "Склад"])
        self.assertEqual([s.sheet for s in result.unmatched_sources], ["Склад"])

    def test_only_summary_sheet_uses_exact_header_match(self):
        filename = self.make_input("exact.xlsx", {
            SOURCE_SHEET: [(1, "Техніка", 5)],
            "Склад": [(1, "Резерв", 42), (2, "Кількість", 13)],
        })
        result = self.session.open_input(filename)
        self.assertEqual([(c.source.sheet, c.destination.column) for c in result.copied],
                         [(SOURCE_SHEET, 3)])
        self.assertEqual(self.session.sheet["C4"].value, 5)
        self.assertIsNone(self.session.sheet["D4"].value)
        self.assertEqual({s.sheet for s in result.unmatched_sources}, {"Склад"})
        self.assertEqual({s.name for s in result.unmatched_sources}, {"Резерв", "Кількість"})

    def test_same_parameter_on_different_sheets_is_independent(self):
        first = self.make_input("first.xlsx", {
            SOURCE_SHEET: [(1, "Техніка", 5)],
            "Склад": [(1, "Техніка", 70)],
        })
        result = self.session.open_input(first)
        self.assertEqual(len(result.copied), 1)
        self.assertEqual(result.unmatched_sources[0].sheet, "Склад")
        self.session.copy_manually(1, 4, sheet="Склад")
        self.assertIsNone(self.aliases.get(self.session.template_id, "Техніка", "Інший"))
        mapping = self.aliases.get(self.session.template_id, "Техніка", "Склад")
        self.assertEqual(mapping["column"], 4)

        session = WorkbookSession(self.summary, AliasStore(self.alias_file))
        second = self.make_input("second.xlsx", {
            SOURCE_SHEET: [(2, "Техніка", 10)],
            "Склад": [(2, "Техніка", 80)],
        })
        result = session.open_input(second)
        self.assertEqual(
            {(r.source.sheet, r.destination.column, r.method) for r in result.copied},
            {(SOURCE_SHEET, 3, "Збіг назви"), ("Склад", 4, "Аліас")},
        )
        self.assertEqual(session.sheet["C5"].value, 10)
        self.assertEqual(session.sheet["D5"].value, 80)

    def test_same_name_and_row_on_different_sheets_can_be_selected_manually(self):
        source = self.make_input("repeated.xlsx", {
            SOURCE_SHEET: [(3, "Чисельність", 11)],
            "Склад": [(3, "Чисельність", 22)],
        })
        result = self.session.open_input(source)
        self.assertEqual(len(result.unmatched_sources), 2)
        self.session.copy_manually(3, 2, sheet=SOURCE_SHEET)
        self.session.copy_manually(3, 4, sheet="Склад")
        entries = self.aliases.list_for_template(self.session.template_id)
        self.assertEqual([(a.sheet, a.destination_column) for a in entries],
                         [("ВИСНОВКИ", 2), ("Склад", 4)])
        self.assertEqual(self.session.sheet["B4"].value, 11)
        self.assertEqual(self.session.sheet["D4"].value, 22)

    def test_delete_one_sheet_alias_keeps_other(self):
        source = self.make_input("names.xlsx", {
            SOURCE_SHEET: [(1, "Чисельність", 1)], "Склад": [(1, "Чисельність", 2)]
        })
        self.session.open_input(source)
        self.session.copy_manually(1, 2, sheet=SOURCE_SHEET)
        self.session.copy_manually(1, 4, sheet="Склад")
        template = self.session.template_id
        self.assertTrue(self.aliases.delete(template, "Чисельність", sheet="Склад"))
        self.assertFalse(self.aliases.delete(template, "Чисельність", sheet="Склад"))
        self.assertIsNotNone(self.aliases.get(template, "Чисельність", SOURCE_SHEET))
        self.assertIsNone(AliasStore(self.alias_file).get(template, "Чисельність", "Склад"))

    def test_duplicate_field_in_one_sheet_does_not_block_other_sheet(self):
        source = self.make_input("dups.xlsx", {
            SOURCE_SHEET: [(1, "Техніка", 1)],
            "Склад": [(1, "Дублікат", 4), (2, "Дублікат", 5)],
        })
        self.aliases.set(self.session.template_id, SourceField(1, "Дублікат", 4, "Склад"),
                         self.session.fields[2])
        result = self.session.open_input(source)
        self.assertEqual([(r.source.sheet, r.source.name) for r in result.copied],
                         [(SOURCE_SHEET, "Техніка")])
        self.assertEqual(len(result.unmatched_sources), 2)

    def test_two_sheets_targeting_same_column_do_not_overwrite(self):
        source = self.make_input("collision.xlsx", {
            SOURCE_SHEET: [(1, "Техніка", 1)],
            "Склад": [(1, "Списана техніка", 2)],
        })
        self.aliases.set(
            self.session.template_id,
            SourceField(1, "Списана техніка", 2, "Склад"),
            self.session.fields[1],
        )
        result = self.session.open_input(source)
        self.assertFalse(result.copied)
        self.assertEqual(len(result.unmatched_sources), 2)
        self.assertTrue(any("претендують" in message for message in result.warnings))
        self.assertIsNone(self.session.sheet["C4"].value)

    def test_gui_passes_sheet_with_selected_row(self):
        session = Mock()
        app = SimpleNamespace(
            session=session,
            source_tree=Mock(selection=Mock(return_value=("I22",))),
            destination_tree=Mock(selection=Mock(return_value=("4",))),
            _unmatched_source_ids={"I22": ("Склад", 3)},
            _refresh_tables=Mock(),
        )
        session.copy_manually.return_value = None
        ExcelCopyApp.copy_selected(app)
        session.copy_manually.assert_called_once_with(3, 4, sheet="Склад")



class SheetSelectorTests(unittest.TestCase):
    """Test sheet filtering without requiring an X server."""

    class FakeVar:
        def __init__(self):
            self.value = ""

        def set(self, value):
            self.value = value

        def get(self):
            return self.value

    class FakeTree:
        def __init__(self):
            self.rows = {}
            self.selected = ()
            self.counter = 0

        def get_children(self):
            return list(self.rows)

        def delete(self, *items):
            for item in items:
                self.rows.pop(item, None)
            self.selected = ()

        def insert(self, _parent, _position, values):
            self.counter += 1
            item_id = f"row{self.counter}"
            self.rows[item_id] = tuple(values)
            return item_id

        def selection(self):
            return self.selected

    def make_ui(self, result):
        app = SimpleNamespace(
            session=SimpleNamespace(current=result),
            _unmatched_source_ids={},
            selected_source_sheet=self.FakeVar(),
            source_sheet_combo=Mock(),
            source_tree=self.FakeTree(),
            _on_selection=Mock(),
        )
        app._refresh_source_rows = lambda: ExcelCopyApp._refresh_source_rows(app)
        return app

    def test_filter_and_sheet_change_only_refreshes_source_rows(self):
        from excel_copy_script.models import ImportResult

        result = ImportResult(filename="file.xlsx", row=4, sheets=[
            SOURCE_SHEET, "Склад", "Порожній",
        ], unmatched_sources=[
            SourceField(2, "Перший", 11, sheet=SOURCE_SHEET),
            SourceField(2, "Другий", 22, sheet="Склад"),
            SourceField(3, "Третій", 33, sheet="Склад"),
        ])
        app = self.make_ui(result)
        ExcelCopyApp._set_source_sheets(app, result.sheets)
        app.source_sheet_combo.configure.assert_called_with(
            values=[SOURCE_SHEET, "Склад", "Порожній"], state="readonly",
        )
        self.assertEqual(app.selected_source_sheet.get(), SOURCE_SHEET)
        ExcelCopyApp._refresh_source_rows(app)
        self.assertEqual(list(app.source_tree.rows.values()), [(2, "Перший", "11")])
        app.selected_source_sheet.set("Склад")
        ExcelCopyApp._on_source_sheet_changed(app)
        self.assertEqual(list(app.source_tree.rows.values()), [
            (2, "Другий", "22"), (3, "Третій", "33"),
        ])
        self.assertEqual(set(app._unmatched_source_ids.values()), {
            ("Склад", 2), ("Склад", 3),
        })
        self.assertEqual(app.source_tree.selection(), ())
        app.selected_source_sheet.set("Порожній")
        ExcelCopyApp._on_source_sheet_changed(app)
        self.assertEqual(app.source_tree.rows, {})
        self.assertEqual(app._unmatched_source_ids, {})
        self.assertGreaterEqual(app._on_selection.call_count, 3)

    def test_manual_copy_uses_correct_sheet_after_selection(self):
        from excel_copy_script.models import ImportResult

        result = ImportResult(filename="file.xlsx", row=4, sheets=[
            SOURCE_SHEET, "Склад",
        ], unmatched_sources=[
            SourceField(2, "Спільне", 1, sheet=SOURCE_SHEET),
            SourceField(2, "Спільне", 2, sheet="Склад"),
        ])
        app = self.make_ui(result)
        session = Mock()
        session.current = result
        session.copy_manually.return_value = None
        app.session = session
        app.destination_tree = Mock()
        app.destination_tree.selection.return_value = ("3",)
        app._refresh_tables = Mock()
        ExcelCopyApp._set_source_sheets(app, result.sheets)
        app.selected_source_sheet.set("Склад")
        ExcelCopyApp._refresh_source_rows(app)
        app.source_tree.selected = (next(iter(app.source_tree.rows)),)
        ExcelCopyApp.copy_selected(app)
        session.copy_manually.assert_called_once_with(2, 3, sheet="Склад")
        app.selected_source_sheet.set(SOURCE_SHEET)
        ExcelCopyApp._refresh_source_rows(app)
        app.source_tree.selected = (next(iter(app.source_tree.rows)),)
        ExcelCopyApp.copy_selected(app)
        session.copy_manually.assert_called_with(2, 3, sheet=SOURCE_SHEET)

    def test_selector_reset_for_another_workbook(self):
        result = SimpleNamespace(sheets=[SOURCE_SHEET, "Склад"])
        app = self.make_ui(result)
        ExcelCopyApp._set_source_sheets(app, result.sheets)
        app.selected_source_sheet.set("Склад")
        ExcelCopyApp._set_source_sheets(app, [SOURCE_SHEET, "Поточний"])
        self.assertEqual(app.selected_source_sheet.get(), SOURCE_SHEET)
        ExcelCopyApp._set_source_sheets(app, [])
        self.assertEqual(app.selected_source_sheet.get(), "")
        app.source_sheet_combo.configure.assert_called_with(
            values=[], state="disabled",
        )


class AliasMigrationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        self.path = self.directory / "aliases.json"

    def test_v1_aliases_migrate_into_summary_and_backup_created(self):
        original = {
            "version": 1,
            "templates": {
                "template-1": {
                    "кількість": {"header": "особовий склад", "column": 2,
                                  "source_name": "Кількість"},
                },
                "template-2": {
                    "інша назва": {"header": "техніка", "column": 3},
                },
            },
        }
        encoded = json.dumps(original, ensure_ascii=False)
        self.path.write_text(encoded, encoding="utf-8")
        store = AliasStore(self.path)
        self.assertEqual(store.data["version"], 2)
        self.assertEqual(store.get("template-1", "Кількість", SOURCE_SHEET)["column"], 2)
        self.assertIsNone(store.get("template-1", "Кількість", "Склад"))
        self.assertEqual(store.get("template-2", "Інша назва", SOURCE_SHEET)["column"], 3)
        self.assertEqual(store.list_for_template("template-1")[0].sheet, SOURCE_SHEET)
        backup = self.directory / "aliases.json.v1.bak"
        self.assertEqual(backup.read_text(encoding="utf-8"), encoded)
        self.assertEqual(json.loads(self.path.read_text(encoding="utf-8"))["version"], 2)
        AliasStore(self.path)
        self.assertEqual(backup.read_text(encoding="utf-8"), encoded)

    def test_legacy_column_a_template_migrates_and_cannot_resurrect_deleted(self):
        target = self.directory / "output.xlsx"
        book = Workbook()
        ws = book.active
        ws.title = "Дані_підрозділів"
        ws["A3"], ws["B3"], ws["C3"] = "№", "Особовий склад", "Техніка"
        book.save(target)
        fields = [DestinationField(1, "№"), DestinationField(2, "Особовий склад"),
                  DestinationField(3, "Техніка")]
        old_id = AliasStore.template_id(fields)
        current_id = AliasStore.template_id(fields[1:])
        data = {"version": 1, "templates": {old_id: {
            "нове ім’я": {"header": "техніка", "column": 3},
            "номер": {"header": "№", "column": 1},
        }}}
        self.path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        initial = AliasStore(self.path)
        session = WorkbookSession(target, initial)
        self.assertIsNotNone(initial.get(current_id, "Нове ім’я", SOURCE_SHEET))
        self.assertIsNone(initial.get(current_id, "Номер", SOURCE_SHEET))
        self.assertNotIn(old_id, initial.data["templates"])
        self.assertTrue(initial.delete(current_id, "Нове ім’я", SOURCE_SHEET))
        restarted = WorkbookSession(target, AliasStore(self.path))
        self.assertIsNone(restarted.aliases.get(current_id, "Нове ім’я", SOURCE_SHEET))

    def test_invalid_schema_is_rejected_without_overwriting(self):
        text = '{"version": 3, "templates": {}}'
        self.path.write_text(text, encoding="utf-8")
        with self.assertRaises(ValueError):
            AliasStore(self.path)
        self.assertEqual(self.path.read_text(encoding="utf-8"), text)
        self.assertFalse((self.directory / "aliases.json.v1.bak").exists())


if __name__ == "__main__":
    unittest.main()
