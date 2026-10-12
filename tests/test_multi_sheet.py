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
        self.assertEqual([s.sheet for s in result.unmatched_sources], [SOURCE_SHEET, "Склад"])

    def test_same_headers_on_any_sheet_never_copy_without_alias(self):
        filename = self.make_input("exact.xlsx", {
            SOURCE_SHEET: [(1, "Техніка", 5)],
            "Склад": [(1, "Резерв", 42), (2, "Кількість", 13)],
        })
        result = self.session.open_input(filename)
        self.assertEqual(result.copied, [])
        self.assertIsNone(self.session.sheet["C4"].value)
        self.assertIsNone(self.session.sheet["D4"].value)
        self.assertEqual({s.sheet for s in result.unmatched_sources}, {SOURCE_SHEET, "Склад"})
        self.assertEqual({s.name for s in result.unmatched_sources},
                         {"Техніка", "Резерв", "Кількість"})

    def test_workbook_without_summary_sheet_uses_aliases_only(self):
        input_file = self.root / "other_sheet_only.xlsx"
        book = Workbook()
        ws = book.active
        ws.title = "Склад"
        ws["B2"], ws["C2"] = "Техніка", 27
        book.save(input_file)

        first = self.session.open_input(input_file)
        self.assertEqual(first.sheets, ["Склад"])
        self.assertFalse(first.copied)
        self.assertEqual([s.name for s in first.unmatched_sources], ["Техніка"])

        self.session.copy_manually(2, 3, sheet="Склад")
        self.assertEqual(self.session.sheet["C4"].value, 27)
        reopened = WorkbookSession(self.summary, AliasStore(self.alias_file))
        later = reopened.open_input(input_file)
        self.assertEqual([(rec.method, rec.source.sheet) for rec in later.copied],
                         [("Аліас", "Склад")])
        self.assertEqual(reopened.sheet["C4"].value, 27)
        self.assertEqual(later.row, 4)

    def test_same_parameter_on_different_sheets_is_independent(self):
        first = self.make_input("first.xlsx", {
            SOURCE_SHEET: [(1, "Техніка", 5)],
            "Склад": [(1, "Техніка", 70)],
        })
        result = self.session.open_input(first)
        self.assertEqual(len(result.copied), 0)
        self.assertEqual({source.sheet for source in result.unmatched_sources},
                         {SOURCE_SHEET, "Склад"})
        self.session.copy_manually(1, 3, sheet=SOURCE_SHEET)
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
            {(SOURCE_SHEET, 3, "Аліас"), ("Склад", 4, "Аліас")},
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
        self.aliases.set(self.session.template_id, SourceField(1, "Техніка", 1),
                         self.session.fields[1])
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
            SourceField(1, "Техніка", 1, SOURCE_SHEET),
            self.session.fields[1],
        )
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

    def test_multiple_aliases_per_output_column_are_listed(self):
        """Destination context menu lists saved aliases absent from the current file."""
        self.aliases.set(
            self.session.template_id,
            SourceField(1, "Основний параметр", 7, sheet=SOURCE_SHEET),
            self.session.fields[1],
        )
        self.aliases.set(
            self.session.template_id,
            SourceField(2, "Запасна назва", 8, sheet="Склад"),
            self.session.fields[1],
        )
        entries = self.session.aliases_for_destination(3)
        self.assertEqual({(e.sheet, e.source_name) for e in entries}, {
            (SOURCE_SHEET, "Основний параметр"), ("Склад", "Запасна назва"),
        })
        self.assertEqual(self.session.aliases_for_destination(1), [])
        self.assertEqual(self.session.aliases_for_destination(2), [])

    def test_ambiguous_output_titles_use_pinned_alias_column(self):
        self.session.sheet["D3"] = "Техніка"
        self.session.fields = [
            DestinationField(2, "Кількість"),
            DestinationField(3, "Техніка"),
            DestinationField(4, "Техніка"),
        ]
        self.session.template_id = self.aliases.template_id(self.session.fields)
        self.aliases.set(
            self.session.template_id, SourceField(1, "Заготовка", 10, sheet="Склад"),
            self.session.fields[2],
        )
        self.assertEqual(self.session.aliases_for_destination(3), [])
        self.assertEqual(
            [e.source_name for e in self.session.aliases_for_destination(4)],
            ["Заготовка"],
        )

    def test_gui_passes_sheet_with_selected_row(self):
        session = Mock()
        source = SourceField(3, "Показник", 42, sheet="Склад")
        session.current.unmatched_sources = [source]
        session.current.unmatched_destinations = [DestinationField(4, "Резерв")]
        app = SimpleNamespace(
            session=session,
            row_only_var=Mock(get=Mock(return_value=False)),
            create_sources=Mock(
                selected_source=Mock(return_value=source),
                source_column=Mock(return_value=3),
            ),
            destination_tree=Mock(selection=Mock(return_value=("4",))),
            _refresh_tables=Mock(),
        )
        session.copy_manually.return_value = None
        ExcelCopyApp.copy_selected(app)
        session.copy_manually.assert_called_once_with(
            3, 4, sheet="Склад", source_column=3, alias_row=None,
        )



class SheetSelectorTests(unittest.TestCase):
    """Exercise two independent source panels without requiring an X server."""

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
            self.tags = {}

        def get_children(self):
            return tuple(self.rows)

        def delete(self, *items):
            for item in items:
                self.rows.pop(item, None)
            self.selected = ()

        def insert(self, _parent, _position, values, iid=None, tags=()):
            self.counter += 1
            item_id = iid or f"row{self.counter}"
            self.rows[item_id] = tuple(values)
            self.tags[item_id] = tuple(tags)
            return item_id

        def selection(self):
            return self.selected

    def make_panel(self):
        from excel_copy_script.source_rows_panel import SourceRowsPanel

        panel = SimpleNamespace(
            _rows=[],
            _sources_by_item_id={},
            _alias_targets={},
            _copied_locations=set(),
            _alias_issues={},
            issue_var=self.FakeVar(),
            _on_selection_changed=Mock(),
            selected_sheet=self.FakeVar(),
            selected_value_column=self.FakeVar(),
            sheet_combo=Mock(),
            tree=self.FakeTree(),
            values_tree=self.FakeTree(),
            COLUMNS=tuple(range(3, 8)),
            layout=__import__("excel_copy_script.workbook_layout", fromlist=["WorkbookLayout"]).WorkbookLayout(),
        )
        panel.refresh = lambda rows=None, alias_targets=None, copied_locations=None, alias_issues=None: SourceRowsPanel.refresh(
            panel, rows, alias_targets=alias_targets,
            copied_locations=copied_locations, alias_issues=alias_issues,
        )
        panel.selected_source = lambda: SourceRowsPanel.selected_source(panel)
        panel.set_sheets = lambda sheets: SourceRowsPanel.set_sheets(panel, sheets)
        panel._notify_selection = lambda: SourceRowsPanel._notify_selection(panel)
        return panel

    def test_each_tab_filters_the_same_rows_independently(self):
        from excel_copy_script.source_rows_panel import SourceRowsPanel

        fields = [
            SourceField(2, "Перший", 11, sheet=SOURCE_SHEET),
            SourceField(2, "Другий", 22, sheet="Склад"),
            SourceField(3, "Третій", 33, sheet="Склад"),
        ]
        review, create = self.make_panel(), self.make_panel()
        for panel in (review, create):
            SourceRowsPanel.set_sheets(panel, [SOURCE_SHEET, "Склад", "Порожній"])
            panel.refresh(fields)
            self.assertEqual(panel.selected_sheet.get(), SOURCE_SHEET)
            self.assertEqual(list(panel.tree.rows.values()), [("", 2, "Перший")])
            self.assertEqual(next(iter(panel.values_tree.rows.values()))[0], "11")
            panel.sheet_combo.configure.assert_called_with(
                values=[SOURCE_SHEET, "Склад", "Порожній"], state="readonly",
            )

        review.selected_sheet.set("Склад")
        SourceRowsPanel._on_sheet_changed(review)
        self.assertEqual(list(review.tree.rows.values()), [
            ("", 2, "Другий"), ("", 3, "Третій"),
        ])
        self.assertEqual(list(create.tree.rows.values()), [("", 2, "Перший")])

        review.selected_sheet.set("Порожній")
        SourceRowsPanel._on_sheet_changed(review)
        self.assertEqual(review.tree.rows, {})
        self.assertEqual(review._sources_by_item_id, {})
        self.assertGreaterEqual(review._on_selection_changed.call_count, 3)

    def test_all_rows_color_and_alias_target_in_frozen_column(self):
        from excel_copy_script.source_rows_panel import SourceRowsPanel

        panel = self.make_panel()
        SourceRowsPanel.set_sheets(panel, [SOURCE_SHEET])
        fields = [
            SourceField(1, "", None, sheet=SOURCE_SHEET, name_valid=False),
            SourceField(2, "Наявний аліас", None, sheet=SOURCE_SHEET,
                        values=(None, None, 42, None, None)),
            SourceField(3, "Звичайний", 17, sheet=SOURCE_SHEET),
        ]
        panel.refresh(
            fields, alias_targets={(SOURCE_SHEET.casefold(), 2): "D: Резерв"},
            copied_locations={(SOURCE_SHEET.casefold(), 2)},
        )
        self.assertEqual(list(panel.tree.rows.values()), [
            ("", 1, ""), ("D: Резерв", 2, "Наявний аліас"),
            ("", 3, "Звичайний"),
        ])
        self.assertEqual(list(panel.tree.tags.values()), [
            ("unusable",), ("aliased",), (),
        ])
        self.assertEqual(panel.values_tree.rows["row-2"], ("", "", "42", "", ""))
        self.assertEqual(panel.selected_sheet.get(), SOURCE_SHEET)

    def test_selection_keeps_sheet_identity_and_resets_when_filtered(self):
        from excel_copy_script.source_rows_panel import SourceRowsPanel

        panel = self.make_panel()
        SourceRowsPanel.set_sheets(panel, [SOURCE_SHEET, "Склад"])
        panel.refresh([
            SourceField(2, "Спільне", 1, sheet=SOURCE_SHEET),
            SourceField(2, "Спільне", 2, sheet="Склад"),
        ])
        panel.selected_sheet.set("Склад")
        SourceRowsPanel._on_sheet_changed(panel)
        panel.tree.selected = (next(iter(panel.tree.rows)),)
        selected = SourceRowsPanel.selected_source(panel)
        self.assertEqual((selected.sheet, selected.row), ("Склад", 2))
        panel.refresh([])
        self.assertIsNone(SourceRowsPanel.selected_source(panel))
        self.assertEqual(panel.selected_sheet.get(), "Склад")

    def test_both_selectors_reset_when_next_file_opens(self):
        from excel_copy_script.source_rows_panel import SourceRowsPanel

        review, create = self.make_panel(), self.make_panel()
        app = SimpleNamespace(review_sources=review, create_sources=create)
        ExcelCopyApp._set_source_sheets(app, [SOURCE_SHEET, "Склад"])
        review.selected_sheet.set("Склад")
        create.selected_sheet.set("Склад")
        ExcelCopyApp._set_source_sheets(app, [SOURCE_SHEET, "Поточний"])
        self.assertEqual(review.selected_sheet.get(), SOURCE_SHEET)
        self.assertEqual(create.selected_sheet.get(), SOURCE_SHEET)
        ExcelCopyApp._set_source_sheets(app, [])
        self.assertEqual(review.selected_sheet.get(), "")
        self.assertEqual(create.selected_sheet.get(), "")
        for panel in (review, create):
            panel.sheet_combo.configure.assert_called_with(
                values=[], state="disabled",
            )

    def test_copy_uses_creation_tab_not_review_tab(self):
        review, create = self.make_panel(), self.make_panel()
        source = SourceField(8, "Дані", 77, sheet="Склад")
        review.selected_source = Mock(return_value=SourceField(1, "Непотрібне", 1))
        create.selected_source = Mock(return_value=source)
        create.source_column = Mock(return_value=5)
        session = Mock()
        session.current.unmatched_sources = [source]
        session.current.unmatched_destinations = [DestinationField(3, "Техніка")]
        session.copy_manually.return_value = None
        app = SimpleNamespace(
            session=session, create_sources=create, review_sources=review,
            row_only_var=Mock(get=Mock(return_value=False)),
            destination_tree=Mock(selection=Mock(return_value=("3",))),
            _refresh_tables=Mock(),
        )
        ExcelCopyApp.copy_selected(app)
        session.copy_manually.assert_called_once_with(
            8, 3, sheet="Склад", source_column=5, alias_row=None,
        )
        review.selected_source.assert_not_called()


class AliasReviewTests(unittest.TestCase):
    """Exercise review actions with a real alias store but no Tk display."""

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "aliases.json"
        self.aliases = AliasStore(self.path)
        self.destination = DestinationField(3, "Техніка")
        self.source = SourceField(7, "Бронетехніка", 45, sheet="Склад")
        self.template = AliasStore.template_id([self.destination])
        self.aliases.set(self.template, self.source, self.destination)
        from excel_copy_script.models import CopyRecord, ImportResult

        self.record = CopyRecord(self.source, self.destination, "Аліас")
        session = SimpleNamespace(
            template_id=self.template,
            fields=[self.destination],
            current=ImportResult(filename="report.xlsx", row=4, copied=[self.record]),
            layout=__import__("excel_copy_script.workbook_layout", fromlist=["WorkbookLayout"]).WorkbookLayout(),
        )
        self.app = SimpleNamespace(
            session=session, aliases=self.aliases,
            copied_tree=SheetSelectorTests.FakeTree(),
            delete_alias_button=Mock(),
            _refresh_source_panels=Mock(),
            _refresh_destination_rows=Mock(),
        )
        self.app._record_has_alias = lambda record: ExcelCopyApp._record_has_alias(self.app, record)
        self.app._aliases_for_source = lambda source: ExcelCopyApp._aliases_for_source(self.app, source)
        self.app._alias_for_source = lambda source, row=None: ExcelCopyApp._alias_for_source(self.app, source, row=row)
        self.app._delete_alias_entry = lambda entry: ExcelCopyApp._delete_alias_entry(self.app, entry)
        self.app._selected_copy_record = lambda: ExcelCopyApp._selected_copy_record(self.app)
        self.app._on_copied_selection = lambda: ExcelCopyApp._on_copied_selection(self.app)
        self.app._refresh_copied = lambda: ExcelCopyApp._refresh_copied(self.app)

    def test_delete_record_alias_preserves_import_record_and_data(self):
        from unittest.mock import patch

        ExcelCopyApp._refresh_copied(self.app)
        self.app.copied_tree.selected = ("0",)
        ExcelCopyApp._on_copied_selection(self.app)
        self.app.delete_alias_button.configure.assert_called_with(state="normal")
        with patch("excel_copy_script.gui.messagebox.askyesno", return_value=True) as confirm:
            ExcelCopyApp.delete_selected_alias(self.app)
        self.assertTrue(confirm.called)
        self.assertIsNone(self.aliases.get(self.template, self.source.name, self.source.sheet))
        self.assertIsNone(AliasStore(self.path).get(self.template, self.source.name, self.source.sheet))
        self.assertEqual(self.app.session.current.copied, [self.record])
        self.assertEqual(self.app.copied_tree.rows["0"][-1], "Аліас видалено")
        self.app.delete_alias_button.configure.assert_called_with(state="disabled")

    def test_cancel_delete_preserves_alias(self):
        from unittest.mock import patch

        self.app.copied_tree.rows["0"] = ("from", "to", 45, "Аліас")
        self.app.copied_tree.selected = ("0",)
        with patch("excel_copy_script.gui.messagebox.askyesno", return_value=False):
            ExcelCopyApp.delete_selected_alias(self.app)
        self.assertIsNotNone(self.aliases.get(self.template, self.source.name, self.source.sheet))
        self.assertEqual(self.app.session.current.copied, [self.record])

    def test_copy_without_saved_alias_cannot_be_deleted(self):
        from excel_copy_script.models import CopyRecord

        exact = CopyRecord(SourceField(1, "Техніка", 42), self.destination, "Вручну")
        self.app.session.current.copied = [exact]
        ExcelCopyApp._refresh_copied(self.app)
        self.app.copied_tree.selected = ("0",)
        ExcelCopyApp._on_copied_selection(self.app)
        self.app.delete_alias_button.configure.assert_called_with(state="disabled")


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
        self.assertEqual(store.data["version"], 4)
        self.assertEqual(store.get("template-1", "Кількість", SOURCE_SHEET)["column"], 2)
        self.assertIsNone(store.get("template-1", "Кількість", "Склад"))
        self.assertEqual(store.get("template-2", "Інша назва", SOURCE_SHEET)["column"], 3)
        self.assertEqual(store.list_for_template("template-1")[0].sheet, SOURCE_SHEET)
        backup = self.directory / "aliases.json.v1.bak"
        self.assertEqual(backup.read_text(encoding="utf-8"), encoded)
        self.assertEqual(json.loads(self.path.read_text(encoding="utf-8"))["version"], 4)
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
        text = '{"version": 5, "templates": {}}'
        self.path.write_text(text, encoding="utf-8")
        with self.assertRaises(ValueError):
            AliasStore(self.path)
        self.assertEqual(self.path.read_text(encoding="utf-8"), text)
        self.assertFalse((self.directory / "aliases.json.v1.bak").exists())


if __name__ == "__main__":
    unittest.main()
