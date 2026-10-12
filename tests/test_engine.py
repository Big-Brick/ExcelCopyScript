import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from excel_copy_script.aliases import AliasStore
from excel_copy_script.models import DestinationField, SourceField
from excel_copy_script.utils import normalize
from excel_copy_script.workbook import WorkbookSession


class TransferTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        self.target = self.directory / "summary.xlsx"
        self.aliases = AliasStore(self.directory / "config" / "aliases.json")
        wb = Workbook()
        ws = wb.active
        ws.title = "Дані_підрозділів"
        ws["A3"] = "Підрозділ"
        ws["B3"] = "Особовий склад"
        ws["C3"] = "Техніка"
        ws["D3"] = "Резерв"
        ws["A4"] = "Вже заповнено"
        ws["B4"] = 99
        wb.save(self.target)
        self.session = WorkbookSession(self.target, self.aliases)

    def source(self, name, rows):
        path = self.directory / name
        wb = Workbook()
        ws = wb.active
        ws.title = "ВИСНОВКИ"
        for row, title, value in rows:
            ws.cell(row, 2, title)
            ws.cell(row, 3, value)
        wb.save(path)
        return path

    def test_exact_manual_alias_next_file_and_reopen(self):
        first = self.source("first.xlsx", [(1, "Особовий  склад", 13), (2, "Бронетехніка", 7),
                                               (4, "Розділ звіту", None)])
        result = self.session.open_input(first)
        self.assertEqual(result.row, 5)
        self.assertEqual(len(result.copied), 1)
        self.assertEqual(result.copied[0].destination.column, 2)
        self.assertEqual({item.row for item in result.unmatched_sources}, {2})
        self.assertIsNone(self.session.copy_manually(2, 3))
        self.assertTrue(self.aliases.path.exists())
        self.assertEqual(len(self.session.current.copied), 2)

        # Persistent aliases are resolved by header, not by source/destination row.
        reloaded_aliases = AliasStore(self.aliases.path)
        self.session = WorkbookSession(self.target, reloaded_aliases)
        second = self.source("second.xlsx", [(6, "особовий склад", 15), (7, "Бронетехніка", 8)])
        result = self.session.open_input(second)
        self.assertEqual(result.row, 6)
        self.assertEqual({item.method for item in result.copied}, {"Збіг назви", "Аліас"})
        ws = load_workbook(self.target, data_only=True)["Дані_підрозділів"]
        self.assertEqual(ws["B4"].value, 99)
        self.assertEqual(ws["B5"].value, 13)
        self.assertEqual(ws["C5"].value, 7)
        self.assertEqual(ws["B6"].value, 15)
        self.assertEqual(ws["C6"].value, 8)
        self.assertEqual(len(list(self.directory.glob("summary.backup-*.xlsx"))), 2)

    def test_duplicate_source_names_are_not_auto_matched(self):
        file = self.source("duplicate.xlsx", [(1, "Техніка", 1), (3, " ТЕХНІКА ", 2)])
        result = self.session.open_input(file)
        self.assertEqual(len(result.copied), 0)
        self.assertEqual(len(result.unmatched_sources), 2)
        self.assertTrue(any("Повторювані" in note for note in result.warnings))
        self.assertFalse(list(self.directory.glob("summary.backup-*.xlsx")))

    def test_column_a_is_excluded_from_fields_and_unmatched_list(self):
        self.assertEqual([field.column for field in self.session.fields], [2, 3, 4])
        source = self.source("index.xlsx", [(1, "Підрозділ", 42)])
        result = self.session.open_input(source)
        self.assertFalse(result.copied)
        self.assertEqual([field.column for field in result.unmatched_destinations], [2, 3, 4])
        self.assertEqual([field.name for field in result.unmatched_sources], ["Підрозділ"])
        with self.assertRaises(ValueError):
            self.session.copy_manually(1, 1)

    def test_prepopulated_ordinal_does_not_block_import(self):
        ws = self.session.sheet
        ws["A5"] = 5
        ws["A6"] = 6
        source = self.source("numbered.xlsx", [(2, "Особовий склад", 21)])
        result = self.session.open_input(source)
        self.assertEqual(result.row, 5)
        self.assertEqual([record.destination.column for record in result.copied], [2])
        second_source = self.source("next_numbered.xlsx", [(3, "Техніка", 11)])
        second_result = self.session.open_input(second_source)
        self.assertEqual(second_result.row, 6)
        self.assertEqual(self.session.sheet["A5"].value, 5)
        self.assertEqual(self.session.sheet["A6"].value, 6)
        saved = load_workbook(self.target)["Дані_підрозділів"]
        self.assertEqual((saved["A5"].value, saved["B5"].value), (5, 21))
        self.assertEqual((saved["A6"].value, saved["C6"].value), (6, 11))

    def test_changing_ordinal_header_does_not_change_alias_template(self):
        other = Workbook()
        ws = other.active
        ws.title = "Дані_підрозділів"
        ws["A3"] = "№"
        ws["B3"] = "Особовий склад"
        ws["C3"] = "Техніка"
        ws["D3"] = "Резерв"
        path = self.directory / "renamed_a.xlsx"
        other.save(path)
        renamed = WorkbookSession(path, self.aliases)
        self.assertEqual(self.session.template_id, renamed.template_id)

    def test_migrate_old_aliases_without_migrating_mapping_to_a(self):
        legacy_fields = [DestinationField(1, "Підрозділ"), *self.session.fields]
        old_id = self.aliases.template_id(legacy_fields)
        self.assertNotEqual(old_id, self.session.template_id)
        self.aliases.set(
            old_id, SourceField(1, "Бронетехніка", 3), self.session.fields[1]
        )
        self.aliases.set(
            old_id, SourceField(2, "Старий номер", 4), legacy_fields[0]
        )
        new_session = WorkbookSession(self.target, AliasStore(self.aliases.path))
        migrated = new_session.aliases.get(new_session.template_id, "Бронетехніка")
        self.assertIsNotNone(migrated)
        self.assertEqual(migrated["column"], 3)
        self.assertIsNone(new_session.aliases.get(new_session.template_id, "Старий номер"))

        source = self.source("legacy.xlsx", [(1, "Бронетехніка", 29), (2, "Старий номер", 7)])
        result = new_session.open_input(source)
        self.assertEqual(len(result.copied), 1)
        self.assertEqual(result.copied[0].method, "Аліас")
        self.assertEqual(result.copied[0].destination.column, 3)
        self.assertEqual(new_session.sheet["A5"].value, None)
        self.assertEqual(new_session.sheet["C5"].value, 29)
        reloaded = AliasStore(self.aliases.path)
        self.assertIsNotNone(reloaded.get(new_session.template_id, "Бронетехніка"))

    def test_existing_current_alias_survives_legacy_migration(self):
        current = self.session.template_id
        old = self.aliases.template_id([DestinationField(1, "Підрозділ"),
                                        *self.session.fields])
        name = SourceField(1, "Чисельність", 1)
        self.aliases.set(old, name, self.session.fields[1])
        self.aliases.set(current, name, self.session.fields[0])
        session = WorkbookSession(self.target, AliasStore(self.aliases.path))
        self.assertEqual(session.aliases.get(current, name.name)["column"], 2)

    def test_first_completely_empty_row_skips_partial(self):
        ws = self.session.sheet
        ws["D5"] = 1
        self.assertEqual(self.session.next_empty_row(), 6)

    def test_aliases_scoped_to_output_template(self):
        first = self.source("match.xlsx", [(1, "Нова назва", 55)])
        self.session.open_input(first)
        self.session.copy_manually(1, 2)
        other = Workbook()
        ws = other.active
        ws.title = "Дані_підрозділів"
        ws["A3"] = "№"
        ws["B3"] = "Інша назва колонки"
        different = self.directory / "other.xlsx"
        other.save(different)
        another_session = WorkbookSession(different, AliasStore(self.aliases.path))
        result = another_session.open_input(first)
        self.assertFalse(result.copied)

    def test_conflicting_candidates_do_not_overwrite_each_other(self):
        source = self.source("conflict.xlsx", [(1, "Техніка", 11), (2, "Запас техніки", 22)])
        self.aliases.set(self.session.template_id,
                         type("S", (), {"key": normalize("Запас техніки")})(), self.session.fields[1])
        result = self.session.open_input(source)
        self.assertFalse(result.copied)
        self.assertEqual(len(result.unmatched_sources), 2)
        self.assertTrue(any("претендують" in note for note in result.warnings))

    def test_missing_formula_cache_warns_and_does_not_copy(self):
        file = self.source("formulas.xlsx", [(3, "Техніка", "=2+2")])
        result = self.session.open_input(file)
        self.assertFalse(result.copied)
        self.assertTrue(any("формула" in warning for warning in result.warnings))

    def test_untrusted_string_does_not_become_formula(self):
        source = self.source("literal.xlsx", [(1, "Техніка", "'=cmd")])
        result = self.session.open_input(source)
        self.assertEqual(len(result.copied), 1)
        wb = load_workbook(self.target)
        self.assertNotEqual(wb["Дані_підрозділів"]["C5"].data_type, "f")

    def test_alias_listing_and_deletion_are_persistent(self):
        source = self.source("aliases.xlsx", [(1, "Бронетехніка", 17)])
        self.session.open_input(source)
        self.session.copy_manually(1, 3)

        template_id = self.session.template_id
        entries = self.aliases.list_for_template(template_id)
        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0].source_name, "Бронетехніка")
        self.assertEqual(entries[0].source_key, normalize("Бронетехніка"))
        self.assertEqual(entries[0].destination_key, normalize("Техніка"))
        self.assertEqual(entries[0].destination_column, 3)

        self.assertTrue(self.aliases.delete(template_id, " БРОНЕТЕХНІКА "))
        self.assertFalse(self.aliases.delete(template_id, "Бронетехніка"))
        self.assertEqual(self.aliases.list_for_template(template_id), [])
        reloaded = AliasStore(self.aliases.path)
        self.assertEqual(reloaded.list_for_template(template_id), [])
        self.assertIsNone(reloaded.get(template_id, "Бронетехніка"))

        next_file = self.source("next.xlsx", [(3, "Бронетехніка", 10)])
        next_result = self.session.open_input(next_file)
        self.assertFalse(next_result.copied)
        self.assertEqual([field.name for field in next_result.unmatched_sources],
                         ["Бронетехніка"])
        # Removing an alias must not modify the already copied value.
        self.assertEqual(load_workbook(self.target)["Дані_підрозділів"]["C5"].value, 17)

    def test_deleting_one_alias_does_not_remove_other_aliases_or_templates(self):
        template_id = self.session.template_id
        source = self.source("many.xlsx", [(1, "Бронетехніка", 10), (2, "Бійці", 20)])
        self.session.open_input(source)
        self.session.copy_manually(1, 3)
        self.session.copy_manually(2, 2)
        self.aliases.set("other-template",
                         type("S", (), {"key": normalize("Альтернативна назва"), "name": "Альтернативна назва"})(),
                         self.session.fields[1])
        self.assertTrue(self.aliases.delete(template_id, "Бронетехніка"))
        self.assertIsNotNone(self.aliases.get(template_id, "Бійці"))
        self.assertIsNotNone(self.aliases.get("other-template", "Альтернативна назва"))

        reloaded = AliasStore(self.aliases.path)
        self.assertEqual(len(reloaded.list_for_template(template_id)), 1)
        self.assertEqual(len(reloaded.list_for_template("other-template")), 1)

    def test_delete_rolls_back_on_save_failure(self):
        from unittest.mock import patch

        template_id = self.session.template_id
        self.aliases.set(template_id,
                         type("S", (), {"key": normalize("Бронетехніка"), "name": "Бронетехніка"})(),
                         self.session.fields[1])
        with patch.object(self.aliases, "_save", side_effect=OSError("disk full")):
            with self.assertRaises(OSError):
                self.aliases.delete(template_id, "Бронетехніка")
        self.assertIsNotNone(self.aliases.get(template_id, "Бронетехніка"))
        self.assertIsNotNone(AliasStore(self.aliases.path).get(template_id, "Бронетехніка"))

    def test_legacy_aliases_can_be_listed_and_deleted(self):
        import json

        template_id = self.session.template_id
        self.aliases.path.parent.mkdir(parents=True, exist_ok=True)
        self.aliases.path.write_text(json.dumps({"version": 1, "templates": {
            template_id: {"бронетехніка": {"header": "техніка", "column": 3}}
        }}, ensure_ascii=False), encoding="utf-8")
        legacy = AliasStore(self.aliases.path)
        entries = legacy.list_for_template(template_id)
        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0].source_name, "бронетехніка")
        self.assertTrue(legacy.delete(template_id, "БРОНЕТЕХНІКА"))
        self.assertEqual(AliasStore(self.aliases.path).list_for_template(template_id), [])

    def test_bad_alias_file_fails_loudly(self):
        self.aliases.path.parent.mkdir(parents=True, exist_ok=True)
        self.aliases.path.write_text("not json", encoding="utf-8")
        with self.assertRaises(ValueError):
            AliasStore(self.aliases.path)


if __name__ == "__main__":
    unittest.main()
