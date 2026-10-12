"""Xvfb integration: skipped aliases must not be displayed as green successes."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from excel_copy_script.aliases import AliasStore
from excel_copy_script.gui import ExcelCopyApp
from excel_copy_script.models import SourceField
from excel_copy_script.workbook import WorkbookSession

PARAMETER = "Загальна результативність виконання завдань"


class SkippedAliasesGuiTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        folder = Path(temp.name)

        dest = folder / "destination.xlsx"
        workbook = Workbook()
        output = workbook.active
        output.title = "Дані_підрозділів"
        output["A3"], output["E3"] = "№", "Результативність"
        workbook.save(dest)
        self.aliases = AliasStore(folder / "aliases.json")
        self.session = WorkbookSession(dest, self.aliases)
        destination = self.session.fields[0]
        for name in (PARAMETER, "Конфліктний показник"):
            self.aliases.set(
                self.session.template_id, SourceField(14, name, "89%"),
                destination,
            )

        source = folder / "source.xlsx"
        workbook = Workbook()
        input_sheet = workbook.active
        input_sheet.title = "ВИСНОВКИ"
        input_sheet["B14"], input_sheet["C14"] = PARAMETER, "89%"
        input_sheet["D14"], input_sheet["E14"] = "80%", "П1"
        input_sheet["F14"] = "у межах норми"
        input_sheet["B15"], input_sheet["C15"] = "Конфліктний показник", "15%"
        workbook.save(source)
        self.result = self.session.open_input(source)

        self.app = ExcelCopyApp()
        self.addCleanup(self.app.destroy)
        self.app.geometry("1400x840")
        self.app.session = self.session
        self.app.aliases = self.aliases
        self.app.notebook.select(1)
        self.app._set_source_sheets(self.result.sheets)
        self.app._refresh_tables()
        self.app.update()

    def test_conflicting_source_rows_yellow_and_destination_e_not_green(self):
        tree = self.app.create_sources.tree
        self.assertEqual(tree.item("row-14", "tags"), ("skipped",))
        self.assertEqual(self.app.create_sources.values_tree.item("row-14", "tags"), ("skipped",))
        self.assertTrue(tree.item("row-14", "values")[0].startswith("⚠ E:"))
        self.assertEqual(tree.item("row-15", "tags"), ("skipped",))
        self.assertNotIn("filled", self.app.destination_tree.item("5", "tags"))
        tree.selection_set("row-14")
        self.app.update()
        message = self.app.create_sources.issue_var.get()
        self.assertIn("B14", message)
        self.assertIn("претендують", message)
        self.assertIn("пропущено аліасів: 2", self.app.status_var.get())

    def test_only_valid_alias_turns_green_after_conflict_is_removed(self):
        self.aliases.delete(self.session.template_id, "Конфліктний показник")
        # The current Excel run is not automatically replayed by alias deletion.
        # A new run applies the remaining alias normally.
        source = self.session.current
        self.assertIsNotNone(source)
        # Re-import into the next available row to simulate opening the file again.
        input_file = self.session.path.parent / "source.xlsx"
        result = self.session.open_input(input_file)
        self.app._refresh_tables()
        self.app.update()
        self.assertEqual(len(result.copied), 1)
        self.assertEqual(self.app.create_sources.tree.item("row-14", "tags"), ("aliased",))
        self.assertEqual(tuple(self.app.create_sources.tree.item("row-15", "tags")), ())
        self.assertEqual(self.app.destination_tree.item("5", "tags"), ("filled",))
        self.assertEqual(self.session.sheet["E4"].value, "89%")


if __name__ == "__main__":
    unittest.main()
