"""Integration checks for the destination list in the real Tk notebook."""

import unittest

from excel_copy_script.gui import ExcelCopyApp
from excel_copy_script.models import CopyRecord, DestinationField, ImportResult, SourceField


class DestinationOverviewTkTests(unittest.TestCase):
    def setUp(self):
        self.app = ExcelCopyApp()
        self.app.geometry("1400x850")
        self.addCleanup(self.app.destroy)
        self.output_fields = [
            DestinationField(2, "Особовий склад"),
            DestinationField(3, "Техніка"),
            DestinationField(4, "Резерв"),
        ]
        self.copied_source = SourceField(
            8, "Бронетехніка", 99, sheet="Склад", value_column=6,
            values=(1, 2, 3, 99, 5),
        )
        self.unmatched_source = SourceField(5, "Вхідний резерв", 40)
        self.copied = CopyRecord(self.copied_source, self.output_fields[1], "Аліас")
        self.result = ImportResult(
            filename="input.xlsx",
            row=4,
            copied=[self.copied],
            unmatched_destinations=[self.output_fields[0], self.output_fields[2]],
            unmatched_sources=[self.unmatched_source],
            sheets=["ВИСНОВКИ", "Склад"],
            all_sources=[self.copied_source, self.unmatched_source],
        )

        class FakeSession:
            def __init__(self, fields, current):
                self.fields = fields
                self.current = current

            def source_alias_targets(self):
                return {}

            @property
            def layout(self):
                from excel_copy_script.workbook_layout import WorkbookLayout
                return WorkbookLayout()

            def skipped_alias_destinations(self):
                return {}

        self.app.session = FakeSession(self.output_fields, self.result)
        self.app.aliases = None  # FakeSession has no real template/alias store.
        self.app.notebook.select(1)
        self.app._set_source_sheets(self.result.sheets)
        self.app._refresh_tables()
        self.app.update()

    def test_all_output_columns_and_provenance(self):
        tree = self.app.destination_tree
        self.assertEqual(tree.get_children(), ("2", "3", "4"))
        self.assertEqual(tuple(tree.item("2", "tags")), ())
        self.assertEqual(tuple(tree.item("4", "tags")), ())
        self.assertEqual(tree.item("3", "tags"), ("filled",))
        self.assertEqual(tree.item("3", "values"), (
            "C", "Техніка", "Склад!F8: Бронетехніка",
        ))
        self.assertEqual(tree.item("4", "values"), ("D", "Резерв", ""))

    def test_filled_column_cannot_be_selected_for_copy(self):
        self.app.create_sources.tree.selection_set("row-5")
        self.app.update()

        self.app.destination_tree.selection_set("3")
        self.app.update()
        self.assertEqual(str(self.app.copy_button.cget("state")), "disabled")

        self.app.destination_tree.selection_set("4")
        self.app.update()
        self.assertEqual(str(self.app.copy_button.cget("state")), "normal")

    def test_newly_copied_column_turns_green_and_shows_source(self):
        self.result.copied.append(CopyRecord(
            self.unmatched_source, self.output_fields[2], "Вручну",
        ))
        self.result.unmatched_destinations.remove(self.output_fields[2])
        self.result.unmatched_sources.remove(self.unmatched_source)
        self.app._refresh_tables()
        self.app.update()

        tree = self.app.destination_tree
        self.assertEqual(tree.item("4", "tags"), ("filled",))
        self.assertEqual(tree.item("4", "values"), (
            "D", "Резерв", "ВИСНОВКИ!C5: Вхідний резерв",
        ))
        self.assertEqual(tree.get_children(), ("2", "3", "4"))


if __name__ == "__main__":
    unittest.main()
