"""GUI regression checks for the taller creation-tab source table."""

import unittest

from excel_copy_script.gui import ExcelCopyApp


class CreationTabHeightTests(unittest.TestCase):
    def setUp(self):
        self.app = ExcelCopyApp()
        self.addCleanup(self.app.destroy)
        self.app.update()

    def test_creation_source_rows_gain_height_from_default_window_size(self):
        screen_height = self.app.winfo_screenheight()
        expected = min(900, max(510, screen_height - 80))
        self.assertEqual(self.app.winfo_height(), expected)
        self.app.notebook.select(1)
        self.app.update()
        enlarged = self.app.create_sources.values_tree.winfo_height()
        self.assertGreater(enlarged, 100)

        if expected > 720:
            self.app.geometry("1260x720")
            self.app.update()
            older_window = self.app.create_sources.values_tree.winfo_height()
            self.assertGreater(enlarged, older_window + 70)

    def test_source_rows_tree_uses_almost_all_destination_tree_height(self):
        # Previously the sheet/value selectors and an empty warning strip
        # consumed ~50-65px more than the neighboring destination header.
        self.app.geometry("1400x850")
        self.app.notebook.select(1)
        self.app.update()
        incoming = self.app.create_sources.values_tree.winfo_height()
        outgoing = self.app.destination_tree.winfo_height()
        self.assertGreater(incoming, 300)
        self.assertGreater(outgoing, 300)
        self.assertLessEqual(
            outgoing - incoming, 35,
            f"Incoming rows are still too short: {incoming=} {outgoing=}",
        )
        self.assertTrue(self.app.create_sources.sheet_combo.winfo_ismapped())
        self.assertTrue(self.app.create_sources.value_combo.winfo_ismapped())
        self.assertFalse(self.app.create_sources.issue_label.winfo_ismapped())

    def test_warning_takes_height_only_while_row_has_a_problem(self):
        from excel_copy_script.models import SourceField

        self.app.geometry("1400x850")
        self.app.notebook.select(1)
        panel = self.app.create_sources
        panel.set_sheets(["ВИСНОВКИ"])
        panel.refresh(
            [SourceField(5, "Техніка", 19), SourceField(8, "Резерв", 22)],
            alias_targets={("висновки", 8): "E: Резерв"},
            alias_issues={("висновки", 8): "Конфлікт двох аліасів"},
        )
        self.app.update()
        full_height = panel.values_tree.winfo_height()
        self.assertFalse(panel.issue_label.winfo_ismapped())

        panel.tree.selection_set("row-8")
        self.app.update()
        self.assertTrue(panel.issue_label.winfo_ismapped())
        self.assertIn("Конфлікт", panel.issue_var.get())
        self.assertLess(panel.values_tree.winfo_height(), full_height)

        panel.tree.selection_set("row-5")
        self.app.update()
        self.assertFalse(panel.issue_label.winfo_ismapped())
        self.assertEqual(panel.values_tree.winfo_height(), full_height)

    def test_copy_action_and_help_share_same_horizontal_row(self):
        self.app.notebook.select(1)
        self.app.update()
        action = self.app.copy_button
        container = action.master
        children = container.winfo_children()
        self.assertEqual(len(children), 3)
        self.assertEqual(action.pack_info()["side"], "left")
        self.assertEqual(children[1].pack_info()["side"], "left")
        # They should occupy one control strip, not two stacked rows.
        self.assertLess(abs(children[1].winfo_y() - action.winfo_y()), 15)


if __name__ == "__main__":
    unittest.main()
