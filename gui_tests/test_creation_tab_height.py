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

    def test_copy_action_and_help_share_same_horizontal_row(self):
        self.app.notebook.select(1)
        self.app.update()
        action = self.app.copy_button
        container = action.master
        children = container.winfo_children()
        self.assertEqual(len(children), 2)
        self.assertEqual(action.pack_info()["side"], "left")
        self.assertEqual(children[1].pack_info()["side"], "left")
        # They should occupy one control strip, not two stacked rows.
        self.assertLess(abs(children[1].winfo_y() - action.winfo_y()), 15)


if __name__ == "__main__":
    unittest.main()
