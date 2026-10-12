"""Real Tk layout resize + restart tests under Xvfb."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from excel_copy_script.gui import ExcelCopyApp
from excel_copy_script.gui_layout import LayoutPreferences


class LayoutPersistenceTkTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.config_path = Path(tmp.name) / "layout.json"

    def app(self):
        with patch(
            "excel_copy_script.gui_layout.LayoutPreferences",
            side_effect=lambda: LayoutPreferences(self.config_path),
        ):
            app = ExcelCopyApp()
        app.geometry("1650x820")
        app.update()
        return app

    def test_all_six_tables_and_four_sashes_restore_on_relaunch(self):
        first = self.app()
        self.addCleanup(lambda: first.winfo_exists() and first.destroy())
        first.notebook.select(1)
        first.update()

        trees = {
            "review_copied": (first.copied_tree, "Звідки", 243),
            "review_source_headers": (first.review_sources.tree, "name", 205),
            "review_source_values": (first.review_sources.values_tree, "E", 181),
            "create_destination": (first.destination_tree, "Назва", 272),
            "create_source_headers": (first.create_sources.tree, "alias", 127),
            "create_source_values": (first.create_sources.values_tree, "D", 189),
        }
        for key, (tree, column, width) in trees.items():
            tree.column(column, width=width)
            first.layout.capture_columns(key, tree)

        for key, pane in (
            ("create_outer", first.create_panes),
            ("create_inner", first.create_sources.divider),
        ):
            pane.sashpos(0, 550 if key == "create_outer" else 460)
            first.update()
            first.layout.capture_sash(key, pane)
        first.notebook.select(0)
        first.update()
        for key, pane in (
            ("review_outer", first.review_panes),
            ("review_inner", first.review_sources.divider),
        ):
            pane.sashpos(0, 590 if key == "review_outer" else 440)
            first.update()
            first.layout.capture_sash(key, pane)

        expected = LayoutPreferences(self.config_path)
        self.assertEqual(len(expected.columns), 6)
        self.assertEqual(len(expected.sashes), 4)
        first.destroy()

        second = self.app()
        self.addCleanup(lambda: second.winfo_exists() and second.destroy())
        # The first tab is restored as it becomes visible on startup.
        self.assertAlmostEqual(second.review_panes.sashpos(0),
                               expected.sashes["review_outer"], delta=10)
        self.assertAlmostEqual(second.review_sources.divider.sashpos(0),
                               expected.sashes["review_inner"], delta=10)
        second.notebook.select(1)
        second.update()
        self.assertAlmostEqual(second.create_panes.sashpos(0),
                               expected.sashes["create_outer"], delta=10)
        self.assertAlmostEqual(second.create_sources.divider.sashpos(0),
                               expected.sashes["create_inner"], delta=10)
        # Check persisted widths on both tabs after they have been made visible.
        new_trees = {
            "review_copied": second.copied_tree,
            "review_source_headers": second.review_sources.tree,
            "review_source_values": second.review_sources.values_tree,
            "create_destination": second.destination_tree,
            "create_source_headers": second.create_sources.tree,
            "create_source_values": second.create_sources.values_tree,
        }
        for key, (original_tree, column, _width) in trees.items():
            actual = int(new_trees[key].column(column, "width"))
            saved = expected.columns[key][column]
            self.assertAlmostEqual(actual, saved, delta=10, msg=key)

    def test_user_drags_are_saved_without_closing(self):
        app = self.app()
        self.addCleanup(lambda: app.winfo_exists() and app.destroy())
        app.notebook.select(1)
        app.update()
        pane = app.create_sources.divider
        original = pane.sashpos(0)
        y = max(50, pane.winfo_height() // 2)
        for event, x in (
            ("<ButtonPress-1>", original),
            ("<B1-Motion>", original + 75),
            ("<ButtonRelease-1>", original + 75),
        ):
            pane.event_generate(event, x=x, y=y)
            app.update()
        self.assertAlmostEqual(
            LayoutPreferences(self.config_path).sashes["create_inner"],
            pane.sashpos(0), delta=3,
        )

        # Resize the first Treeview column by its actual header separator.
        tree = app.create_sources.tree
        tree.update()
        current = int(tree.column("alias", "width"))
        separator = next((
            x for x in range(max(5, current - 12), current + 13)
            if tree.identify_region(x, 8) == "separator"
        ), None)
        self.assertIsNotNone(separator, "Could not locate the Tk heading divider")
        for event, x in (
            ("<ButtonPress-1>", separator),
            ("<B1-Motion>", separator + 40),
            ("<ButtonRelease-1>", separator + 40),
        ):
            tree.event_generate(event, x=x, y=8)
            app.update()

        saved = LayoutPreferences(self.config_path)
        self.assertIn("create_source_headers", saved.columns)
        self.assertEqual(
            saved.columns["create_source_headers"]["alias"],
            int(tree.column("alias", "width")),
        )
        self.assertGreater(saved.columns["create_source_headers"]["alias"], current + 20)

    def test_unvisited_tab_keeps_its_saved_sash_positions(self):
        prefs = LayoutPreferences(self.config_path)
        prefs.set_sash("create_outer", 710)
        prefs.set_sash("create_inner", 445)
        app = self.app()
        self.addCleanup(lambda: app.winfo_exists() and app.destroy())
        self.assertFalse(app.create_panes.winfo_ismapped())
        # Repeated attempts to restore a hidden page must not reset its settings.
        app.layout.restore_tab("create")
        self.assertEqual(LayoutPreferences(self.config_path).sashes["create_outer"], 710)
        app.notebook.select(1)
        app.update()
        self.assertAlmostEqual(app.create_panes.sashpos(0), 710, delta=10)
        self.assertAlmostEqual(app.create_sources.divider.sashpos(0), 445, delta=10)


if __name__ == "__main__":
    unittest.main()
