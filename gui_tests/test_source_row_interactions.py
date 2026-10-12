"""Real Tk regression tests for synchronized source row selection.

Run under Xvfb (see .github/workflows/tests.yml). These tests deliberately
send mouse events to the right-hand value cells: that was the freeze trigger.
"""

import tkinter as tk
import unittest

from excel_copy_script.models import SourceField
from excel_copy_script.source_rows_panel import SourceRowsPanel


class SourceRowsPanelInteractionTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.geometry("1150x550")
        self.events = 0
        self.panel = SourceRowsPanel(self.root, on_selection_changed=self._count_event)
        self.panel.pack(fill="both", expand=True)
        self.panel.set_sheets(["ВИСНОВКИ"])
        self.panel.refresh([
            SourceField(i, f"Показник {i}", i * 10, values=(
                i * 10, i * 20, i * 30, i * 40, i * 50,
            ))
            for i in range(1, 65)
        ])
        self.root.update()
        self.addCleanup(self.root.destroy)

    def _count_event(self):
        self.events += 1
        if self.events > 60:
            self.fail("Recursive <<TreeviewSelect>> event storm")

    def _click(self, tree, row, column):
        rect = tree.bbox(f"row-{row}", column)
        self.assertTrue(rect, f"Expected row {row}, column {column} to be visible")
        x, y, width, height = rect
        self.assertGreater(width, 10)
        tree.event_generate(
            "<ButtonPress-1>", x=x + 8, y=y + height // 2,
        )
        tree.event_generate(
            "<ButtonRelease-1>", x=x + 8, y=y + height // 2,
        )
        # Deliver *queued* <<TreeviewSelect>> callbacks. Before the fix,
        # update() never returned due to the peer selection event ping-pong.
        self.root.update()

    def test_clicks_on_value_cells_do_not_freeze(self):
        self._click(self.panel.values_tree, 1, "C")
        self.assertEqual(self.panel.selected_source().row, 1)
        self.assertEqual(self.panel.tree.selection(), ("row-1",))
        self.assertEqual(self.panel.values_tree.selection(), ("row-1",))
        self._click(self.panel.values_tree, 2, "D")
        self.assertEqual(self.panel.selected_source().row, 2)
        self.assertEqual(self.panel.source_column(), 4)
        self._click(self.panel.values_tree, 2, "D")
        self.assertLess(self.events, 20)

    def test_frozen_table_click_syncs_only_once(self):
        self._click(self.panel.tree, 3, "name")
        self.assertEqual(self.panel.values_tree.selection(), ("row-3",))
        self._click(self.panel.values_tree, 1, "C")
        self.assertEqual(self.panel.tree.selection(), ("row-1",))
        self.assertLess(self.events, 20)

    def test_vertical_scroll_keeps_both_halves_aligned(self):
        self.panel.values_tree.yview_moveto(0.6)
        self.root.update()
        left, right = self.panel.tree.yview(), self.panel.values_tree.yview()
        self.assertAlmostEqual(left[0], right[0], delta=0.015)
        self.panel.tree.yview_moveto(0.3)
        self.root.update()
        left, right = self.panel.tree.yview(), self.panel.values_tree.yview()
        self.assertAlmostEqual(left[0], right[0], delta=0.015)


if __name__ == "__main__":
    unittest.main()
