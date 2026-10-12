"""Persistent Tk layout settings: validation, durability and corruption recovery."""

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from excel_copy_script.gui_layout import LayoutPreferences


class LayoutPreferencesTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "preferences" / "layout.json"

    def test_independent_tables_and_splitters_persist_across_instances(self):
        layout = LayoutPreferences(self.path)
        layout.set_columns("review_copied", {"Звідки": 190, "Куди": 250})
        layout.set_columns("create_destination", {"Назва": 420})
        layout.set_columns("review_source_headers", {"name": 285})
        layout.set_columns("create_source_values", {"E": 161})
        layout.set_sash("review_outer", 520)
        layout.set_sash("create_outer", 710)
        layout.set_sash("review_inner", 370)
        layout.set_sash("create_inner", 650)

        loaded = LayoutPreferences(self.path)
        self.assertEqual(loaded.columns["review_copied"], {"Звідки": 190, "Куди": 250})
        self.assertEqual(loaded.columns["create_destination"], {"Назва": 420})
        self.assertEqual(loaded.columns["review_source_headers"], {"name": 285})
        self.assertEqual(loaded.columns["create_source_values"], {"E": 161})
        self.assertEqual(loaded.sashes, {
            "review_outer": 520, "create_outer": 710,
            "review_inner": 370, "create_inner": 650,
        })
        self.assertEqual(json.loads(self.path.read_text())["version"], 1)
        self.assertEqual(list(self.path.parent.glob(".layout-*.tmp")), [])

    def test_invalid_values_ignored_without_crashing(self):
        self.path.parent.mkdir()
        self.path.write_text(json.dumps({
            "version": 1,
            "columns": {
                "valid": {"A": 100, "B": True, "C": -1, "D": 99999, "E": 175.5},
                "malformed": "not a mapping",
            },
            "sashes": {
                "a": 400, "b": False, "c": 5, "d": "400",
            },
        }))
        loaded = LayoutPreferences(self.path)
        self.assertEqual(loaded.columns, {"valid": {"A": 100}})
        self.assertEqual(loaded.sashes, {"a": 400})

        self.path.write_text("{broken", encoding="utf-8")
        self.assertEqual(LayoutPreferences(self.path).columns, {})
        self.assertEqual(LayoutPreferences(self.path).sashes, {})
        self.path.write_text(json.dumps({"version": 99, "columns": {"a": {"x": 200}}}))
        self.assertEqual(LayoutPreferences(self.path).columns, {})

    def test_missing_file_uses_defaults_without_creating_file(self):
        loaded = LayoutPreferences(self.path)
        self.assertEqual(loaded.columns, {})
        self.assertEqual(loaded.sashes, {})
        self.assertFalse(self.path.exists())

    def test_invalid_resizes_do_not_replace_last_good_settings(self):
        p = LayoutPreferences(self.path)
        p.set_columns("columns", {"A": 240})
        p.set_columns("columns", {"B": True, "C": 30, "D": 40000})
        p.set_sash("sash", 470)
        p.set_sash("sash", -20)
        self.assertEqual(LayoutPreferences(self.path).columns["columns"], {"A": 240})
        self.assertEqual(LayoutPreferences(self.path).sashes["sash"], 470)

    def test_atomic_save_failure_keeps_original_file(self):
        p = LayoutPreferences(self.path)
        p.set_columns("table", {"A": 180})
        snapshot = self.path.read_bytes()
        with patch("excel_copy_script.gui_layout.os.replace", side_effect=OSError("full disk")):
            with self.assertRaisesRegex(OSError, "full disk"):
                p.set_sash("split", 320)
        self.assertEqual(self.path.read_bytes(), snapshot)
        self.assertEqual(list(self.path.parent.glob(".layout-*.tmp")), [])


if __name__ == "__main__":
    unittest.main()
