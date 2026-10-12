"""Regression tests for remembered open-file dialog directories."""

import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from excel_copy_script.gui import ExcelCopyApp
from excel_copy_script.recent_directories import RecentDirectories


class RecentDirectoriesTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.destination_dir = self.root / "summaries"
        self.source_dir = self.root / "reports"
        self.destination_dir.mkdir()
        self.source_dir.mkdir()
        self.config_path = self.root / "settings" / "recent_directories.json"
        self.store = RecentDirectories(self.config_path)

    def test_separate_directories_survive_restart(self):
        self.store.remember_file("destination", self.destination_dir / "summary.xlsx")
        self.store.remember_file("source", self.source_dir / "report.xlsx")
        reloaded = RecentDirectories(self.config_path)
        self.assertEqual(reloaded.initial_directory("destination"), self.destination_dir)
        self.assertEqual(reloaded.initial_directory("source"), self.source_dir)
        data = json.loads(self.config_path.read_text(encoding="utf-8"))
        self.assertEqual(data["version"], 1)
        self.assertEqual(data["directories"]["source"], str(self.source_dir))
        self.assertEqual(data["directories"]["destination"], str(self.destination_dir))

        next_dir = self.root / "next_reports"
        next_dir.mkdir()
        reloaded.remember_file("source", next_dir / "next.xlsx")
        self.assertEqual(
            RecentDirectories(self.config_path).initial_directory("source"), next_dir,
        )
        self.assertEqual(
            RecentDirectories(self.config_path).initial_directory("destination"),
            self.destination_dir,
        )

    def test_fallback_for_first_run_and_removed_folder(self):
        self.assertEqual(
            self.store.initial_directory("source", self.destination_dir),
            self.destination_dir,
        )
        self.assertEqual(self.store.initial_directory("destination"), Path.home())
        self.store.remember_file("source", self.source_dir / "report.xlsx")
        self.source_dir.rmdir()
        self.assertEqual(
            RecentDirectories(self.config_path).initial_directory(
                "source", fallback=self.destination_dir,
            ),
            self.destination_dir,
        )
        self.assertEqual(
            RecentDirectories(self.config_path).initial_directory("source"),
            Path.home(),
        )

    def test_bad_preferences_do_not_prevent_opening(self):
        self.config_path.parent.mkdir()
        self.config_path.write_text("{invalid", encoding="utf-8")
        fresh = RecentDirectories(self.config_path)
        self.assertEqual(fresh.initial_directory("source", self.source_dir), self.source_dir)
        fresh.remember_file("source", self.source_dir / "file.xlsx")
        self.assertEqual(
            RecentDirectories(self.config_path).initial_directory("source"),
            self.source_dir,
        )

    def test_failed_save_does_not_change_remembered_directory(self):
        self.store.remember_file("destination", self.destination_dir / "summary.xlsx")
        with patch.object(self.store, "_save", side_effect=OSError("read only")):
            with self.assertRaises(OSError):
                self.store.remember_file("destination", self.source_dir / "other.xlsx")
        self.assertEqual(self.store.initial_directory("destination"), self.destination_dir)
        self.assertEqual(
            RecentDirectories(self.config_path).initial_directory("destination"),
            self.destination_dir,
        )

    def test_unknown_kind_is_not_saved(self):
        with self.assertRaises(ValueError):
            self.store.remember_file("other", self.source_dir / "x.xlsx")
        with self.assertRaises(ValueError):
            self.store.initial_directory("other")
        self.assertFalse(self.config_path.exists())

    def test_gui_open_file_dialogs_use_remembered_folders(self):
        """Call Tk event handlers without constructing a display-dependent Tk root."""
        destination_path = self.destination_dir / "summary.xlsx"
        source_path = self.source_dir / "report.xlsx"
        session = SimpleNamespace(
            path=destination_path, fields=[object()],
            next_empty_row=lambda: 4,
            clear_previous_imports=Mock(return_value=0),
            open_input=lambda _, **kwargs: SimpleNamespace(row=4, sheets=["ВИСНОВКИ"]),
        )
        app = SimpleNamespace(
            recent_directories=self.store,
            session=None,
            aliases=None,
            destination_var=Mock(),
            source_var=Mock(),
            status_var=Mock(),
            source_button=Mock(),
            clear_rows_var=Mock(get=Mock(return_value=False)),
            _choose_destination_row=Mock(),
            restart_button=Mock(),
            alias_button=Mock(),
            opened_sources=[],
            _clear_tables=Mock(),
            _refresh_tables=Mock(),
            _set_source_sheets=Mock(),
            _refresh_destination_rows=Mock(),
            _remember_directory=lambda kind, path: self.store.remember_file(kind, path),
        )
        with (
            patch("excel_copy_script.gui.filedialog.askopenfilename", return_value=str(destination_path)) as open_dialog,
            patch("excel_copy_script.gui.WorkbookSession", return_value=session),
            patch("excel_copy_script.gui.AliasStore"),
        ):
            ExcelCopyApp.open_destination(app)
            session.clear_previous_imports.assert_not_called()
            self.assertEqual(Path(open_dialog.call_args.kwargs["initialdir"]), Path.home())

        with patch("excel_copy_script.gui.filedialog.askopenfilename", return_value=str(source_path)) as open_dialog:
            ExcelCopyApp.open_source(app)
            self.assertEqual(Path(open_dialog.call_args.kwargs["initialdir"]), self.destination_dir)
        self.assertEqual(
            RecentDirectories(self.config_path).initial_directory("source"),
            self.source_dir,
        )

        self.assertEqual(app.opened_sources, [source_path])

        # A new app session will use the folder from the previous run.
        app.recent_directories = RecentDirectories(self.config_path)
        with patch("excel_copy_script.gui.filedialog.askopenfilename", return_value="") as open_dialog:
            ExcelCopyApp.open_source(app)
            self.assertEqual(Path(open_dialog.call_args.kwargs["initialdir"]), self.source_dir)

        with patch("excel_copy_script.gui.filedialog.askopenfilename", return_value="") as open_dialog:
            ExcelCopyApp.open_destination(app)
            self.assertEqual(Path(open_dialog.call_args.kwargs["initialdir"]), self.destination_dir)

    def test_cancel_or_invalid_file_does_not_update_directory(self):
        app = SimpleNamespace(
            recent_directories=self.store,
            session=SimpleNamespace(path=self.destination_dir / "summary.xlsx"),
            _remember_directory=Mock(),
        )
        with patch("excel_copy_script.gui.filedialog.askopenfilename", return_value=""):
            ExcelCopyApp.open_source(app)
        app._remember_directory.assert_not_called()

        with (
            patch("excel_copy_script.gui.filedialog.askopenfilename", return_value=str(self.source_dir / "bad.xlsx")),
            patch("excel_copy_script.gui.messagebox.showerror"),
        ):
            app.session.open_input = Mock(side_effect=ValueError("missing sheet"))
            ExcelCopyApp.open_source(app)
        app._remember_directory.assert_not_called()
        self.assertFalse(self.config_path.exists())


if __name__ == "__main__":
    unittest.main()
