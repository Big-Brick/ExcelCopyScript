"""Real Tk dialogs for creating a new, empty profile without copying aliases."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from excel_copy_script.aliases import AliasStore
from excel_copy_script.gui import ExcelCopyApp
from excel_copy_script.models import DestinationField, SourceField
from excel_copy_script.new_profile_dialog import NewProfileDialog, suggested_profile_path
from excel_copy_script.profiles import Profile, ProfileManager
from excel_copy_script.workbook_layout import WorkbookLayout


class NewProfileGuiTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.folder = Path(temp.name)
        self.manager = ProfileManager(self.folder / "config")
        profile_patch = patch(
            "excel_copy_script.gui.ProfileManager", return_value=self.manager,
        )
        profile_patch.start()
        self.addCleanup(profile_patch.stop)
        self.app = ExcelCopyApp()
        self.addCleanup(self.app.destroy)
        self.app.geometry("1300x860")
        self.app.update()

    def dialog(self):
        dialog = NewProfileDialog(self.app, self.manager.config_dir / "profiles")
        self.addCleanup(lambda: dialog.winfo_exists() and dialog.destroy())
        self.app.update()
        return dialog

    def test_new_button_creates_blank_profile_and_activates_it(self):
        previous = self.app.profile
        store = AliasStore.from_profile(previous)
        store.set(
            "old-template", SourceField(14, "Old", 42),
            DestinationField(2, "Old Target"),
        )

        new_path = self.folder / "my-blank-profile.json"
        def fill_modal():
            modal = next(
                child for child in self.app.winfo_children()
                if isinstance(child, NewProfileDialog)
            )
            self.assertEqual(modal.fields["destination_sheet"].get(), "Дані_підрозділів")
            modal.path_var.set(str(new_path))
            modal.fields["destination_sheet"].set("Звіт")
            modal.fields["header_row"].set("5")
            modal.fields["first_data_row"].set("7")
            modal.fields["source_name_column"].set("4")
            modal.fields["source_value_first_column"].set("6")
            modal.fields["source_value_last_column"].set("8")
            modal.create_button.invoke()

        # Invoke the actual visible button. The modal wait must finish only
        # after the user action creates the profile.
        bars = [frame for frame in self.app.winfo_children()
                if any(getattr(child, "cget", None)
                       and child.winfo_class() == "TButton"
                       for child in frame.winfo_children())]
        buttons = [widget for bar in bars for widget in bar.winfo_children()
                   if widget.winfo_class() == "TButton"
                   and widget.cget("text") == "Новий профіль…"]
        self.assertEqual(len(buttons), 1)
        self.app.after(20, fill_modal)
        buttons[0].invoke()
        self.app.update()

        current = self.app.profile
        self.assertEqual(current.path, new_path.resolve())
        self.assertEqual(current.layout.destination_sheet, "Звіт")
        self.assertEqual(current.layout.header_row, 5)
        self.assertEqual(current.layout.first_data_row, 7)
        self.assertEqual(current.aliases_data["templates"], {})
        self.assertEqual(self.manager.load_last().path, new_path.resolve())
        self.assertEqual(tuple(self.app.create_sources.values_tree["columns"]),
                         ("F", "G", "H"))
        self.assertIn("old-template", Profile.load(previous.path).aliases_data["templates"])
        self.assertEqual(Profile.load(new_path).aliases_data["templates"], {})

    def test_cancellation_does_not_create_file_or_switch_profile(self):
        before = self.app.profile
        suggested = suggested_profile_path(self.manager.config_dir / "profiles")
        def cancel():
            modal = next(child for child in self.app.winfo_children()
                         if isinstance(child, NewProfileDialog))
            modal.destroy()
        self.app.after(20, cancel)
        self.app.create_profile()
        self.assertIs(self.app.profile, before)
        self.assertFalse(suggested.exists())
        self.assertFalse(self.manager.state_path.exists())

    def test_original_hardcoded_defaults_even_with_different_active_profile(self):
        # The new-profile dialog must NOT inherit coordinates or sheet names
        # from whichever profile happens to be active.
        altered = self.app.profile.copy_as(self.folder / "altered.json")
        altered.layout = WorkbookLayout(
            destination_sheet="Another sheet", header_row=9, first_data_row=12,
            first_destination_column=5, source_name_column=3,
            first_source_row=4, source_value_first_column=6,
            source_value_last_column=9, preferred_source_sheet="Different",
        )
        altered.save()
        self.app._activate_profile(altered)
        dialog = self.dialog()
        expected = {
            "destination_sheet": "Дані_підрозділів",
            "header_row": "3",
            "first_data_row": "4",
            "first_destination_column": "2",
            "source_name_column": "2",
            "first_source_row": "1",
            "source_value_first_column": "3",
            "source_value_last_column": "7",
            "preferred_source_sheet": "ВИСНОВКИ",
        }
        self.assertEqual(
            {key: value.get() for key, value in dialog.fields.items()},
            expected,
        )
        dialog.destroy()

    def test_invalid_settings_report_error_and_keep_dialog_open(self):
        dialog = self.dialog()
        target = self.folder / "invalid-test.json"
        dialog.path_var.set(str(target))
        dialog.fields["first_data_row"].set("2")
        with patch("excel_copy_script.new_profile_dialog.messagebox.showerror") as error:
            dialog.create_button.invoke()
        error.assert_called_once()
        self.assertIn("Перший рядок даних", str(error.call_args))
        self.assertTrue(dialog.winfo_exists())
        self.assertFalse(target.exists())
        self.assertIsNone(dialog.profile)

        dialog.fields["first_data_row"].set("4")
        dialog.fields["source_value_first_column"].set("oops")
        with patch("excel_copy_script.new_profile_dialog.messagebox.showerror") as error:
            dialog.create_button.invoke()
        self.assertIn("Перша колонка значення", str(error.call_args))
        self.assertTrue(dialog.winfo_exists())
        self.assertFalse(target.exists())

    def test_existent_profile_file_is_never_overwritten(self):
        existing = self.folder / "do-not-touch.json"
        existing.write_bytes(b"do not replace")
        dialog = self.dialog()
        dialog.path_var.set(str(existing))
        with patch("excel_copy_script.new_profile_dialog.messagebox.showerror") as error:
            dialog.create_button.invoke()
        error.assert_called_once()
        self.assertTrue(dialog.winfo_exists())
        self.assertEqual(existing.read_bytes(), b"do not replace")
        self.assertIsNone(dialog.profile)

        # The dialog remains usable: select another name, create successfully.
        dialog.path_var.set(str(self.folder / "unique.json"))
        dialog.create_button.invoke()
        self.assertEqual(dialog.profile.path, self.folder / "unique.json")
        self.assertFalse(dialog.winfo_exists())

    def test_default_filename_skips_existing_files(self):
        target = self.manager.config_dir / "profiles"
        target.mkdir(parents=True, exist_ok=True)
        (target / "new-profile.json").write_text("existing")
        (target / "new-profile-2.json").write_text("existing")
        dialog = self.dialog()
        self.assertEqual(Path(dialog.path_var.get()),
                         target / "new-profile-3.json")
        dialog.destroy()

    def test_browse_updates_filename_without_creating_profile(self):
        dialog = self.dialog()
        target = self.folder / "picked.json"
        with patch(
            "excel_copy_script.new_profile_dialog.filedialog.asksaveasfilename",
            return_value=str(target),
        ) as browser:
            dialog._browse_path()
        browser.assert_called_once()
        self.assertEqual(dialog.path_var.get(), str(target))
        self.assertFalse(target.exists())
        dialog.destroy()


if __name__ == "__main__":
    unittest.main()
