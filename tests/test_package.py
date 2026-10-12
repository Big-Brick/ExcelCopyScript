"""Smoke tests for module boundaries and backwards-compatible imports."""

import unittest


class PackageTests(unittest.TestCase):
    def test_backwards_compatible_imports(self):
        from excel_copy import AliasStore as LegacyAliases
        from excel_copy import WorkbookSession as LegacySession
        from excel_copy import normalize as legacy_normalize
        from excel_copy_script.aliases import AliasStore
        from excel_copy_script.workbook import WorkbookSession
        from excel_copy_script.utils import normalize

        self.assertIs(LegacyAliases, AliasStore)
        self.assertIs(LegacySession, WorkbookSession)
        self.assertIs(legacy_normalize, normalize)

    def test_gui_entry_points_import_without_creating_window(self):
        from excel_copy_script.alias_dialog import AliasManagerDialog
        from excel_copy_script.gui import ExcelCopyApp, main
        from excel_copy_script import __main__ as package_main

        self.assertTrue(callable(main))
        self.assertIs(package_main.main, main)
        self.assertTrue(issubclass(ExcelCopyApp, __import__("tkinter").Tk))
        self.assertTrue(issubclass(AliasManagerDialog, __import__("tkinter").Toplevel))


if __name__ == "__main__":
    unittest.main()
