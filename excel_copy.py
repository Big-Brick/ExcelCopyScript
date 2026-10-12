"""Backward-compatible launcher for ExcelCopyScript.

Prefer `python -m excel_copy_script` or the installed `excel-copy-script` command.
"""

from excel_copy_script import AliasStore, WorkbookSession, normalize
from excel_copy_script.gui import main

if __name__ == "__main__":
    main()
