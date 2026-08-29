#!/usr/bin/env python3
"""
Inject VBA code into an existing Excel file using pyOpenVBA.
"""

import sys
from pathlib import Path

from pyopenvba import ExcelFile


def inject_vba(xlsm_path: Path, module_name: str, vba_code: str) -> None:
    """Inject or update VBA module in an Excel file."""
    with ExcelFile(str(xlsm_path)) as wb:
        existing_modules = wb.module_names()
        print(f"Existing modules: {existing_modules}")

        # Set/update the module
        wb.set_module(module_name, vba_code)
        wb.save()
        print(f"Injected module '{module_name}' into {xlsm_path}")


def main():
    if len(sys.argv) < 3:
        print("Usage: python inject_vba.py <xlsm_file> <vba_file>")
        return 1

    xlsm_path = Path(sys.argv[1])
    vba_path = Path(sys.argv[2])

    if not xlsm_path.exists():
        print(f"Excel file not found: {xlsm_path}")
        return 1

    if not vba_path.exists():
        print(f"VBA file not found: {vba_path}")
        return 1

    # Read VBA code
    vba_code = vba_path.read_text(encoding="utf-8")

    # Get module name from file or Attribute line
    module_name = vba_path.stem
    lines = vba_code.split("\n")
    if lines[0].startswith("Attribute VB_Name"):
        # Extract module name from Attribute line
        module_name = lines[0].split('"')[1]
        vba_code = "\n".join(lines[1:])

    inject_vba(xlsm_path, module_name, vba_code)
    print("Done!")
    return 0


if __name__ == "__main__":
    sys.exit(main())
