#!/usr/bin/env python3
"""
Create an Excel file with VBA macros using pyOpenVBA.
This script demonstrates how to programmatically add VBA code to Excel files.
"""

import sys
from pathlib import Path

from openpyxl import Workbook
from pyopenvba import ExcelFile


def create_base_excel(output_path: Path) -> None:
    """Create a base Excel file with openpyxl."""
    wb = Workbook()
    ws = wb.active
    ws.title = "TestSheet"
    ws["A1"] = "Test Data"
    ws["A2"] = 10
    ws["B2"] = 20
    # Save as .xlsx first
    xlsx_path = output_path.with_suffix(".xlsx")
    wb.save(xlsx_path)
    print(f"Created base Excel file: {xlsx_path}")
    return xlsx_path


def inject_vba(xlsx_path: Path, vba_code: str, output_path: Path) -> None:
    """Inject VBA code into Excel file and save as .xlsm."""
    # pyOpenVBA can create a new xlsm file
    with ExcelFile.create_new(str(output_path)) as wb:
        wb.set_module("SampleModule", vba_code)
        wb.save()
    print(f"Created Excel file with VBA: {output_path}")


def main():
    # Paths
    project_root = Path(__file__).parent.parent
    output_dir = project_root / "output"
    output_dir.mkdir(exist_ok=True)

    vba_path = project_root / "vba" / "sample_module.bas"
    output_xlsm = output_dir / "test_workbook.xlsm"

    # Read VBA code
    vba_code = vba_path.read_text(encoding="utf-8")

    # Remove Attribute line if present (pyOpenVBA handles this)
    lines = vba_code.split("\n")
    if lines[0].startswith("Attribute VB_Name"):
        vba_code = "\n".join(lines[1:])

    # Create Excel with VBA
    inject_vba(None, vba_code, output_xlsm)

    print("Done!")
    return 0


if __name__ == "__main__":
    sys.exit(main())
