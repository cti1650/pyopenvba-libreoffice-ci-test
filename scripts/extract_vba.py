#!/usr/bin/env python3
"""
Extract VBA code from an Excel file using pyOpenVBA.
"""

import sys
from pathlib import Path

from pyopenvba import ExcelFile


def extract_vba(xlsm_path: Path, output_dir: Path) -> None:
    """Extract all VBA modules from an Excel file."""
    output_dir.mkdir(exist_ok=True)

    with ExcelFile(str(xlsm_path)) as wb:
        modules = wb.module_names()
        print(f"Found {len(modules)} modules: {modules}")

        for module_name in modules:
            source = wb.get_module(module_name)
            output_file = output_dir / f"{module_name}.bas"
            output_file.write_text(source, encoding="utf-8")
            print(f"Extracted: {output_file}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_vba.py <xlsm_file> [output_dir]")
        return 1

    xlsm_path = Path(sys.argv[1])
    output_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("./extracted_vba")

    if not xlsm_path.exists():
        print(f"File not found: {xlsm_path}")
        return 1

    extract_vba(xlsm_path, output_dir)
    print("Done!")
    return 0


if __name__ == "__main__":
    sys.exit(main())
