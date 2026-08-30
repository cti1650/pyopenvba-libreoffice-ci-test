#!/usr/bin/env python3
"""
Extract VBA code from an Excel file using pyOpenVBA.

Standard modules are written as .bas and class/document/designer modules as
.cls, matching what the VBE's own export produces. UserForm layout (.frx) is
not exported -- pyOpenVBA keeps it inside the file and never round-trips it
through disk, so an extracted UserForm carries its code but not its controls.
"""

import sys
from pathlib import Path

from pyopenvba import ExcelFile, VBAModuleKind


def extract_vba(xlsm_path: Path, output_dir: Path) -> None:
    """Extract all VBA modules from an Excel file."""
    output_dir.mkdir(parents=True, exist_ok=True)

    with ExcelFile(str(xlsm_path)) as wb:
        modules = wb.vba_project().modules
        print(f"Found {len(modules)} modules: {[m.name for m in modules]}")

        for module in modules:
            ext = ".bas" if module.kind == VBAModuleKind.standard else ".cls"
            output_file = output_dir / f"{module.name}{ext}"
            output_file.write_text(wb.get_module(module.name), encoding="utf-8")
            print(f"Extracted: {output_file} ({module.kind.name})")


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
