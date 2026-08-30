#!/usr/bin/env python3
"""
Inject VBA code into an existing Excel file using pyOpenVBA.

Handles both standard modules (.bas) and class modules (.cls). A module that
does not exist yet is created via VBAProject.add_module(); ExcelFile.set_module()
alone would raise KeyError, since it only replaces existing modules.

Note that this cannot create document modules (ThisWorkbook, Sheet1) or
designer/UserForm modules -- those carry host-specific CLSIDs that pyOpenVBA
deliberately refuses to invent.
"""

import sys
from pathlib import Path

from pyopenvba import ExcelFile, VBAModuleKind

_KIND_BY_SUFFIX = {
    ".cls": VBAModuleKind.other,
    ".bas": VBAModuleKind.standard,
}


def module_name_from_source(path: Path, source: str) -> str:
    """Prefer the name declared in the attribute header over the file stem."""
    for line in source.splitlines():
        if line.startswith("Attribute VB_Name"):
            return line.split('"')[1]
    return path.stem


def read_source(path: Path) -> str:
    """Read VBA source with CRLF line endings, as the VBE writes them."""
    text = path.read_text(encoding="utf-8")
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def inject_vba(xlsm_path: Path, vba_path: Path) -> str:
    """Inject or update one VBA module. Returns the module name."""
    suffix = vba_path.suffix.lower()
    if suffix not in _KIND_BY_SUFFIX:
        raise ValueError(f"Unsupported VBA source extension: {suffix} (expected .bas or .cls)")

    kind = _KIND_BY_SUFFIX[suffix]
    source = read_source(vba_path)
    name = module_name_from_source(vba_path, source)

    with ExcelFile(str(xlsm_path)) as wb:
        existing = wb.module_names()
        print(f"Existing modules: {existing}")

        if any(m.casefold() == name.casefold() for m in existing):
            # set_module normalizes .cls export form and preserves the
            # module's existing VB_Base (important for document modules).
            wb.set_module(name, source)
            print(f"Updated module '{name}' ({kind.name})")
        else:
            wb.vba_project().add_module(name, source, kind=kind)
            print(f"Added module '{name}' ({kind.name})")

        wb.save()

    return name


def main():
    if len(sys.argv) < 3:
        print("Usage: python inject_vba.py <xlsm_file> <vba_file> [<vba_file> ...]")
        return 1

    xlsm_path = Path(sys.argv[1])
    if not xlsm_path.exists():
        print(f"Excel file not found: {xlsm_path}")
        return 1

    vba_paths = [Path(p) for p in sys.argv[2:]]
    missing = [p for p in vba_paths if not p.exists()]
    if missing:
        for p in missing:
            print(f"VBA file not found: {p}")
        return 1

    # Class modules first, so a standard module referencing a class type is
    # never written into a project that does not contain it yet.
    vba_paths.sort(key=lambda p: 0 if p.suffix.lower() == ".cls" else 1)

    for vba_path in vba_paths:
        inject_vba(xlsm_path, vba_path)

    print("Done!")
    return 0


if __name__ == "__main__":
    sys.exit(main())
