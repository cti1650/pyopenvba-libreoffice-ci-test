#!/usr/bin/env python3
"""
Build the VBA class-module test workbook.

Reads every .cls / .bas under vba/class_test/ and adds them to a fresh
.xlsm via pyOpenVBA's VBAProject.add_module(). Unlike ExcelFile.set_module(),
which only replaces existing modules, add_module() can create new standard
and class modules -- class modules only need the universal class CLSID, which
pyOpenVBA synthesizes. (Document and designer/UserForm modules carry
host-specific CLSIDs and cannot be created this way.)
"""

import sys
from pathlib import Path

from pyopenvba import ExcelFile, VBAModuleKind

# Classes are added before standard modules so that a module referencing a
# class type is never resolved against a project that lacks it.
_KIND_BY_SUFFIX = {
    ".cls": VBAModuleKind.other,
    ".bas": VBAModuleKind.standard,
}


def read_source(path: Path) -> str:
    """Read a VBA source file with CRLF line endings, as the VBE writes them."""
    text = path.read_text(encoding="utf-8")
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def build_workbook(src_dir: Path, output_path: Path) -> list[str]:
    """Create output_path containing every module found in src_dir."""
    cls_files = sorted(src_dir.glob("*.cls"))
    bas_files = sorted(src_dir.glob("*.bas"))

    if not cls_files and not bas_files:
        raise FileNotFoundError(f"No .cls/.bas files found in {src_dir}")

    added = []
    with ExcelFile.create_new(str(output_path)) as wb:
        project = wb.vba_project()
        print(f"Template modules: {wb.module_names()}")

        for path in cls_files + bas_files:
            kind = _KIND_BY_SUFFIX[path.suffix.lower()]
            name = path.stem
            try:
                project.add_module(name, read_source(path), kind=kind)
            except ValueError:
                # Name collides with a template module (e.g. Module1) -- replace.
                wb.set_module(name, read_source(path))
                print(f"  replaced {name} ({kind.name})")
            else:
                print(f"  added    {name} ({kind.name})")
            added.append(name)

        # validate() compares the parsed project against the CFB on disk, so
        # it is only meaningful after save() has written the module streams.
        wb.save()

    print(f"Created: {output_path} ({output_path.stat().st_size} bytes)")
    return added


def verify_roundtrip(output_path: Path, expected: list[str]) -> bool:
    """Reopen the workbook and confirm every module survived the save."""
    print(f"\n--- Round-trip check: {output_path.name} ---")
    ok = True
    with ExcelFile(str(output_path)) as wb:
        names = wb.module_names()
        print(f"Modules after reopen: {names}")

        problems = wb.validate()
        if problems:
            ok = False
            print("  VALIDATION PROBLEMS:")
            for p in problems:
                print(f"    {p}")
        else:
            print("  Validation: OK")

        for name in expected:
            if name not in names:
                print(f"  MISSING: {name}")
                ok = False
                continue

            source = wb.get_module(name)
            if "Attribute VB_Name" not in source:
                print(f"  {name}: attribute header missing")
                ok = False
                continue

            # A class module is unusable in Excel without VB_Base; confirm
            # pyOpenVBA inserted it and preserved any VB_PredeclaredId.
            lines = [ln.strip() for ln in source.strip().splitlines()]
            code = [ln for ln in lines if not ln.startswith("'")]
            flags = []
            if "Attribute VB_Base" in source:
                flags.append("VB_Base")
            if "Attribute VB_PredeclaredId = True" in source:
                flags.append("PredeclaredId=True")
            if any(ln.startswith("Implements ") for ln in code):
                flags.append("Implements")
            if any("WithEvents " in ln for ln in code):
                flags.append("WithEvents")
            if any(ln.startswith("Public Event ") for ln in code):
                flags.append("Event")
            body_lines = len(lines)
            suffix = f" [{', '.join(flags)}]" if flags else ""
            print(f"  {name}: {body_lines} lines{suffix}")

    return ok


def main() -> int:
    project_root = Path(__file__).parent.parent
    src_dir = project_root / "vba" / "class_test"
    output_dir = project_root / "output"
    output_dir.mkdir(exist_ok=True)
    output_path = output_dir / "class_test_workbook.xlsm"

    print("=" * 60)
    print("Building VBA class-module test workbook")
    print("=" * 60)

    added = build_workbook(src_dir, output_path)

    if not verify_roundtrip(output_path, added):
        print("\nRound-trip check FAILED")
        return 1

    print("\nRound-trip check OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
