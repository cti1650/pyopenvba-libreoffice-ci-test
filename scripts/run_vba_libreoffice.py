#!/usr/bin/env python3
"""
Run VBA macros in LibreOffice via UNO API.
This script demonstrates how to execute VBA code in LibreOffice from Python.
"""

import subprocess
import sys
import time
from pathlib import Path

from pyopenvba import ExcelFile


def find_libreoffice():
    """Find LibreOffice executable path."""
    # Windows paths
    windows_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    # macOS paths
    macos_paths = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]

    # Linux paths
    linux_paths = [
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
    ]

    if sys.platform == "win32":
        paths = windows_paths
    elif sys.platform == "darwin":
        paths = macos_paths
    else:
        paths = linux_paths

    for path in paths:
        if Path(path).exists():
            return path

    # Try to find in PATH
    try:
        result = subprocess.run(
            ["where" if sys.platform == "win32" else "which", "soffice"],
            capture_output=True,
            text=True,
        )
        if result.returncode == 0:
            return result.stdout.strip().split("\n")[0]
    except Exception:
        pass

    return None


def convert_xlsm_to_ods(xlsm_path: Path, output_dir: Path) -> Path:
    """Convert Excel file to ODS format using LibreOffice."""
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError("LibreOffice not found")

    output_dir.mkdir(exist_ok=True)

    cmd = [
        soffice,
        "--headless",
        "--convert-to",
        "ods",
        "--outdir",
        str(output_dir),
        str(xlsm_path),
    ]

    print(f"Running: {' '.join(cmd)}")
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    except subprocess.TimeoutExpired:
        print("Conversion timed out after 30 seconds")
        raise RuntimeError("Conversion timed out")

    if result.returncode != 0:
        print(f"STDERR: {result.stderr}")
        raise RuntimeError(f"Conversion failed: {result.returncode}")

    ods_path = output_dir / xlsm_path.with_suffix(".ods").name
    print(f"Converted to: {ods_path}")
    return ods_path


def run_macro_in_libreoffice(file_path: Path, macro_name: str) -> str:
    """
    Run a macro in LibreOffice.
    Note: This is a simplified approach using command line.
    For full VBA execution, UNO API connection is recommended.
    """
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError("LibreOffice not found")

    # LibreOffice macro execution syntax:
    # macro:///Library.Module.Function
    # For document-embedded macros: macro://document/Standard.Module.Function

    cmd = [
        soffice,
        "--headless",
        "--nofirststartwizard",
        "--norestore",
        str(file_path),
        f"macro:///Standard.{macro_name}",
    ]

    print(f"Running macro: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)

    return result.stdout


def verify_libreoffice():
    """Verify LibreOffice installation and version."""
    soffice = find_libreoffice()
    if not soffice:
        print("LibreOffice not found!")
        return False

    print(f"LibreOffice found at: {soffice}")

    cmd = [soffice, "--version"]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
        print(f"Version: {result.stdout.strip()}")
    except subprocess.TimeoutExpired:
        print("Warning: --version command timed out, but LibreOffice exists")

    return True


def verify_vba_in_file(xlsm_path: Path) -> None:
    """Verify VBA code was injected and display summary."""
    print(f"\n--- Verifying: {xlsm_path.name} ---")

    if not xlsm_path.exists():
        print(f"  ERROR: File not found")
        return

    print(f"  File size: {xlsm_path.stat().st_size} bytes")

    try:
        with ExcelFile(str(xlsm_path)) as wb:
            modules = wb.module_names()
            print(f"  Modules: {modules}")

            for module_name in modules:
                try:
                    source = wb.get_module(module_name)
                    lines = source.strip().split("\n")
                    print(f"\n  [{module_name}] ({len(lines)} lines)")

                    # Check for Win32 API declarations
                    declare_count = sum(1 for line in lines if "Declare" in line)
                    function_count = sum(1 for line in lines if line.strip().startswith(("Public Function", "Public Sub", "Private Function", "Private Sub")))

                    if declare_count > 0:
                        print(f"    Win32 API Declarations: {declare_count}")
                    if function_count > 0:
                        print(f"    Functions/Subs: {function_count}")

                    # Show first few lines
                    print("    Preview:")
                    for line in lines[:5]:
                        if line.strip():
                            print(f"      {line[:60]}{'...' if len(line) > 60 else ''}")
                    if len(lines) > 5:
                        print(f"      ... ({len(lines) - 5} more lines)")

                except Exception as e:
                    print(f"    Could not read: {e}")

    except Exception as e:
        print(f"  ERROR reading file: {e}")


def main():
    print("=== LibreOffice VBA Test ===\n")
    print(f"Platform: {sys.platform}")

    # Verify LibreOffice installation
    if not verify_libreoffice():
        print("\nPlease install LibreOffice to run VBA tests.")
        return 1

    # Check if test file exists
    project_root = Path(__file__).parent.parent
    test_xlsm = project_root / "output" / "test_workbook.xlsm"

    if not test_xlsm.exists():
        print(f"\nTest file not found: {test_xlsm}")
        print("Run 'python scripts/create_excel_with_vba.py' first.")
        return 1

    print(f"\nTest file found: {test_xlsm}")
    print(f"File size: {test_xlsm.stat().st_size} bytes")

    # Verify VBA in all xlsm files
    print("\n" + "=" * 50)
    print("VBA Verification Results")
    print("=" * 50)
    output_dir = project_root / "output"
    for xlsm_file in output_dir.glob("*.xlsm"):
        verify_vba_in_file(xlsm_file)

    # Convert to ODS (LibreOffice native format)
    print(f"\nConverting {test_xlsm} to ODS format...")
    try:
        ods_path = convert_xlsm_to_ods(test_xlsm, project_root / "output")
        if ods_path.exists():
            print(f"Successfully converted to: {ods_path}")
            print(f"ODS file size: {ods_path.stat().st_size} bytes")
        else:
            print(f"Warning: ODS file not created at {ods_path}")
    except Exception as e:
        print(f"Conversion failed: {e}")
        # Don't return error - this is expected to fail in some CI environments
        print("Note: Conversion may fail in headless CI environments")

    # List output directory contents
    output_dir = project_root / "output"
    print(f"\nOutput directory contents:")
    for f in output_dir.iterdir():
        print(f"  {f.name} ({f.stat().st_size} bytes)")

    print("\n=== Test Complete ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
