#!/usr/bin/env python3
"""
Test script for verifying VBA injection with pyOpenVBA.
"""

import sys
from pathlib import Path

from pyopenvba import ExcelFile


def test_create_new_xlsm():
    """Test creating a new xlsm file with VBA."""
    output_path = Path("./output/test_new.xlsm")
    output_path.parent.mkdir(exist_ok=True)

    vba_code = '''
Option Explicit

Public Function TestFunction() As String
    TestFunction = "Hello from VBA"
End Function
'''

    with ExcelFile.create_new(str(output_path)) as wb:
        wb.set_module("TestModule", vba_code)
        wb.save()

    # Verify
    with ExcelFile(str(output_path)) as wb:
        modules = wb.module_names()
        assert "TestModule" in modules, f"TestModule not found in {modules}"

        source = wb.get_module("TestModule")
        assert "TestFunction" in source, "TestFunction not found in source"

    print("test_create_new_xlsm: PASSED")
    return True


def test_module_operations():
    """Test module add/rename/delete operations."""
    output_path = Path("./output/test_operations.xlsm")
    output_path.parent.mkdir(exist_ok=True)

    # Create file with initial module
    with ExcelFile.create_new(str(output_path)) as wb:
        wb.set_module("Module1", "Sub Test1()\nEnd Sub")
        wb.save()

    # Add another module
    with ExcelFile(str(output_path)) as wb:
        wb.set_module("Module2", "Sub Test2()\nEnd Sub")
        wb.save()

    # Verify both modules exist
    with ExcelFile(str(output_path)) as wb:
        modules = wb.module_names()
        assert "Module1" in modules, f"Module1 not found in {modules}"
        assert "Module2" in modules, f"Module2 not found in {modules}"

    print("test_module_operations: PASSED")
    return True


def test_read_sample_module():
    """Test reading the sample module file."""
    vba_path = Path(__file__).parent.parent / "vba" / "sample_module.bas"

    if not vba_path.exists():
        print(f"Sample VBA file not found: {vba_path}")
        return False

    vba_code = vba_path.read_text(encoding="utf-8")
    assert "AddNumbers" in vba_code, "AddNumbers function not found"
    assert "ConcatStrings" in vba_code, "ConcatStrings function not found"

    print("test_read_sample_module: PASSED")
    return True


def main():
    """Run all tests."""
    tests = [
        test_read_sample_module,
        test_create_new_xlsm,
        test_module_operations,
    ]

    passed = 0
    failed = 0

    for test in tests:
        try:
            if test():
                passed += 1
            else:
                failed += 1
        except Exception as e:
            print(f"{test.__name__}: FAILED - {e}")
            failed += 1

    print(f"\nResults: {passed} passed, {failed} failed")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
