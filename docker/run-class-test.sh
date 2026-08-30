#!/usr/bin/env bash
# Build the class-module test workbook, then probe it with LibreOffice.
set -euo pipefail

python3 scripts/create_class_test_excel.py
echo
python3 scripts/run_class_test_libreoffice.py "$@"
