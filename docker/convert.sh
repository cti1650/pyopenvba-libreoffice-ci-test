#!/usr/bin/env bash
# Convert every .xlsm in output/ to the requested format (default: ods).
# Usage: docker compose run --rm convert [format]
set -euo pipefail

FORMAT="${1:-ods}"

shopt -s nullglob
files=(output/*.xlsm)
if [ ${#files[@]} -eq 0 ]; then
    echo "No .xlsm files in output/. Run 'docker compose run --rm build' first."
    exit 1
fi

for f in "${files[@]}"; do
    echo "--- $f -> $FORMAT ---"
    timeout 120 soffice \
        -env:UserInstallation=file:///tmp/lo_convert_profile \
        --headless --norestore --nologo \
        --convert-to "$FORMAT" --outdir output "$f"
done

echo
ls -la output/
