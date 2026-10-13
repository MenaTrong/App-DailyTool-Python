#!/bin/bash
# Đóng DailyTool.exe cho Windows. Chạy trên macOS hoặc Linux, cần Docker.
set -euo pipefail
cd "$(dirname "$0")"

STAGE="$(mktemp -d)"
cleanup() { rm -rf "$STAGE"; }
trap cleanup EXIT

mkdir -p "$STAGE/dailytool" dist
cp run_dailytool.py xuat_ngay_nghi.py DailyTool.spec requirements.txt "$STAGE/"
cp dailytool/*.py dailytool/*.json "$STAGE/dailytool/"

echo "Đang dựng DailyTool.exe cho Windows..."
docker run --rm \
  -v "$STAGE":/work \
  -w /work \
  -e WINEDEBUG=-all \
  tobix/pywine:3.12 \
  sh -c "wine python -m pip install --disable-pip-version-check -r requirements.txt && wine pyinstaller --noconfirm --clean --distpath dist --workpath /tmp/pyi-work DailyTool.spec && chown -R $(id -u):$(id -g) dist"

cp "$STAGE/dist/DailyTool.exe" dist/DailyTool.exe
echo "Xong: dist/DailyTool.exe"
