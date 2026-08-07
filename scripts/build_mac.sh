#!/usr/bin/env bash
# Build macOS .app
set -euo pipefail
cd "$(dirname "$0")/.."
python3 -m venv .venv 2>/dev/null || true
source .venv/bin/activate
pip install -r requirements.txt -r requirements-build.txt
rm -rf build dist
pyinstaller packaging/vt_rate.spec --noconfirm
echo ""
echo "Xong: dist/VT Rate Mail Sender.app"
echo "Copy vào /Applications hoặc gửi ZIP cho văn phòng."
