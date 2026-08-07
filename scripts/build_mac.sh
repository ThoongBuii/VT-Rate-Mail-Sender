#!/usr/bin/env bash
# Build macOS .app + ZIP sẵn sàng phát hành (xử lý Gatekeeper quarantine)
set -euo pipefail
cd "$(dirname "$0")/.."
python3 -m venv .venv 2>/dev/null || true
# shellcheck disable=SC1091
source .venv/bin/activate
pip install -r requirements.txt -r requirements-build.txt
rm -rf build dist
pyinstaller packaging/vt_rate.spec --noconfirm

APP="dist/VT Rate Mail Sender.app"
ZIP="dist/VT-Rate-Mail-Sender-macOS.zip"

# Ad-hoc sign lại toàn bundle (tránh báo damaged do thiếu chữ ký)
codesign --force --deep --sign - "$APP"
xattr -cr "$APP"

rm -f "$ZIP"
# ditto giữ metadata .app đúng hơn zip thông thường
ditto -c -k --sequesterRsrc --keepParent "$APP" "$ZIP"

echo ""
echo "Xong: $APP"
echo "ZIP:  $ZIP"
echo ""
echo "Người dùng tải từ GitHub nếu bị 'damaged', chạy:"
echo "  xattr -cr \"/path/to/VT Rate Mail Sender.app\""
