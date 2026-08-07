# VT Rate Mail Sender

**Desktop app** nội bộ (cửa sổ riêng) — gửi mail báo giá qua **Outlook đã đăng nhập** (Semi-Auto).

> Bên trong app dùng engine cục bộ `127.0.0.1` — **không phải** website public. Không cần mở Chrome/Edge thủ công.

## Dành cho nhân viên văn phòng

### macOS
1. Tải `VT-Rate-Mail-Sender-macOS.zip` từ [Releases](https://github.com/ThoongBuii/VT-Rate-Mail-Sender/releases)
2. Giải nén → kéo `VT Rate Mail Sender.app` vào **Applications**
3. Nếu macOS báo *“damaged” / không mở được* (do tải từ internet, app chưa Apple notarize), chạy Terminal:

```bash
xattr -cr "/Applications/VT Rate Mail Sender.app"
```

4. Mở app → Outlook đã login → Allow **Automation** nếu được hỏi

### Windows
1. Tải `VT-Rate-Mail-Sender-Windows.zip` từ [Releases](https://github.com/ThoongBuii/VT-Rate-Mail-Sender/releases)
2. (Khuyến nghị) Chuột phải ZIP → Properties → **Unblock** → OK
3. Giải nén → chạy `VTRateMailSender.exe` (giữ nguyên cả thư mục)
4. Cần Outlook desktop đã login + Microsoft Edge/WebView2

> **Không xóa được thư mục / Outlook không mở:** đóng tab trình duyệt **không** đủ — app vẫn chạy nền.  
> Task Manager → End task → `VTRateMailSender.exe` và `OUTLOOK.EXE`,  
> hoặc chạy `scripts\stop-app-windows.bat` (trong repo / kèm bản phát hành).  
> App ưu tiên cửa sổ desktop; nếu thiếu WebView2 sẽ mở **Edge/Chrome dạng --app** (không thanh địa chỉ). Log: `%APPDATA%\VTRateMailSender\logs\startup.log`

### Excel danh bạ (4 cột)
`Agency Company` · `Account Name` · `Account Mail` · `Mail cc`

Mẫu: `samples/VT_Rate_TEST.xlsx`

---

## Chạy từ source (dev)

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python run.py
```

Windows thêm: `pip install pywin32`

## Đóng gói desktop

**macOS** (trên máy Mac):

```bash
chmod +x scripts/build_mac.sh
./scripts/build_mac.sh
# → dist/VT Rate Mail Sender.app + ZIP
```

**Windows** (CI hoặc máy Windows):

```bat
scripts\build_win.bat
# → dist\VTRateMailSender\
```

## Lưu ý vận hành

- App chỉ chạy **localhost nội bộ** trong cửa sổ desktop — không mở ra internet
- Không lưu mật khẩu mailbox (dùng Outlook đã login)
- Mỗi agency = 1 email riêng · delay 10–20s · Pause/Stop
- Dữ liệu user (config, log, attachment) lưu tại:
  - macOS: `~/Library/Application Support/VTRateMailSender/`
  - Windows: `%APPDATA%\VTRateMailSender\`
  - Dev: thư mục project

## GitHub

Repo: [ThoongBuii/VT-Rate-Mail-Sender](https://github.com/ThoongBuii/VT-Rate-Mail-Sender) — tải bản build tại Releases.
