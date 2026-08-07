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
1. Tải `VTRateMailSender.exe` (hoặc ZIP) từ [Releases](https://github.com/ThoongBuii/VT-Rate-Mail-Sender/releases)
2. **Một file `.exe` độc lập** — copy đi đâu cũng được (Desktop, USB, …), không cần thư mục `_internal`
3. Chuột phải → Pin to taskbar (pin từ `.exe`, không pin từ Chrome)
4. Cần **Outlook desktop** đã login + **Edge WebView2** (Win10/11 thường sẵn). Thiếu thì cài: https://go.microsoft.com/fwlink/p/?LinkId=2124703

> Đây là **desktop app** (cửa sổ riêng của `.exe`). Localhost bên trong chỉ là engine — không phải website.

> Nếu **Outlook không mở / vẫn còn process cũ:** Task Manager → End task `VTRateMailSender.exe` / `OUTLOOK.EXE` hoặc `scripts\stop-app-windows.bat`.
> Log: `%APPDATA%\VTRateMailSender\logs\startup.log`

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
