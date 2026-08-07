# VT Rate Mail Sender

Desktop app nội bộ — gửi mail báo giá qua **Outlook đã đăng nhập** (Semi-Auto).

## Dành cho nhân viên văn phòng

### macOS
1. Cài **Microsoft Outlook** và đăng nhập account gửi (ví dụ `overseas@…`)
2. Mở `VT Rate Mail Sender.app`
3. Bấm **Mở Outlook** → soạn/dán nội dung → Import Excel → **Preview** → Semi-Auto

> Lần đầu macOS có thể hỏi quyền **Automation** (điều khiển Outlook) → Allow.

### Windows
1. Cài Outlook desktop + đăng nhập
2. Chạy `VTRateMailSender.exe` (trong thư mục đã giải nén)
3. Cài thêm `pywin32` chỉ khi chạy từ source; bản `.exe` đã đóng gói sẵn

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
# → dist/VT Rate Mail Sender.app
```

**Windows** (trên máy Windows):

```bat
scripts\build_win.bat
# → dist\VTRateMailSender\
```

Zip file `.app` / thư mục `.exe` rồi gửi nội bộ hoặc đính Release trên GitHub.

## Lưu ý vận hành

- App chỉ chạy **localhost** — không mở ra internet
- Không lưu mật khẩu mailbox (dùng Outlook đã login)
- Mỗi agency = 1 email riêng · delay 10–20s · Pause/Stop
- Dữ liệu user (config, log, attachment) lưu tại:
  - macOS: `~/Library/Application Support/VTRateMailSender/`
  - Windows: `%APPDATA%\VTRateMailSender\`
  - Dev: thư mục project

## GitHub

Repo public — dùng Issues/Releases để phát hành bản build cho văn phòng.
