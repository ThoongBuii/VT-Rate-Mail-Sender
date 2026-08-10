"""macOS Legacy Outlook — chữ ký gốc (còn ảnh) + dán body phía trên → Send.

Không set content (tránh gãy CID ảnh chữ ký).
"""

from __future__ import annotations

import re
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

from ..models import AppConfig
from ..outlook_html import prepare_body_html_for_outlook
from .html_merge import html_visible_text


def _esc(s: str) -> str:
    return (s or "").replace("\\", "\\\\").replace('"', '\\"')


def _account_block(from_email: str) -> str:
    if not from_email:
        return ""
    fe = _esc(from_email)
    return (
        "\n"
        "  try\n"
        "    repeat with acc in (get exchange accounts)\n"
        "      try\n"
        f'        if (email address of acc as string) contains "{fe}" then\n'
        "          set account of msg to acc\n"
        "          exit repeat\n"
        "        end if\n"
        "      end try\n"
        "    end repeat\n"
        "  end try\n"
        "  try\n"
        "    repeat with acc in (get imap accounts)\n"
        "      try\n"
        f'        if (email address of acc as string) contains "{fe}" then\n'
        "          set account of msg to acc\n"
        "          exit repeat\n"
        "        end if\n"
        "      end try\n"
        "    end repeat\n"
        "  end try\n"
    )


def _run_osascript(script: str, timeout: int = 180) -> str:
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "vt_mac.applescript"
        path.write_text(script, encoding="utf-8")
        result = subprocess.run(
            ["osascript", str(path)],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    if result.returncode != 0:
        raise RuntimeError((result.stderr or result.stdout or "osascript failed").strip())
    return (result.stdout or "").strip()


def _set_html_clipboard(html: str) -> None:
    try:
        from AppKit import NSPasteboard, NSPasteboardTypeHTML, NSPasteboardTypeString  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Thiếu PyObjC (AppKit). Cài: pip install pyobjc-framework-Cocoa"
        ) from exc

    plain = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", html or "")
    plain = re.sub(r"(?is)<br\s*/?>", "\n", plain)
    plain = re.sub(r"(?is)</p>", "\n", plain)
    plain = re.sub(r"<[^>]+>", " ", plain)
    plain = re.sub(r"\s+", " ", plain).strip() or " "

    pb = NSPasteboard.generalPasteboard()
    pb.clearContents()
    ok_html = pb.setString_forType_(html, NSPasteboardTypeHTML)
    ok_plain = pb.setString_forType_(plain, NSPasteboardTypeString)
    if not ok_html and not ok_plain:
        raise RuntimeError("Không ghi được clipboard HTML.")


class MacOutlookSender:
    def __init__(self, config: AppConfig):
        self.config = config

    def open_outlook(self) -> str:
        subprocess.run(["open", "-a", "Microsoft Outlook"], check=False)
        result = subprocess.run(
            ["osascript", "-l", "JavaScript", "-e", 'Application("Microsoft Outlook").name();'],
            capture_output=True,
            text=True,
            timeout=60,
        )
        if result.returncode != 0:
            err = (result.stderr or result.stdout or "").strip()
            raise RuntimeError(
                "Không điều khiển được Microsoft Outlook trên Mac.\n"
                "Cần Legacy Outlook ON + đã đăng nhập.\n"
                f"Chi tiết: {err or 'osascript failed'}"
            )
        return self.config.from_email or "Outlook (Mac)"

    def signature_status(self) -> dict:
        return {
            "ready": True,
            "path": "",
            "message": (
                "macOS: mở New Mail (chữ ký gốc còn ảnh) → dán body phía trên → Send."
            ),
        }

    def capture_signature(self) -> str:
        return "Không cần chụp chữ ký — giữ chữ ký gốc Outlook (có ảnh)."

    def send(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Path],
    ) -> None:
        """
        1) To/Cc/Subject + mở New Mail (chữ ký Outlook nguyên ảnh)
        2) Clipboard HTML body → Tab tới body → Cmd+V (không set content)
        3) Send cùng mail
        """
        prepared = prepare_body_html_for_outlook(body_html)
        if not (prepared or "").strip():
            raise ValueError("Nội dung mail trống — không gửi.")
        if not to_list:
            raise ValueError("Thiếu địa chỉ To — không gửi.")

        probe_raw = re.sub(r"\s+", " ", html_visible_text(prepared)).strip()
        probe = "Dear"
        for token in re.findall(r"[A-Za-zÀ-ỹ]{3,}", probe_raw or ""):
            probe = token[:12]
            break

        to_block = "\n".join(
            f'  make new to recipient at msg with properties {{email address:{{address:"{_esc(t)}"}}}}'
            for t in to_list
        )
        cc_block = "\n".join(
            f'  make new cc recipient at msg with properties {{email address:{{address:"{_esc(cc)}"}}}}'
            for cc in cc_list
        )
        if attachment:
            att = _esc(str(attachment.resolve()))
            att_block = (
                f'  make new attachment at msg with properties {{file:POSIX file "{att}"}}'
            )
        else:
            att_block = ""

        _set_html_clipboard(prepared)

        script = f'''
tell application "Microsoft Outlook"
  activate
  set msg to make new outgoing message
{_account_block(self.config.from_email)}
  set subject of msg to "{_esc(subject)}"
{to_block}
{cc_block}
{att_block}

  -- Chữ ký account (Outlook gắn kèm ảnh, không rewrite HTML)
  try
    set accObj to account of msg
    set sigs to signatures of accObj
    if (count of sigs) > 0 then set signature of msg to item 1 of sigs
  end try
  try
    if (count of signatures) > 0 then set signature of msg to item 1 of signatures
  end try

  open msg
  delay 1.8
end tell

tell application "System Events"
  if not UI elements enabled then
    error "Bật Accessibility: System Settings → Privacy & Security → Accessibility → VT Rate Mail Sender"
  end if
  tell process "Microsoft Outlook"
    set frontmost to true
    delay 0.4
    -- Layout Legacy: To → Cc → Subject → Body
    keystroke tab
    delay 0.12
    keystroke tab
    delay 0.12
    keystroke tab
    delay 0.25
    -- Caret đầu body (trên chữ ký)
    key code 126 using {{command down}}
    delay 0.15
    keystroke "v" using {{command down}}
    delay 1.2
  end tell
end tell

tell application "Microsoft Outlook"
  set checkText to ""
  try
    set checkText to plain text content of msg
  end try
  if checkText is "" then
    try
      set checkText to content of msg
    end try
  end if

  set ok to false
  if checkText contains "{_esc(probe)}" then set ok to true
  if checkText contains "Dear" then set ok to true
  if checkText contains "dear" then set ok to true
  if ok is false then
    error "Body chưa dán vào New Mail — hủy Send. Đóng Untitled, bật Accessibility, thử 1 mail."
  end if

  set toCount to 0
  try
    set toCount to count of (to recipients of msg)
  end try
  if toCount < 1 then error "To trống — hủy gửi"

  send msg
end tell
'''
        try:
            _run_osascript(script, timeout=180)
        except RuntimeError as exc:
            raise RuntimeError(
                "Outlook Mac gửi thất bại.\n"
                f"{exc}\n"
                "Gợi ý:\n"
                "- Legacy Outlook = ON\n"
                "- Accessibility: cho phép VT Rate Mail Sender\n"
                "- Đóng Untitled / xóa Outbox lỗi · thử 1 mail"
            ) from exc
