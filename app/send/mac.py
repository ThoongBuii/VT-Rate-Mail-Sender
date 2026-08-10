"""macOS Legacy Outlook — chụp chữ ký 1 lần, gửi bằng set content (không Tab/Cmd+V)."""

from __future__ import annotations

import subprocess
import tempfile
from pathlib import Path
from typing import Optional

from ..models import AppConfig
from ..outlook_html import prepare_body_html_for_outlook
from ..paths import user_data_dir
from .html_merge import html_visible_text, merge_body_with_outlook_signature


def signature_path() -> Path:
    return user_data_dir() / "signature_mac.html"


def has_signature() -> bool:
    p = signature_path()
    if not p.is_file():
        return False
    return len(html_visible_text(p.read_text(encoding="utf-8"))) >= 8


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


def _clipboard_html() -> str:
    try:
        from AppKit import NSPasteboard, NSPasteboardTypeHTML, NSPasteboardTypeString  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Thiếu PyObjC (AppKit). Cài: pip install pyobjc-framework-Cocoa"
        ) from exc

    pb = NSPasteboard.generalPasteboard()
    html = pb.stringForType_(NSPasteboardTypeHTML) or ""
    if html.strip():
        return str(html)
    plain = pb.stringForType_(NSPasteboardTypeString) or ""
    if plain.strip():
        return f"<div>{plain}</div>"
    return ""


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
        ok = has_signature()
        return {
            "ready": ok,
            "path": str(signature_path()),
            "message": (
                "Đã có chữ ký Mac — sẵn sàng gửi."
                if ok
                else "Chưa chụp chữ ký. Bấm «Chụp chữ ký Outlook» (1 lần)."
            ),
        }

    def capture_signature(self) -> str:
        """
        Mở New Mail trống → chờ chữ ký UI → lấy HTML (content hoặc Copy)
        → lưu signature_mac.html → đóng cửa sổ (không gửi).
        """
        self.open_outlook()
        script = f'''
tell application "Microsoft Outlook"
  activate
  set msg to make new outgoing message
{_account_block(self.config.from_email)}
  set subject of msg to "[VT-SIG-CAPTURE] — đóng sau khi app lấy chữ ký"
  try
    set accObj to account of msg
    set sigs to signatures of accObj
    if (count of sigs) > 0 then set signature of msg to item 1 of sigs
  end try
  try
    if (count of signatures) > 0 then set signature of msg to item 1 of signatures
  end try
  open msg
  delay 2.5
  set mid to ""
  try
    set mid to (id of msg) as string
  end try
  set htmlSig to ""
  try
    set htmlSig to content of msg
  end try
  if (length of htmlSig) < 40 then
    try
      set htmlSig to plain text content of msg
    end try
  end if
  return mid & "|||" & htmlSig
end tell
'''
        out = _run_osascript(script)
        mid, _, html = out.partition("|||")
        html = html or ""

        if len(html_visible_text(html)) < 12:
            # Fallback: user-visible Copy body (1 lần khi chụp chữ ký)
            copy_script = '''
tell application "System Events"
  if not UI elements enabled then
    error "Bật Accessibility cho VT Rate Mail Sender"
  end if
  tell process "Microsoft Outlook"
    set frontmost to true
    delay 0.4
    keystroke tab
    delay 0.12
    keystroke tab
    delay 0.12
    keystroke tab
    delay 0.2
    keystroke "a" using {command down}
    delay 0.2
    keystroke "c" using {command down}
    delay 0.4
  end tell
end tell
'''
            try:
                _run_osascript(copy_script)
                html = _clipboard_html()
            except Exception:  # noqa: BLE001
                pass

        if len(html_visible_text(html)) < 12:
            # Đóng cửa sổ capture nếu có
            self._delete_capture_message(mid)
            raise RuntimeError(
                "Không lấy được chữ ký từ New Mail.\n"
                "Hãy: Legacy Outlook ON → New Mail tay vẫn thấy chữ ký → "
                "bấm lại «Chụp chữ ký Outlook» và không đụng chuột trong ~3 giây."
            )

        signature_path().write_text(html, encoding="utf-8")
        self._delete_capture_message(mid)
        return f"Đã lưu chữ ký Mac ({len(html)} ký tự). Có thể Semi-Auto."

    def _delete_capture_message(self, mid: str) -> None:
        if not mid:
            return
        script = f'''
tell application "Microsoft Outlook"
  try
    set msg to message id "{_esc(mid)}"
    delete msg
  end try
  try
    set msgs to (outgoing messages whose subject contains "[VT-SIG-CAPTURE]")
    repeat with m in msgs
      try
        delete m
      end try
    end repeat
  end try
end tell
'''
        try:
            _run_osascript(script, timeout=30)
        except Exception:  # noqa: BLE001
            pass

    def send(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Path],
    ) -> None:
        if not has_signature():
            raise RuntimeError(
                "Chưa chụp chữ ký Mac.\n"
                "Bấm «Chụp chữ ký Outlook» một lần rồi gửi lại."
            )
        prepared = prepare_body_html_for_outlook(body_html)
        if not (prepared or "").strip():
            raise ValueError("Nội dung mail trống — không gửi.")
        if not to_list:
            raise ValueError("Thiếu địa chỉ To — không gửi.")

        sig = signature_path().read_text(encoding="utf-8")
        full_html = merge_body_with_outlook_signature(prepared, sig)
        if len(html_visible_text(full_html)) < 8:
            raise ValueError("Body+chữ ký trống sau merge — chụp lại chữ ký.")

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

        with tempfile.TemporaryDirectory() as tmp:
            body_path = Path(tmp) / "full.html"
            body_path.write_text(full_html, encoding="utf-8")
            script = f'''
set bodyPath to "{body_path}"
set bodyText to do shell script "cat " & quoted form of bodyPath
if bodyText is "" then error "HTML trống"

tell application "Microsoft Outlook"
  activate
  set msg to make new outgoing message
{_account_block(self.config.from_email)}
  set subject of msg to "{_esc(subject)}"
{to_block}
{cc_block}
{att_block}
  set content of msg to bodyText
  delay 0.3
  set toCount to 0
  try
    set toCount to count of (to recipients of msg)
  end try
  if toCount < 1 then error "To trống — hủy gửi"
  send msg
end tell
'''
            _run_osascript(script, timeout=180)
