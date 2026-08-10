"""macOS — giống Windows: chữ ký Outlook chuẩn (cache 1 lần) + set content mỗi lần gửi.

Windows: COM tạo mail → HTMLBody đã có chữ ký/ảnh → merge body → Send (không UI paste).
Mac không có COM tương đương; set content + cid: làm gãy ảnh.

Cách Mac ≈ Windows:
1) Một lần: mở New Mail → Copy chữ ký gốc (NSAppleScript + Accessibility app)
   → lưu HTML + ảnh base64 (signature.html).
2) Mỗi lần gửi: merge body + chữ ký đã lưu → set content → Send
   (không System Events, không Accessibility).

Lưu ý: Legacy Outlook phải ON. New Outlook (toggle OFF) không tự động tốt.
"""

from __future__ import annotations

import re
import subprocess
import tempfile
import time
from ctypes import byref, c_bool, c_void_p, cdll
from pathlib import Path
from typing import Any, Optional

from ..models import AppConfig
from ..outlook_html import (
    convert_px_font_sizes_to_pt,
    harden_tables_for_outlook,
    strip_trailing_empty_blocks,
)
from ..paths import user_data_dir
from .html_merge import html_visible_text, merge_body_with_outlook_signature


def signature_dir() -> Path:
    d = user_data_dir() / "mac_signature"
    d.mkdir(parents=True, exist_ok=True)
    return d


def signature_path() -> Path:
    return signature_dir() / "signature.html"


def has_ready_signature() -> bool:
    p = signature_path()
    if not p.is_file():
        return False
    html = p.read_text(encoding="utf-8", errors="ignore")
    if len(html_visible_text(html)) < 20:
        return False
    # Bắt buộc có ảnh nhúng — giống chữ ký Outlook có logo
    return bool(re.search(r"data:image/", html, flags=re.I))


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


def _run_osascript_file(script: str, timeout: int = 180) -> str:
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


def _run_nsapplescript(source: str) -> str:
    """AppleScript trong process app → Accessibility của «VT Rate Mail Sender»."""
    try:
        from Foundation import NSAppleScript  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Thiếu PyObjC Foundation. Cài: pip install pyobjc-framework-Cocoa"
        ) from exc

    script = NSAppleScript.alloc().initWithSource_(source)
    if script is None:
        raise RuntimeError("Không tạo được NSAppleScript.")
    result, error = script.executeAndReturnError_(None)
    if error:
        try:
            desc = str(error.objectForKey_("NSAppleScriptErrorMessage") or error)
        except Exception:  # noqa: BLE001
            desc = str(error)
        raise RuntimeError(desc or "NSAppleScript failed")
    if result is None:
        return ""
    try:
        return str(result.stringValue() or "")
    except Exception:  # noqa: BLE001
        return str(result)


def ensure_assistive_access(*, prompt: bool = True) -> bool:
    """Kiểm tra Accessibility của process app; có thể mở Settings + prompt."""
    try:
        lib = cdll.LoadLibrary(
            "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices"
        )
        lib.AXIsProcessTrusted.restype = c_bool
        if bool(lib.AXIsProcessTrusted()):
            return True
        if not prompt:
            return False
        # Hiện dialog hệ thống nếu có thể
        try:
            import objc  # type: ignore
            from Foundation import NSDictionary  # type: ignore

            opts = NSDictionary.dictionaryWithObject_forKey_(
                True, "AXTrustedCheckOptionPrompt"
            )
            lib.AXIsProcessTrustedWithOptions.argtypes = [c_void_p]
            lib.AXIsProcessTrustedWithOptions.restype = c_bool
            ptr = c_void_p(objc.pyobjc_id(opts))
            if bool(lib.AXIsProcessTrustedWithOptions(ptr)):
                return True
        except Exception:  # noqa: BLE001
            pass
        subprocess.run(
            [
                "open",
                "x-apple.systempreferences:com.apple.preference.security?Privacy_Accessibility",
            ],
            check=False,
        )
        time.sleep(0.3)
        return bool(lib.AXIsProcessTrusted())
    except Exception:  # noqa: BLE001
        return False


def prepare_mac_html(html: str) -> str:
    s = (html or "").strip()
    if not s:
        return s
    s = convert_px_font_sizes_to_pt(s)
    s = harden_tables_for_outlook(s)
    s = strip_trailing_empty_blocks(s)
    return s


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
                "Bật Legacy Outlook = ON · đã đăng nhập.\n"
                f"Chi tiết: {err or 'osascript failed'}"
            )
        return self.config.from_email or "Outlook (Mac)"

    def signature_status(self) -> dict[str, Any]:
        ok = has_ready_signature()
        return {
            "ready": ok,
            "path": str(signature_path()),
            "has_file": ok,
            "has_images": ok,
            "message": (
                "Đã có chữ ký chuẩn Mac (giống New Mail, ảnh nhúng). Gửi = set content như Windows."
                if ok
                else "Chưa lấy chữ ký chuẩn. Bấm «Lấy chữ ký chuẩn (1 lần)» — Legacy ON + Accessibility."
            ),
        }

    def load_signature_html(self) -> str:
        if not has_ready_signature():
            return ""
        return signature_path().read_text(encoding="utf-8", errors="ignore")

    def capture_signature(self) -> str:
        """Alias nút UI — lấy chữ ký chuẩn 1 lần từ New Mail."""
        return self.capture_signature_from_new_mail()

    def capture_signature_from_new_mail(self) -> str:
        """
        Một lần (cần Accessibility + Legacy ON):
        New Mail + chữ ký AMBER → app tự focus body → Cmd+A → Cmd+C
        → đọc WebArchive/HTML → lưu base64.
        Nếu tự Copy thất bại: để New Mail mở, hướng dẫn user Cmd+A/C rồi bấm lại nút.
        """
        from ..clipboard_html import clipboard_html_for_compose, read_mac_clipboard_html

        self.open_outlook()

        # Nếu New Mail [VT-SIG] đang mở và clipboard đã có chữ ký (user vừa Cmd+C tay)
        existing = ""
        try:
            existing = clipboard_html_for_compose("")
        except Exception:  # noqa: BLE001
            existing = ""
        if (
            len(html_visible_text(existing)) >= 40
            and re.search(r"data:image/", existing, flags=re.I)
            and ("best regards" in existing.lower() or "vt logistics" in existing.lower() or "amber" in existing.lower())
        ):
            signature_path().write_text(existing, encoding="utf-8")
            n_img = len(re.findall(r"data:image/", existing, flags=re.I))
            self._close_vt_sig_windows()
            return (
                f"Đã lưu chữ ký từ Clipboard ({len(existing)} ký tự, {n_img} ảnh).\n"
                "Từ giờ Semi-Auto không cần Accessibility."
            )

        trusted = ensure_assistive_access(prompt=True)
        if not trusted:
            raise RuntimeError(
                "«VT Rate Mail Sender» chưa được Accessibility.\n"
                "System Settings → Privacy & Security → Accessibility → "
                "GỠ rồi BẬT lại «VT Rate Mail Sender» → mở lại app → thử lần nữa.\n"
                "(Mỗi lần build/codesign mới, macOS hay mất quyền dù toggle vẫn xanh.)"
            )

        # Mở New Mail với chữ ký gốc — chờ đủ để logo render
        open_script = f'''
tell application "Microsoft Outlook"
  activate
  set msg to make new outgoing message
{_account_block(self.config.from_email)}
  set subject of msg to "[VT-SIG] capture — will close"
  try
    set signature of msg to signature "AMBER"
  end try
  try
    set accObj to account of msg
    set sigs to signatures of accObj
    if (count of sigs) > 0 then set signature of msg to item 1 of sigs
  end try
  try
    if (count of signatures) > 0 then set signature of msg to item 1 of signatures
  end try
  open msg
  delay 3.0
end tell
'''
        _run_osascript_file(open_script, timeout=120)

        # Thử Copy tự động nhiều lần (nhiều điểm click thân thư)
        html = ""
        last_copy_err: Optional[Exception] = None
        for ratio in (0.72, 0.65, 0.80, 0.58):
            try:
                self._auto_select_copy_compose_body(y_ratio=ratio)
                time.sleep(0.5)
                html = clipboard_html_for_compose("")
                if len(html_visible_text(html)) >= 40 and re.search(
                    r"data:image/", html, flags=re.I
                ):
                    break
                # Có HTML nhưng chưa ảnh — thử đọc raw WebArchive lần nữa
                raw_html, _cmap = read_mac_clipboard_html()
                if raw_html.strip():
                    html = clipboard_html_for_compose(raw_html)
                if len(html_visible_text(html)) >= 40 and re.search(
                    r"data:image/", html, flags=re.I
                ):
                    break
                html = ""
            except Exception as exc:  # noqa: BLE001
                last_copy_err = exc
                html = ""
                continue

        if len(html_visible_text(html)) < 40 or not re.search(
            r"data:image/", html or "", flags=re.I
        ):
            # Giữ cửa sổ New Mail — user có thể Cmd+A Cmd+C rồi bấm lại nút
            detail = f"\nChi tiết: {last_copy_err}" if last_copy_err else ""
            raise RuntimeError(
                "App chưa Copy được HTML chữ ký tự động.\n"
                "Cửa sổ [VT-SIG] vẫn mở — làm tay rồi bấm lại «Lấy chữ ký chuẩn»:\n"
                "1) Click vào thân thư (chữ ký)\n"
                "2) Cmd+A → Cmd+C\n"
                "3) Quay lại app → bấm lại nút (app sẽ đọc Clipboard)\n"
                "Nhớ: Legacy Outlook = ON."
                + detail
            )

        signature_path().write_text(html, encoding="utf-8")
        self._close_vt_sig_windows()
        n_img = len(re.findall(r"data:image/", html, flags=re.I))
        return (
            f"Đã lưu chữ ký chuẩn từ New Mail ({len(html)} ký tự, {n_img} ảnh).\n"
            "Từ giờ Semi-Auto gửi như Windows (set content) — không cần Accessibility."
        )

    def _auto_select_copy_compose_body(self, *, y_ratio: float = 0.72) -> None:
        """Focus body + Cmd+A + Cmd+C (NSAppleScript / Accessibility app)."""
        script = f'''
tell application "Microsoft Outlook" to activate
delay 0.35
tell application "System Events"
  tell process "Microsoft Outlook"
    set frontmost to true
    delay 0.2
    set win to front window
    set {{wx, wy}} to position of win
    set {{ww, wh}} to size of win
    set clickX to (wx + (ww / 2)) as integer
    set clickY to (wy + (wh * {y_ratio})) as integer
    click at {{clickX, clickY}}
    delay 0.25
    -- Escape khỏi To/Cc nếu đang focus sai
    key code 53
    delay 0.1
    click at {{clickX, clickY}}
    delay 0.2
    keystroke "a" using {{command down}}
    delay 0.3
    keystroke "c" using {{command down}}
    delay 0.9
  end tell
end tell
return "ok"
'''
        _run_nsapplescript(script)

    def _close_vt_sig_windows(self) -> None:
        try:
            _run_osascript_file(
                '''
tell application "Microsoft Outlook"
  try
    repeat with m in (get outgoing messages)
      try
        if (subject of m as string) contains "[VT-SIG]" then
          delete m
        end if
      end try
    end repeat
  end try
end tell
''',
                timeout=60,
            )
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
        """
        Giống Windows merge: body app + chữ ký cache → set content → Send.
        Không UI paste.
        """
        if not has_ready_signature():
            raise RuntimeError(
                "Chưa có chữ ký chuẩn Mac.\n"
                "Bấm «Lấy chữ ký chuẩn (1 lần)» (Legacy ON + Accessibility lần đầu),\n"
                "rồi Semi-Auto — các lần sau không cần Accessibility."
            )

        prepared = prepare_mac_html(body_html)
        if not (prepared or "").strip():
            raise ValueError("Nội dung mail trống — không gửi.")
        if not to_list:
            raise ValueError("Thiếu địa chỉ To — không gửi.")

        sig = self.load_signature_html()
        full = merge_body_with_outlook_signature(prepared, sig)

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
            full_path = Path(tmp) / "full.html"
            full_path.write_text(full, encoding="utf-8")
            script = f'''
set fullPath to "{full_path}"
set fullText to do shell script "cat " & quoted form of fullPath
if fullText is "" then error "HTML trống"

tell application "Microsoft Outlook"
  activate
  set msg to make new outgoing message
{_account_block(self.config.from_email)}
  set subject of msg to "{_esc(subject)}"
{to_block}
{cc_block}
  set content of msg to fullText
{att_block}
  set toCount to 0
  try
    set toCount to count of (to recipients of msg)
  end try
  if toCount < 1 then error "To trống — hủy gửi"
  send msg
end tell
'''
            try:
                _run_osascript_file(script, timeout=180)
            except RuntimeError as exc:
                raise RuntimeError(
                    "Outlook Mac gửi thất bại.\n"
                    f"{exc}\n"
                    "Legacy Outlook = ON · đã «Lấy chữ ký chuẩn (1 lần)» · thử 1 mail."
                ) from exc
