"""macOS Outlook — user dán chữ ký 1 lần + set content (nhanh như bản cũ).

1) User: Outlook New Mail → Copy chữ ký → app nút «Chữ ký» → Dán/Lưu
   (ảnh http → data URI internal — hết banner privacy)
2) Mỗi gửi: merge body + chữ ký cache → set content → Send
   (nhiều đính kèm, ảnh trong message; không mở cửa sổ)

Không tự mở New Mail lấy chữ ký. Legacy Outlook = ON.
"""

from __future__ import annotations

import re
import subprocess
import tempfile
import time
from ctypes import c_bool, c_void_p, cdll
from pathlib import Path
from typing import Any, Optional

from ..models import AppConfig
from ..outlook_html import prepare_body_html_for_outlook
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
    try:
        lib = cdll.LoadLibrary(
            "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices"
        )
        lib.AXIsProcessTrusted.restype = c_bool
        if bool(lib.AXIsProcessTrusted()):
            return True
        if not prompt:
            return False
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


def _frontmost_app_name() -> str:
    try:
        return _run_osascript_file(
            'tell application "System Events" to get name of first application process whose frontmost is true',
            timeout=10,
        )
    except Exception:  # noqa: BLE001
        return ""


def _activate_app(name: str) -> None:
    if not name or name.lower() in {"microsoft outlook", "outlook"}:
        return
    try:
        _run_osascript_file(
            f'tell application "{_esc(name)}" to activate',
            timeout=10,
        )
    except Exception:  # noqa: BLE001
        pass


def _prepare_signature_html(html: str) -> str:
    """Nhúng ảnh remote → data URI (internal) — hết banner privacy."""
    from ..clipboard_html import (
        embed_local_and_cid_images,
        embed_remote_http_images,
        strip_cf_html_headers,
    )

    s = strip_cf_html_headers(html or "")
    s = embed_local_and_cid_images(s)
    s = embed_remote_http_images(s)
    return s.strip()


class MacOutlookSender:
    def __init__(self, config: AppConfig):
        self.config = config

    def _list_outlook_accounts(self) -> list[str]:
        script = """
tell application "Microsoft Outlook"
  set outList to {}
  try
    repeat with acc in exchange accounts
      try
        set end of outList to (email address of acc as string)
      end try
    end repeat
  end try
  try
    repeat with acc in imap accounts
      try
        set end of outList to (email address of acc as string)
      end try
    end repeat
  end try
  try
    repeat with acc in pop accounts
      try
        set end of outList to (email address of acc as string)
      end try
    end repeat
  end try
  set AppleScript's text item delimiters to "|"
  set joined to outList as string
  set AppleScript's text item delimiters to ""
  return joined
end tell
"""
        try:
            raw = _run_osascript_file(script, timeout=60)
        except Exception:  # noqa: BLE001
            return []
        emails: list[str] = []
        seen: set[str] = set()
        for part in (raw or "").split("|"):
            em = part.strip()
            if not em or "@" not in em:
                continue
            key = em.lower()
            if key in seen:
                continue
            seen.add(key)
            emails.append(em)
        return emails

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
        accounts = self._list_outlook_accounts()
        if accounts:
            return accounts[0]
        return "tài khoản mặc định trên máy"

    def accessibility_status(self) -> dict[str, Any]:
        trusted = ensure_assistive_access(prompt=False)
        return {
            "trusted": trusted,
            "needed": False,
            "message": "Không bắt buộc — chữ ký do bạn dán thủ công.",
        }

    def ensure_accessibility(self, *, prompt: bool = True) -> dict[str, Any]:
        ensure_assistive_access(prompt=prompt)
        return self.accessibility_status()

    def signature_status(self) -> dict[str, Any]:
        ok = has_ready_signature()
        return {
            "ready": ok,
            "path": str(signature_path()) if ok else "",
            "has_file": ok,
            "has_images": ok,
            "cached": ok,
            "accessibility_trusted": ensure_assistive_access(prompt=False),
            "message": (
                "Đã có chữ ký (ảnh internal). Có thể Semi-Auto."
                if ok
                else "Chưa có chữ ký — bấm «Chữ ký» → dán chữ ký Outlook vào rồi Lưu."
            ),
        }

    def load_signature_html(self) -> str:
        if not has_ready_signature():
            return ""
        raw = signature_path().read_text(encoding="utf-8", errors="ignore")
        cleaned = _prepare_signature_html(raw)
        if cleaned != raw and len(html_visible_text(cleaned)) >= 20:
            try:
                signature_path().write_text(cleaned, encoding="utf-8")
            except Exception:  # noqa: BLE001
                pass
            return cleaned
        return raw

    def save_signature_html(self, html: str = "") -> str:
        """
        Lưu chữ ký do user Copy/dán từ Outlook.
        Ưu tiên clipboard hệ thống (giữ ảnh); html từ khung dán là dự phòng.
        """
        from ..clipboard_html import clipboard_html_for_compose

        try:
            # Đọc clip Mac trước; nếu trống thì dùng html user dán vào khung
            normalized = clipboard_html_for_compose(html or "")
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(
                "Không đọc được chữ ký.\n"
                "Outlook: New Mail → Copy chữ ký (Cmd+C) → app: «Dán từ Clipboard»."
            ) from exc

        prepared = _prepare_signature_html(normalized)
        if len(html_visible_text(prepared)) < 30:
            raise RuntimeError(
                "Chữ ký quá ngắn / clipboard trống.\n"
                "Outlook → New Mail → chọn chữ ký + logo → Cmd+C → «Dán từ Clipboard»."
            )

        signature_path().write_text(prepared, encoding="utf-8")
        n_img = len(re.findall(r"data:image/", prepared, flags=re.I))
        note = (
            f"{n_img} ảnh internal (không banner privacy)."
            if n_img
            else "không có ảnh — nên Copy cả logo từ Outlook."
        )
        return f"Đã lưu chữ ký ({len(prepared)} ký tự, {note}) Có thể Semi-Auto."

    def capture_signature(self) -> str:
        """Tương thích API cũ — lưu từ clipboard hệ thống."""
        return self.save_signature_html("")

    def capture_signature_from_new_mail(self) -> str:
        return self.save_signature_html("")

    def _normalize_files(
        self, attachment: Optional[Path | list[Path]]
    ) -> list[Path]:
        if isinstance(attachment, Path):
            return [attachment]
        if attachment:
            return [p for p in attachment if p is not None]
        return []

    def send(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Path | list[Path]] = None,
    ) -> None:
        """
        Cache chữ ký + merge body → set content → Send.
        Không mở cửa sổ, không paste UI, không đọc lại content (tránh timeout).
        """
        from ..clipboard_html import strip_cf_html_headers

        if not has_ready_signature():
            raise RuntimeError(
                "Chưa có chữ ký Mac.\n"
                "Bấm «Chữ ký» → Copy chữ ký trong Outlook → Dán / Lưu,\n"
                "rồi Semi-Auto."
            )

        prepared = prepare_body_html_for_outlook(body_html)
        prepared = strip_cf_html_headers(prepared)
        # Ảnh trong message: giữ data URI / harden (prepare_body đã xử lý)
        if not (prepared or "").strip():
            raise ValueError("Nội dung mail trống — không gửi.")
        if not to_list:
            raise ValueError("Thiếu địa chỉ To — không gửi.")

        sig = self.load_signature_html()
        full = merge_body_with_outlook_signature(prepared, sig)
        full = _prepare_signature_html(full)  # chắc chắn không còn http ảnh
        if len(html_visible_text(full)) < 80:
            raise RuntimeError("Nội dung mail sau merge quá ngắn — hủy gửi.")

        files = self._normalize_files(attachment)
        prev_app = _frontmost_app_name()

        to_block = "\n".join(
            f'  make new to recipient at msg with properties {{email address:{{address:"{_esc(t)}"}}}}'
            for t in to_list
        )
        cc_block = "\n".join(
            f'  make new cc recipient at msg with properties {{email address:{{address:"{_esc(cc)}"}}}}'
            for cc in cc_list
        )
        att_block = "\n".join(
            f'  make new attachment at msg with properties {{file:POSIX file "{_esc(str(p.resolve()))}"}}'
            for p in files
        )

        with tempfile.TemporaryDirectory() as tmp:
            full_path = Path(tmp) / "full.html"
            full_path.write_text(full, encoding="utf-8")

            # Đọc utf8 từ file — tránh cat cắt HTML; không open / không đọc lại content
            script = f'''
set fullPath to "{full_path}"
set fullText to read (POSIX file fullPath) as «class utf8»
if fullText is "" then error "HTML trống"

tell application "Microsoft Outlook"
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
                _run_osascript_file(script, timeout=120)
            except RuntimeError as exc:
                raise RuntimeError(
                    "Outlook Mac gửi thất bại.\n"
                    f"{exc}\n"
                    "Legacy Outlook = ON · đã «Lấy chữ ký chuẩn (1 lần)»."
                ) from exc

        _activate_app(prev_app)
