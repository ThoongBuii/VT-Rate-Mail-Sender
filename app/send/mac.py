"""macOS Outlook Legacy — thao tác trên cửa sổ New Mail đang mở.

Bắt buộc: Legacy ON · Accessibility · chữ ký mặc định account.

1) New Mail + gắn To / Cc / Subject (AppleScript)
2) Focus body (Subject → Tab) rồi paste Soạn — không Cmd+V khi focus còn ở To
3) Gắn lại To/Cc (xóa rác nếu lỡ dán vào To) + Send

Không set content HTML (phá CID). Không WordEditor.
"""

from __future__ import annotations

import re
import subprocess
import tempfile
import time
import uuid
from ctypes import c_bool, c_void_p, cdll
from pathlib import Path
from typing import Any, Optional

from ..models import AppConfig
from ..outlook_html import prepare_body_html_for_outlook
from ..paths import user_data_dir
from .html_merge import html_visible_text


def signature_dir() -> Path:
    d = user_data_dir() / "mac_signature"
    d.mkdir(parents=True, exist_ok=True)
    return d


def signature_path() -> Path:
    return signature_dir() / "signature.html"


def has_ready_signature() -> bool:
    """Cache chữ ký (tuỳ chọn Preview) — Hướng 2 không bắt buộc khi gửi."""
    p = signature_path()
    if not p.is_file():
        return False
    html = p.read_text(encoding="utf-8", errors="ignore")
    if len(html_visible_text(html)) < 20:
        return False
    return bool(re.search(r"data:image/", html, flags=re.I))


def _strip_body_images(html: str) -> str:
    """Body gửi Hướng 2: bỏ mọi img — logo để chữ ký Outlook gắn."""
    from ..clipboard_html import strip_cf_html_headers

    s = strip_cf_html_headers(html or "")
    s = re.sub(r"(?is)<img\b[^>]*>", "", s)
    s = re.sub(r"(?is)<p[^>]*>\s*(?:&nbsp;|\u00a0|\s|<br\s*/?>)*</p>", "", s)
    return s.strip()


_SIG_NAME = "AMBER"


def _esc(s: str) -> str:
    return (s or "").replace("\\", "\\\\").replace('"', '\\"')


def _normalize_recipient_lists(
    to_list: list[str], cc_list: list[str]
) -> tuple[list[str], list[str]]:
    """Làm sạch + tách To/Cc trước khi đưa vào AppleScript."""
    from ..sender_smtp import extract_email, is_valid_email

    def _clean(items: list[str]) -> list[str]:
        out: list[str] = []
        seen: set[str] = set()
        for raw in items or []:
            em = extract_email(raw) or (raw or "").strip()
            if not is_valid_email(em):
                continue
            key = em.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(em)
        return out

    tos = _clean(to_list)
    to_keys = {e.lower() for e in tos}
    ccs = [e for e in _clean(cc_list) if e.lower() not in to_keys]
    return tos, ccs


def _applescript_set_recipients(to_list: list[str], cc_list: list[str]) -> str:
    """
    Xóa To/Cc cũ → gắn lại từng địa chỉ vào đúng ô.
    Không đụng content (tránh dán nhầm body vào To).
    """
    lines: list[str] = [
        "  -- To / Cc: xóa hết rồi gắn lại (đúng ô, đúng địa chỉ)",
        "  try",
        "    delete (every to recipient of msg)",
        "  end try",
        "  try",
        "    delete (every cc recipient of msg)",
        "  end try",
        "  try",
        "    delete (every bcc recipient of msg)",
        "  end try",
    ]
    for t in to_list:
        lines.append(
            "  make new to recipient at msg with properties "
            f'{{email address:{{address:"{_esc(t)}"}}}}'
        )
    for cc in cc_list:
        lines.append(
            "  make new cc recipient at msg with properties "
            f'{{email address:{{address:"{_esc(cc)}"}}}}'
        )
    return "\n".join(lines)


def _applescript_verify_recipients(to_list: list[str], cc_list: list[str]) -> str:
    """
    Kiểm tra To/Cc trước Send.

    Legacy Outlook thường lỗi khi ép `address of email address … as string`
    (-2700). Vì vậy: bắt buộc đúng số lượng; đọc địa chỉ kiểu mềm
    (đọc được thì so khớp, không đọc được thì chỉ tin count + lần set của ta).
    """
    want_to = len(to_list)
    want_cc = len(cc_list)
    checks: list[str] = [
        "  set toCount to 0",
        "  set ccCount to 0",
        "  try",
        "    set toCount to count of to recipients of msg",
        "  end try",
        "  try",
        "    set ccCount to count of cc recipients of msg",
        "  end try",
        f'  if toCount is not {want_to} then error '
        f'"To sai số lượng: " & toCount & " (cần {want_to})"',
        f'  if ccCount is not {want_cc} then error '
        f'"Cc sai số lượng: " & ccCount & " (cần {want_cc})"',
    ]
    for i, t in enumerate(to_list, start=1):
        te = _esc(t)
        checks.append(
            f"  -- soft-check To[{i}]\n"
            "  set tAddr to \"\"\n"
            "  try\n"
            f"    set ea to email address of to recipient {i} of msg\n"
            "    try\n"
            "      set tAddr to (address of ea) as text\n"
            "    end try\n"
            "  end try\n"
            "  if tAddr is not \"\" then\n"
            f'    if tAddr does not contain "{te}" then '
            f'error "To[{i}] sai: " & tAddr & " (cần {te})"\n'
            f"    if (length of tAddr) > 120 then "
            f'error "To[{i}] bị dán nhầm nội dung dài"\n'
            '    if tAddr contains "MicrosoftInternetExplorer" then '
            'error "To bẩn — hủy"\n'
            "  end if"
        )
    for i, cc in enumerate(cc_list, start=1):
        ce = _esc(cc)
        checks.append(
            f"  -- soft-check Cc[{i}]\n"
            "  set cAddr to \"\"\n"
            "  try\n"
            f"    set ea to email address of cc recipient {i} of msg\n"
            "    try\n"
            "      set cAddr to (address of ea) as text\n"
            "    end try\n"
            "  end try\n"
            "  if cAddr is not \"\" then\n"
            f'    if cAddr does not contain "{ce}" then '
            f'error "Cc[{i}] sai: " & cAddr & " (cần {ce})"\n'
            f"    if (length of cAddr) > 120 then "
            f'error "Cc[{i}] bị dán nhầm nội dung dài"\n'
            "  end if"
        )
    return "\n".join(checks)


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
    """Nhúng ảnh → data URI; từ chối nếu còn ảnh ngoài (banner privacy)."""
    from ..clipboard_html import (
        embed_local_and_cid_images,
        ensure_images_internal,
        strip_cf_html_headers,
    )

    s = strip_cf_html_headers(html or "")
    s = embed_local_and_cid_images(s)
    s = ensure_images_internal(s)
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

    def _close_stuck_vt_drafts(self) -> None:
        """Xóa draft [VT-…] còn sót từ lần paste New Mail cũ (tránh treo)."""
        try:
            _run_osascript_file(
                '''
tell application "Microsoft Outlook"
  try
    repeat with m in (get outgoing messages)
      try
        set s to subject of m as string
        if s starts with "[VT-" then
          try
            delete m
          end try
        end if
      end try
    end repeat
  end try
end tell
''',
                timeout=45,
            )
        except Exception:  # noqa: BLE001
            pass

    def open_outlook(self) -> str:
        """Legacy ON · Accessibility · AppleScript chạy được."""
        ensure_assistive_access(prompt=True)
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
                "Bật **Legacy Outlook = ON** (New Outlook tắt) · đã đăng nhập.\n"
                f"Chi tiết: {err or 'osascript failed'}"
            )
        if not ensure_assistive_access(prompt=False):
            raise RuntimeError(
                "Cần Accessibility cho «VT Rate Mail Sender».\n"
                "System Settings → Privacy & Security → Accessibility → bật app.\n"
                "(Dán body phía trên chữ ký AMBER — hết banner privacy.)"
            )
        accounts = self._list_outlook_accounts()
        if accounts:
            return accounts[0]
        return self.config.from_email or "tài khoản mặc định trên máy"

    def accessibility_status(self) -> dict[str, Any]:
        trusted = ensure_assistive_access(prompt=False)
        return {
            "trusted": trusted,
            "needed": True,
            "message": (
                "Accessibility OK — New Mail đang mở → paste → Send."
                if trusted
                else "Cần Accessibility để paste/Send trên cửa sổ New Mail."
            ),
        }

    def ensure_accessibility(self, *, prompt: bool = True) -> dict[str, Any]:
        ensure_assistive_access(prompt=prompt)
        return self.accessibility_status()

    def signature_status(self) -> dict[str, Any]:
        """Semi-Auto: cần Accessibility. Cache chữ ký chỉ tuỳ chọn Preview."""
        ax = ensure_assistive_access(prompt=False)
        cached = has_ready_signature()
        return {
            "ready": ax,
            "path": str(signature_path()) if cached else "",
            "has_file": cached,
            "has_images": cached,
            "cached": cached,
            "smtp_configured": False,
            "accessibility_trusted": ax,
            "message": (
                "Sẵn sàng: Legacy · New Mail mở · paste Soạn · Send."
                if ax
                else "Bật Accessibility → Legacy ON → New Mail tay thấy chữ ký mặc định."
            ),
        }

    def load_signature_html(self) -> str:
        if not has_ready_signature():
            return ""
        from ..clipboard_html import (
            embed_local_and_cid_images,
            embed_remote_http_images,
            strip_cf_html_headers,
        )

        raw = signature_path().read_text(encoding="utf-8", errors="ignore")
        cleaned = embed_remote_http_images(
            embed_local_and_cid_images(strip_cf_html_headers(raw))
        )
        if cleaned != raw and len(html_visible_text(cleaned)) >= 20:
            try:
                # Chỉ ghi lại nếu đã internal hoàn toàn
                _prepare_signature_html(cleaned)
                signature_path().write_text(cleaned, encoding="utf-8")
            except Exception:  # noqa: BLE001
                pass
            return cleaned
        return raw

    def save_signature_html(self, html: str = "", *, prefer_clipboard: bool = True) -> str:
        """
        Lưu chữ ký do user Copy từ Outlook.
        Ưu tiên clipboard hệ thống (WebArchive giữ ảnh). Không lưu nếu còn ảnh ngoài.
        """
        from ..clipboard_html import clipboard_html_for_compose

        browser = (html or "").strip()
        normalized = ""
        try:
            if prefer_clipboard:
                normalized = clipboard_html_for_compose("")
            if len(html_visible_text(normalized)) < 30 and browser:
                normalized = clipboard_html_for_compose(browser)
            elif not normalized.strip() and browser:
                normalized = browser
        except Exception as exc:  # noqa: BLE001
            if browser:
                normalized = browser
            else:
                raise RuntimeError(
                    "Không đọc được chữ ký từ Clipboard.\n"
                    "Outlook: New Mail → chọn chữ ký + logo → Cmd+C → «Dán từ Clipboard»."
                ) from exc

        prepared = _prepare_signature_html(normalized)
        if len(html_visible_text(prepared)) < 30:
            raise RuntimeError(
                "Chữ ký quá ngắn / clipboard trống.\n"
                "Legacy Outlook → New Mail → chọn chữ ký + logo → Cmd+C → «Dán từ Clipboard»."
            )
        # New Outlook hay copy ra signature_NNNN (ảnh vỡ) — từ chối lưu
        if re.search(r"signature_\d+", prepared, flags=re.I) and not re.search(
            r"data:image/", prepared, flags=re.I
        ):
            raise RuntimeError(
                "Clipboard chữ ký hỏng (signature_* / ảnh ?).\n"
                "New Outlook không copy được logo.\n"
                "Bật Legacy Outlook = ON → New Mail → chọn cả logo → Cmd+C lại."
            )
        n_img = len(re.findall(r"data:image/", prepared, flags=re.I))
        if n_img < 1:
            raise RuntimeError(
                "Chữ ký không có ảnh nhúng (logo).\n"
                "Legacy ON → chọn cả vùng chữ ký có logo → Cmd+C.\n"
                "Không dùng Cmd+V vào khung nếu thấy icon ảnh vỡ."
            )

        signature_path().write_text(prepared, encoding="utf-8")
        return (
            f"Đã lưu chữ ký preview ({len(prepared)} ký tự, {n_img} ảnh). "
            "Semi-Auto dùng chữ ký AMBER sống trong Outlook (Hướng 2)."
        )

    def preview_signature_from_clipboard(self) -> str:
        """Đọc clipboard → HTML đã nhúng ảnh (không lưu)."""
        from ..clipboard_html import clipboard_html_for_compose

        normalized = clipboard_html_for_compose("")
        return _prepare_signature_html(normalized)

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

    def _focus_body_then_paste(self) -> None:
        """
        Đưa focus vào ô nội dung rồi Cmd+V.

        Sau AppleScript gắn To, focus thường còn ở ô To → Cmd+V sẽ dán HTML
        thành hàng trăm «người nhận» rác (Style Definitions / Dear…).
        Cách an toàn: click Subject → Tab vào body → mới paste.
        """
        script = '''
tell application "Microsoft Outlook" to activate
delay 0.1
tell application "System Events"
  tell process "Microsoft Outlook"
    set frontmost to true
    delay 0.08
    set win to front window
    set {wx, wy} to position of win
    set {ww, wh} to size of win
    set clickX to (wx + (ww / 2)) as integer
    -- ~24% từ đỉnh: vùng Subject (dưới To/Cc, trên body)
    set clickY to (wy + (wh * 0.24)) as integer
    key code 53
    delay 0.04
    click at {clickX, clickY}
    delay 0.12
    -- Tab: Subject → Message body (Legacy Outlook Mac)
    keystroke tab
    delay 0.18
    -- Đầu vùng body (trên chữ ký)
    key code 126 using {command down}
    delay 0.06
    keystroke "v" using {command down}
    delay 0.3
  end tell
end tell
return "ok"
'''
        _run_nsapplescript(script)

    def _reapply_recipients_and_send(
        self,
        *,
        subject: str,
        to_list: list[str],
        cc_list: list[str],
    ) -> None:
        """
        Gắn lại To/Cc đúng (xóa rác nếu paste nhầm vào To) rồi Send.
        Tìm đúng New Mail vừa mở theo Subject — không dùng token [VT-].
        """
        recip_block = _applescript_set_recipients(to_list, cc_list)
        verify_block = _applescript_verify_recipients(to_list, cc_list)
        want = _esc(subject)
        script = f'''
tell application "Microsoft Outlook"
  set msg to missing value
  set wantSubj to "{want}"
  try
    repeat with m in (get outgoing messages)
      try
        if (subject of m as string) is wantSubj then
          set msg to m
          exit repeat
        end if
      end try
    end repeat
  end try
  if msg is missing value then
    try
      repeat with m in (messages of drafts folder)
        try
          if (subject of m as string) is wantSubj then
            set msg to m
            exit repeat
          end if
        end try
      end repeat
    end try
  end if
  if msg is missing value then error "Không thấy New Mail với subject vừa đặt"
{recip_block}
{verify_block}
  set subject of msg to wantSubj
  send msg
end tell
'''
        _run_osascript_file(script, timeout=60)

    def send(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Path | list[Path]] = None,
    ) -> None:
        """
        New Mail → To/Cc/Subject → focus body → paste → gắn lại To/Cc → Send.
        Tránh Cmd+V khi focus còn ở ô To (lỗi To đầy HTML rác).
        """
        from ..clipboard_html import set_mac_clipboard_html, strip_cf_html_headers

        if not ensure_assistive_access(prompt=True):
            raise RuntimeError(
                "Cần Accessibility cho «VT Rate Mail Sender».\n"
                "System Settings → Privacy & Security → Accessibility → bật app."
            )

        prepared = prepare_body_html_for_outlook(body_html)
        prepared = strip_cf_html_headers(prepared)
        prepared = _strip_body_images(prepared)
        if not (prepared or "").strip():
            raise ValueError("Nội dung mail trống — không gửi.")
        if len(html_visible_text(prepared)) < 20:
            raise ValueError(
                "Body quá ngắn sau khi bỏ ảnh.\n"
                "Soạn chỉ bảng/text — logo để chữ ký mặc định Outlook."
            )
        if re.search(r"(?is)<img\b", prepared) or re.search(
            r"(?i)(?:src|background)\s*=\s*[\"']?(?:data:|https?:|file:)",
            prepared,
        ):
            raise ValueError(
                "Body còn ảnh / URL ảnh.\n"
                "Xóa logo khỏi Soạn; New Mail sẽ tự có chữ ký AMBER."
            )

        to_list, cc_list = _normalize_recipient_lists(to_list, cc_list or [])
        if not to_list:
            raise ValueError("Thiếu / sai địa chỉ To — không gửi.")

        body_for_paste = (
            f"{prepared}"
            f'<div style="margin:0;padding:0"><br/></div>'
        )

        files = self._normalize_files(attachment)
        prev_app = _frontmost_app_name()
        recip_block = _applescript_set_recipients(to_list, cc_list)
        verify_block = _applescript_verify_recipients(to_list, cc_list)
        att_block = "\n".join(
            f'  make new attachment at msg with properties '
            f'{{file:POSIX file "{_esc(str(p.resolve()))}"}}'
            for p in files
        )

        self._close_stuck_vt_drafts()
        set_mac_clipboard_html(body_for_paste)

        # 1) New Mail + To/Cc/Subject đúng
        create_script = f'''
tell application "Microsoft Outlook"
  activate
  set msg to make new outgoing message
{_account_block(self.config.from_email)}
  set subject of msg to "{_esc(subject)}"
{recip_block}
{att_block}
{verify_block}
  open msg
end tell
return "ok"
'''
        try:
            try:
                _run_nsapplescript(create_script)
            except Exception:  # noqa: BLE001
                _run_osascript_file(create_script, timeout=45)
        except RuntimeError as exc:
            raise RuntimeError(
                "Không mở được New Mail.\n"
                f"{exc}\n"
                "Legacy Outlook = ON · đã login."
            ) from exc

        time.sleep(0.4)

        # 2) Paste vào BODY (Subject→Tab), không paste khi focus còn ở To
        try:
            set_mac_clipboard_html(body_for_paste)
            self._focus_body_then_paste()
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(
                "Không paste được vào ô nội dung New Mail.\n"
                "Accessibility · cửa sổ New Mail không bị che.\n"
                f"{exc}"
            ) from exc

        time.sleep(0.25)

        # 3) Gắn lại To/Cc (xóa rác nếu lỡ dán vào To) + Send
        try:
            self._reapply_recipients_and_send(
                subject=subject, to_list=to_list, cc_list=cc_list
            )
        except RuntimeError as exc:
            raise RuntimeError(
                "Gửi thất bại sau khi gắn lại To/Cc.\n"
                f"{exc}\n"
                f"To cần: {', '.join(to_list)}\n"
                f"Subject: {subject}"
            ) from exc

        time.sleep(0.2)
        _activate_app(prev_app)







