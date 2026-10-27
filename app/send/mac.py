"""macOS Outlook Legacy — quy trình chuẩn.

1) AS: New Mail + To/Cc/Subject
2) AX focused Subject → Tab 1 → Cmd+V → Cmd+Return
3) Trả focus → verify Sent
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


def _strip_external_body_images(html: str) -> str:
    """
    Giữ ảnh data: (paste OK, người nhận thấy được).
    Bỏ img http/https/file/cid — dễ banner privacy / vỡ ảnh.
    """
    from ..clipboard_html import strip_cf_html_headers

    s = strip_cf_html_headers(html or "")

    def _img_repl(m: re.Match[str]) -> str:
        tag = m.group(0)
        src_m = re.search(r"""(?is)\bsrc\s*=\s*['"]([^'"]*)['"]""", tag)
        if not src_m:
            return ""
        src = (src_m.group(1) or "").strip()
        if src.lower().startswith("data:image/"):
            return tag
        return ""

    s = re.sub(r"(?is)<img\b[^>]*>", _img_repl, s)
    s = re.sub(r"(?is)<p[^>]*>\s*(?:&nbsp;|\u00a0|\s|<br\s*/?>)*</p>", "", s)
    return s.strip()


def _prepare_mac_paste_body(body_html: str) -> str:
    """Đồng bộ Preview/Windows: pt/bảng + sát chữ ký (không <br> thừa)."""
    from ..clipboard_html import strip_cf_html_headers
    from ..outlook_html import prepare_body_html_for_outlook, strip_trailing_empty_blocks
    from .html_merge import zero_last_block_margin

    s = prepare_body_html_for_outlook(body_html)
    s = strip_cf_html_headers(s)
    s = _strip_external_body_images(s)
    s = strip_trailing_empty_blocks(s)
    s = zero_last_block_margin(s)
    # Wrapper app: không tạo khoảng dưới trước Remark chữ ký
    m = re.match(r"^(<div(?:\s[^>]*)?)>", s or "", flags=re.I)
    if m:
        from .html_merge import inject_style_attr

        s = inject_style_attr(m.group(1) + ">", "margin-bottom:0;padding-bottom:0") + s[m.end() :]
    return (s or "").strip()


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
        "    set oldTos to to recipients of msg",
        "    repeat with r in oldTos",
        "      try",
        "        delete r",
        "      end try",
        "    end repeat",
        "  end try",
        "  try",
        "    set oldCcs to cc recipients of msg",
        "    repeat with r in oldCcs",
        "      try",
        "        delete r",
        "      end try",
        "    end repeat",
        "  end try",
        "  try",
        "    set oldBccs to bcc recipients of msg",
        "    repeat with r in oldBccs",
        "      try",
        "        delete r",
        "      end try",
        "    end repeat",
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

    def _any_compose_window_open(self) -> bool:
        script = '''
tell application "System Events"
  tell process "Microsoft Outlook"
    try
      repeat with w in windows
        try
          set n to name of w as text
          if (n contains "@") or (n contains "Untitled") then return "yes"
        end try
      end repeat
    end try
  end tell
end tell
return "no"
'''
        try:
            return _run_osascript_file(script, timeout=12).strip().lower().startswith(
                "yes"
            )
        except Exception:  # noqa: BLE001
            return False

    def _close_compose_ui_windows(self) -> None:
        """
        Đóng New Mail bằng Cmd+W (không click chuột / không delete outgoing).

        CẤM `delete` outgoing: Outlook Mac `outgoing messages` còn gồm mail
        đã nằm Sent (cùng id) → delete = xóa Sent Items.
        """
        try:
            from .mac_hid import close_front_compose_keyboard

            for _ in range(3):
                if not self._any_compose_window_open():
                    break
                close_front_compose_keyboard()
                time.sleep(0.2)
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.15)

    def _discard_all_outgoing(self) -> None:
        """Không xóa mail — chỉ đóng UI compose."""
        self._close_compose_ui_windows()

    def _discard_outgoing_subject(self, subject: str) -> None:
        """Không xóa mail — chỉ đóng UI compose."""
        _ = subject
        self._close_compose_ui_windows()

    def _close_stuck_vt_drafts(self) -> None:
        self._close_compose_ui_windows()

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
                "(Paste Soạn trên chữ ký AMBER sống — không banner privacy.)"
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
                "Accessibility OK — G: New Mail+field → paste compose → gửi."
                if trusted
                else "Cần Accessibility để paste / Cmd+Return."
            ),
        }

    def ensure_accessibility(self, *, prompt: bool = True) -> dict[str, Any]:
        ensure_assistive_access(prompt=prompt)
        return self.accessibility_status()

    def signature_status(self) -> dict[str, Any]:
        """Semi-Auto: chữ ký sống trong New Mail; cache chỉ tuỳ chọn Preview."""
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
                "Sẵn sàng: Legacy · New Mail+To/Cc/Subject → paste compose · chữ ký AMBER."
                if ax
                else "Bật Accessibility + Legacy ON · New Mail tay thấy chữ ký AMBER."
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

    def _outgoing_ids(self) -> list[str]:
        """Danh sách id outgoing hiện có (tránh gắn field vào draft cũ)."""
        script = '''
tell application "Microsoft Outlook"
  set outIds to {}
  try
    repeat with m in (get outgoing messages)
      try
        set end of outIds to (id of m as string)
      end try
    end repeat
  end try
  set AppleScript's text item delimiters to "|"
  set joined to outIds as string
  set AppleScript's text item delimiters to ""
  return joined
end tell
'''
        try:
            raw = _run_osascript_file(script, timeout=20)
        except Exception:  # noqa: BLE001
            return []
        return [p.strip() for p in (raw or "").split("|") if p.strip()]

    def _newest_outgoing_id(self, *, exclude_ids: list[str]) -> str:
        """Id outgoing mới sau Cmd+N (không trùng draft cũ)."""
        exclude = {i for i in exclude_ids}
        for _ in range(12):
            for oid in self._outgoing_ids():
                if oid not in exclude:
                    return oid
            time.sleep(0.12)
        return ""

    def _paste_and_send_burst(self, *, has_cc: bool, subject: str = "") -> None:
        """AX Subject → Tab 1 → paste → Cmd+Return."""
        from .mac_hid import paste_and_send_keyboard

        paste_and_send_keyboard(has_cc=has_cc, subject_hint=subject)

    def _ui_send(self, *, subject: str = "") -> None:
        from .mac_hid import send_compose_cmd_return

        send_compose_cmd_return(already_in_body=True, subject_hint=subject)

    def _message_in_sent(self, subject: str, *, within_minutes: int = 5) -> bool:
        """
        True nếu subject vừa gửi — quét Sent Exchange/IMAP/POP + folder tên Sent*
        (POP «On My Computer» thường không nằm trong sent items mặc định).
        """
        want = _esc(subject)
        mins = max(1, int(within_minutes))
        # Một script phẳng: khớp subject; time sent optional (POP đôi khi thiếu).
        script = f'''
tell application "Microsoft Outlook"
  set wantSubj to "{want}"
  set cutoff to (current date) - ({mins} * minutes)

  -- 1) sent items mặc định
  try
    repeat with m in (get messages of sent items)
      try
        if (subject of m as string) is wantSubj then
          try
            if (time sent of m) ≥ cutoff then return "yes"
          on error
            return "yes"
          end try
        end if
      end try
    end repeat
  end try

  -- 2) Exchange
  try
    repeat with acc in (get exchange accounts)
      try
        repeat with m in (get messages of sent items of acc)
          try
            if (subject of m as string) is wantSubj then
              try
                if (time sent of m) ≥ cutoff then return "yes"
              on error
                return "yes"
              end try
            end if
          end try
        end repeat
      end try
    end repeat
  end try

  -- 3) IMAP
  try
    repeat with acc in (get imap accounts)
      try
        repeat with m in (get messages of sent items of acc)
          try
            if (subject of m as string) is wantSubj then
              try
                if (time sent of m) ≥ cutoff then return "yes"
              on error
                return "yes"
              end try
            end if
          end try
        end repeat
      end try
    end repeat
  end try

  -- 4) POP (On My Computer / local)
  try
    repeat with acc in (get pop accounts)
      try
        repeat with m in (get messages of sent items of acc)
          try
            if (subject of m as string) is wantSubj then
              try
                if (time sent of m) ≥ cutoff then return "yes"
              on error
                return "yes"
              end try
            end if
          end try
        end repeat
      end try
      try
        repeat with fld in (get mail folders of acc)
          try
            set nm to name of fld as string
            if nm is "Sent" or nm is "Sent Items" or nm is "Sent Messages" then
              repeat with m in (get messages of fld)
                try
                  if (subject of m as string) is wantSubj then
                    try
                      if (time sent of m) ≥ cutoff then return "yes"
                    on error
                      return "yes"
                    end try
                  end if
                end try
              end repeat
            end if
          end try
        end repeat
      end try
    end repeat
  end try

  -- 5) Mọi mail folder tên Sent* (fallback On My Computer)
  try
    repeat with fld in (get mail folders)
      try
        set nm to name of fld as string
        if nm is "Sent" or nm is "Sent Items" or nm is "Sent Messages" then
          repeat with m in (get messages of fld)
            try
              if (subject of m as string) is wantSubj then
                try
                  if (time sent of m) ≥ cutoff then return "yes"
                on error
                  return "yes"
                end try
              end if
            end try
          end repeat
        end if
      end try
    end repeat
  end try

  return "no"
end tell
'''
        try:
            return _run_osascript_file(script, timeout=40).strip().lower().startswith(
                "yes"
            )
        except Exception:  # noqa: BLE001
            return False

    def _compose_window_still_open(self, subject: str) -> bool:
        """
        True nếu còn cửa sổ *compose* (title có subject + @ / Untitled).
        Tránh khớp nhầm cửa sổ chính khi đang xem Sent.
        """
        needle = _esc((subject or "").strip()[:48])
        if not needle:
            return False
        script = f'''
tell application "System Events"
  tell process "Microsoft Outlook"
    try
      repeat with w in windows
        try
          set n to name of w as text
          if n contains "{needle}" then
            if n contains "@" or n contains "Untitled" then return "yes"
          end if
        end try
      end repeat
    end try
  end tell
end tell
return "no"
'''
        try:
            return _run_osascript_file(script, timeout=15).strip().lower().startswith(
                "yes"
            )
        except Exception:  # noqa: BLE001
            return False

    def _outgoing_with_subject(self, subject: str) -> bool:
        """
        True nếu còn draft/outgoing chưa gửi (cùng subject) và CHƯA nằm trong Sent.
        Quét outgoing + folder Drafts (POP On My Computer).
        """
        want = _esc(subject)
        script = f'''
tell application "Microsoft Outlook"
  set wantSubj to "{want}"
  try
    repeat with m in (get outgoing messages)
      try
        if (subject of m as string) is wantSubj then
          set mid to (id of m as string)
          set alreadySent to false
          try
            repeat with s in (get messages of sent items)
              try
                if ((id of s) as string) is mid then
                  set alreadySent to true
                  exit repeat
                end if
              end try
            end repeat
          end try
          if alreadySent is false then return "yes"
        end if
      end try
    end repeat
  end try
  -- Drafts folder (POP / On My Computer)
  try
    repeat with fld in (get mail folders)
      try
        set nm to name of fld as string
        if nm is "Drafts" or nm is "Draft" then
          repeat with m in (get messages of fld)
            try
              if (subject of m as string) is wantSubj then return "yes"
            end try
          end repeat
        end if
      end try
    end repeat
  end try
  try
    repeat with acc in (get pop accounts)
      try
        repeat with fld in (get mail folders of acc)
          try
            set nm to name of fld as string
            if nm is "Drafts" or nm is "Draft" then
              repeat with m in (get messages of fld)
                try
                  if (subject of m as string) is wantSubj then return "yes"
                end try
              end repeat
            end if
          end try
        end repeat
      end try
    end repeat
  end try
  return "no"
end tell
'''
        try:
            return _run_osascript_file(script, timeout=30).strip().lower().startswith(
                "yes"
            )
        except Exception:  # noqa: BLE001
            return False

    def send(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Path | list[Path]] = None,
    ) -> None:
        """
        Focus Outlook → New Mail + To/Cc/Subject (AS một phát) → paste compose → Send.
        """
        from ..clipboard_html import set_mac_clipboard_html
        from .mac_hid import raise_compose_window, wait_compose_window

        if not ensure_assistive_access(prompt=True):
            raise RuntimeError(
                "Cần Accessibility / Input Monitoring cho «VT Rate Mail Sender».\n"
                "System Settings → Privacy & Security → Accessibility (và Input Monitoring nếu có) → bật app.\n"
                "Outlook: Legacy Outlook = ON."
            )

        prepared = _prepare_mac_paste_body(body_html)
        if not prepared:
            raise ValueError("Nội dung mail trống — không gửi.")
        if len(html_visible_text(prepared)) < 8:
            raise ValueError("Body quá ngắn — kiểm tra Soạn rồi thử lại.")

        to_list, cc_list = _normalize_recipient_lists(to_list, cc_list or [])
        if not to_list:
            raise ValueError("Thiếu / sai địa chỉ To — không gửi.")

        files = self._normalize_files(attachment)
        prev_app = _frontmost_app_name()
        att_block = "\n".join(
            f'  make new attachment at msg with properties '
            f'{{file:POSIX file "{_esc(str(p.resolve()))}"}}'
            for p in files
        )
        recip_block = _applescript_set_recipients(to_list, cc_list)
        has_cc = bool(cc_list)

        # 0) Đóng compose sót — tránh trùng / Draft backup
        self._close_compose_ui_windows()
        set_mac_clipboard_html(prepared)

        # 1) Focus + New Mail + To/Cc/Subject liền (một AppleScript — như trước, chính xác)
        create_script = f'''
tell application "Microsoft Outlook"
  activate
  set msg to make new outgoing message
{_account_block(self.config.from_email)}
  set subject of msg to "{_esc(subject)}"
{recip_block}
{att_block}
  open msg
  return id of msg as string
end tell
'''
        try:
            try:
                msg_id = _run_nsapplescript(create_script).strip()
            except Exception:  # noqa: BLE001
                msg_id = _run_osascript_file(create_script, timeout=30).strip()
        except RuntimeError as exc:
            _activate_app(prev_app)
            raise RuntimeError(
                "Không mở New Mail / gắn To-Cc-Subject.\n"
                f"{exc}\n"
                "Legacy Outlook = ON · đã login."
            ) from exc

        if not msg_id:
            _activate_app(prev_app)
            raise RuntimeError("Không lấy được id New Mail.")

        # 2) Kéo focus Outlook ngay (tránh kẹt app Semi-Auto/Cursor) → AX paste
        from .mac_hid import activate_outlook

        activate_outlook(settle=0.06)
        if not wait_compose_window(timeout=3.0, subject_hint=subject):
            activate_outlook(settle=0.08)
            if not raise_compose_window(subject_hint=subject):
                self._close_compose_ui_windows()
                _activate_app(prev_app)
                raise RuntimeError(
                    "New Mail đã tạo nhưng chưa kéo được focus vào cửa sổ compose.\n"
                    "Đừng đổi Space lúc gửi · Legacy ON · thử lại."
                )

        try:
            set_mac_clipboard_html(prepared)
            time.sleep(0.04)
            self._paste_and_send_burst(has_cc=has_cc, subject=subject)
        except Exception as exc:  # noqa: BLE001
            self._close_compose_ui_windows()
            _activate_app(prev_app)
            raise RuntimeError(
                "Paste/gửi lỗi trên New Mail.\n"
                "Giữ Space lúc Outlook nhảy (AX Subject → Tab → paste → Send).\n"
                f"{exc}"
            ) from exc

        try:
            set_mac_clipboard_html("<div></div>")
        except Exception:  # noqa: BLE001
            pass

        # 3) Compose còn mở → Send thêm 1 lần
        time.sleep(0.2)
        if self._compose_window_still_open(subject):
            try:
                self._ui_send(subject=subject)
                time.sleep(0.3)
            except Exception:  # noqa: BLE001
                pass

        # 4) Trả focus → verify
        _activate_app(prev_app)

        compose_closed = False
        for _ in range(6):
            if not self._compose_window_still_open(subject):
                compose_closed = True
                break
            time.sleep(0.2)

        sent_ok = False
        for _ in range(8):
            if self._message_in_sent(subject, within_minutes=5):
                sent_ok = True
                break
            if compose_closed and not self._outgoing_with_subject(subject):
                sent_ok = True
                break
            time.sleep(0.35)

        if sent_ok:
            return

        if self._compose_window_still_open(subject):
            self._close_compose_ui_windows()
            raise RuntimeError(
                "New Mail vẫn mở — gửi/paste có thể hụt.\n"
                "Giữ Space lúc Outlook nhảy · Legacy ON · thử lại agency này.\n"
                f"To: {', '.join(to_list)} · Subject: {subject}"
            )

        if self._outgoing_with_subject(subject):
            raise RuntimeError(
                "Mail nằm Drafts (chưa vào Sent) — paste/send hụt hoặc mất focus.\n"
                "Xóa draft trong Outlook · giữ Space khi paste · thử lại.\n"
                f"To: {', '.join(to_list)} · Subject: {subject}"
            )

        raise RuntimeError(
            "Không thấy mail trong Sent (đã quét POP/IMAP/Exchange).\n"
            "Kiểm tra Drafts · Legacy Outlook = ON · thử lại agency này.\n"
            f"To: {', '.join(to_list)} · Subject: {subject}"
        )







