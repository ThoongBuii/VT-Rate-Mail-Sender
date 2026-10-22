"""macOS Outlook — New Mail + paste body (tương đương InsertFile Windows).

Windows: Display → WordEditor.InsertFile (ghi thẳng vào body, chữ ký giữ nguyên).
Mac Legacy không có WordEditor. `set content` HTML lớn (bảng + chữ ký base64)
qua AppleScript rất chậm (~phút) và dễ báo SENT khi mail mới vào Outbox.

Cách Mac nhanh + đúng chỗ:
1) New Mail + chữ ký sống Outlook (không merge base64 vào AS)
2) Gắn To·Cc·Subject TRƯỚC — tránh dán nhầm vào ô To
3) Clipboard HTML = chỉ body → click vùng body → Cmd+V (cần Accessibility)
4) Kiểm tra To sạch + body có marker → Send
5) Chờ Outbox xử xong — chưa ra Outbox thì không báo OK

Legacy Outlook = ON.
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
from .html_merge import html_visible_text, merge_body_with_outlook_signature


def signature_dir() -> Path:
    d = user_data_dir() / "mac_signature"
    d.mkdir(parents=True, exist_ok=True)
    return d


def signature_path() -> Path:
    return signature_dir() / "signature.html"


def has_ready_signature() -> bool:
    """Cache dự phòng — không bắt buộc cho Semi-Auto nữa."""
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
        # Thông báo Accessibility ngay khi kết nối — chỉ hiện nếu chưa cấp quyền
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
            "needed": True,
            "message": (
                "Accessibility đã bật."
                if trusted
                else "Cần bật Accessibility cho «VT Rate Mail Sender» "
                "(System Settings → Privacy & Security → Accessibility)."
            ),
        }

    def ensure_accessibility(self, *, prompt: bool = True) -> dict[str, Any]:
        ensure_assistive_access(prompt=prompt)
        return self.accessibility_status()

    def signature_status(self) -> dict[str, Any]:
        """UI parity Windows: không bắt buộc cache — New Mail tự gắn chữ ký."""
        cached = has_ready_signature()
        ax = ensure_assistive_access(prompt=False)
        return {
            "ready": True,  # Semi-Auto không bị chặn
            "path": str(signature_path()) if cached else "",
            "has_file": cached,
            "has_images": cached,
            "cached": cached,
            "accessibility_trusted": ax,
            "message": (
                "Mac: New Mail + chữ ký Outlook (giống Windows). "
                + ("Có cache dự phòng. " if cached else "Cache dự phòng chưa lấy (tuỳ chọn). ")
                + (
                    "Accessibility OK."
                    if ax
                    else "Chưa Accessibility — gửi sẽ chậm / dễ lỗi."
                )
            ),
        }

    def load_signature_html(self) -> str:
        if not has_ready_signature():
            return ""
        return signature_path().read_text(encoding="utf-8", errors="ignore")

    def capture_signature(self) -> str:
        return self.capture_signature_from_new_mail()

    def capture_signature_from_new_mail(self) -> str:
        """Tuỳ chọn: cache chữ ký dự phòng (không bắt buộc Semi-Auto)."""
        from ..clipboard_html import clipboard_html_for_compose, read_mac_clipboard_html

        self.open_outlook()

        existing = ""
        try:
            existing = clipboard_html_for_compose("")
        except Exception:  # noqa: BLE001
            existing = ""
        if (
            len(html_visible_text(existing)) >= 40
            and re.search(r"data:image/", existing, flags=re.I)
            and (
                "best regards" in existing.lower()
                or "vt logistics" in existing.lower()
                or "amber" in existing.lower()
            )
        ):
            signature_path().write_text(existing, encoding="utf-8")
            n_img = len(re.findall(r"data:image/", existing, flags=re.I))
            self._close_vt_sig_windows()
            return (
                f"Đã lưu chữ ký dự phòng từ Clipboard ({len(existing)} ký tự, {n_img} ảnh)."
            )

        trusted = ensure_assistive_access(prompt=True)
        if not trusted:
            raise RuntimeError(
                "«VT Rate Mail Sender» chưa Accessibility.\n"
                "System Settings → Privacy & Security → Accessibility → bật app → thử lại.\n"
                "(Chỉ cần cho bước lấy chữ ký dự phòng.)"
            )

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
            detail = f"\nChi tiết: {last_copy_err}" if last_copy_err else ""
            raise RuntimeError(
                "Chưa Copy được chữ ký dự phòng.\n"
                "Cửa sổ [VT-SIG] vẫn mở — Cmd+A → Cmd+C rồi bấm lại «Chữ ký Outlook»."
                + detail
            )

        signature_path().write_text(html, encoding="utf-8")
        self._close_vt_sig_windows()
        n_img = len(re.findall(r"data:image/", html, flags=re.I))
        return f"Đã lưu chữ ký dự phòng ({len(html)} ký tự, {n_img} ảnh)."

    def _auto_select_copy_compose_body(self, *, y_ratio: float = 0.72) -> None:
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

    def _normalize_files(
        self, attachment: Optional[Path | list[Path]]
    ) -> list[Path]:
        if isinstance(attachment, Path):
            return [attachment]
        if attachment:
            return [p for p in attachment if p is not None]
        return []

    def _close_vt_send_drafts(self, token: str = "") -> None:
        """Đóng/xóa draft [VT-…] còn sót — tránh mail 2+ tìm nhầm draft cũ."""
        try:
            if token:
                cond = f'(subject of m as string) contains "{_esc(token)}"'
            else:
                cond = '(subject of m as string) starts with "[VT-"'
            _run_osascript_file(
                f'''
tell application "Microsoft Outlook"
  try
    repeat with m in (get outgoing messages)
      try
        if {cond} then
          try
            delete m
          end try
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

    def _find_draft_by_token_script(self, token: str) -> str:
        return f'''
  set msg to missing value
  repeat with m in (get outgoing messages)
    try
      if (subject of m as string) contains "{_esc(token)}" then
        set msg to m
        exit repeat
      end if
    end try
  end repeat
  if msg is missing value then error "Không tìm thấy draft {_esc(token)}"
'''

    def _focus_compose_body_and_paste(self, *, y_ratio: float = 0.62) -> None:
        """Click vùng body (không phải To) rồi Cmd+V — cần Accessibility."""
        script = f'''
tell application "Microsoft Outlook" to activate
delay 0.2
tell application "System Events"
  tell process "Microsoft Outlook"
    set frontmost to true
    delay 0.15
    set win to front window
    set {{wx, wy}} to position of win
    set {{ww, wh}} to size of win
    set clickX to (wx + (ww / 2)) as integer
    set clickY to (wy + (wh * {y_ratio})) as integer
    -- Escape bỏ selection ô To nếu đang focus
    key code 53
    delay 0.08
    click at {{clickX, clickY}}
    delay 0.2
    keystroke "v" using {{command down}}
    delay 0.45
  end tell
end tell
return "ok"
'''
        _run_nsapplescript(script)

    def _count_outbox_matching(self, subject: str, to_email: str) -> int:
        """Đếm mail trong Outbox khớp subject/to — 0 = đã gửi xong (hoặc không còn)."""
        script = f'''
tell application "Microsoft Outlook"
  set n to 0
  set wantSubj to "{_esc(subject)}"
  try
    repeat with m in (get messages of outbox)
      try
        set s to subject of m as string
        if s is wantSubj then set n to n + 1
      end try
    end repeat
  on error
    return "-1"
  end try
  return n as string
end tell
'''
        try:
            raw = _run_osascript_file(script, timeout=30)
            return max(0, int((raw or "0").strip() or "0"))
        except Exception:  # noqa: BLE001
            return -1  # không đọc được Outbox

    def _trigger_send_receive(self) -> None:
        try:
            _run_osascript_file(
                '''
tell application "Microsoft Outlook"
  try
    send
  end try
end tell
''',
                timeout=20,
            )
        except Exception:  # noqa: BLE001
            pass
        # Legacy: Cmd+K / menu Send & Receive nếu `send` không có
        try:
            if ensure_assistive_access(prompt=False):
                _run_nsapplescript(
                    '''
tell application "Microsoft Outlook" to activate
delay 0.1
tell application "System Events"
  tell process "Microsoft Outlook"
    keystroke "k" using {command down}
  end tell
end tell
'''
                )
        except Exception:  # noqa: BLE001
            pass

    def _wait_outbox_clear(
        self, subject: str, to_email: str, *, timeout: float = 40.0
    ) -> None:
        """
        send msg chỉ xếp vào Outbox. Chờ Outbox hết mail khớp — nếu kẹt
        (thông báo «A message could not be sent») thì báo lỗi, không ghi SENT OK.
        """
        # Chờ mail kịp vào Outbox trước khi đếm (tránh báo OK sớm)
        time.sleep(1.8)
        self._trigger_send_receive()
        deadline = time.time() + timeout
        last = -1
        while time.time() < deadline:
            last = self._count_outbox_matching(subject, to_email)
            if last == 0:
                return
            if last < 0:
                # Không đọc được folder Outbox — không chặn Semi-Auto
                time.sleep(1.5)
                return
            time.sleep(1.2)
            if time.time() < deadline:
                self._trigger_send_receive()
        raise RuntimeError(
            "Outlook Mac: mail kẹt Outbox / «A message could not be sent».\n"
            f"Subject: {subject}\n"
            f"To: {to_email}\n"
            "Kiểm tra kết nối tài khoản Exchange · Send & Receive · "
            "xóa mail lỗi trong Outbox rồi thử lại.\n"
            f"(Outbox còn ~{last} mail khớp sau {int(timeout)}s)"
        )

    def send(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Path | list[Path]] = None,
    ) -> None:
        """
        New Mail + chữ ký sống → To/Cc trước → paste body vào vùng nội dung → Send.
        Nhanh hơn set content HTML lớn; chờ Outbox xong mới return.
        """
        from ..clipboard_html import set_mac_clipboard_html

        prepared = prepare_body_html_for_outlook(body_html)
        if not (prepared or "").strip():
            raise ValueError("Nội dung mail trống — không gửi.")
        if not to_list:
            raise ValueError("Thiếu địa chỉ To — không gửi.")

        # Span ẩn — Outlook hay xóa HTML comment khi paste; id vẫn giữ được
        marker_id = f"VTBODY-{uuid.uuid4().hex[:10]}"
        body_marker = (
            f'<span id="{marker_id}" style="font-size:0;line-height:0;'
            f'color:transparent;mso-hide:all">.</span>'
        )
        prepared_marked = body_marker + prepared
        marker_token = marker_id

        files = self._normalize_files(attachment)
        prev_app = _frontmost_app_name()
        token = f"[VT-{uuid.uuid4().hex[:8]}]"
        temp_subject = f"{token} drafting"[:180]
        primary_to = to_list[0]

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
        find_draft = self._find_draft_by_token_script(token)

        ax_ok = ensure_assistive_access(prompt=True)
        self._close_vt_send_drafts()

        # 1) New Mail + chữ ký sống + To/Cc/Subject TRƯỚC (chưa đụng body)
        create_script = f'''
tell application "Microsoft Outlook"
  set msg to make new outgoing message
{_account_block(self.config.from_email)}
  set subject of msg to "{_esc(temp_subject)}"
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
{to_block}
{cc_block}
{att_block}
  open msg
  delay 0.7
  return "ok"
end tell
'''
        try:
            try:
                _run_nsapplescript(create_script)
            except Exception:  # noqa: BLE001
                _run_osascript_file(create_script, timeout=90)
        except RuntimeError as exc:
            raise RuntimeError(
                "Outlook Mac không tạo được New Mail.\n"
                f"{exc}\n"
                "Legacy Outlook = ON · đã đăng nhập."
            ) from exc

        pasted = False
        if ax_ok:
            try:
                set_mac_clipboard_html(prepared_marked)
                self._focus_compose_body_and_paste()
                pasted = True
            except Exception:  # noqa: BLE001
                pasted = False

        # Fallback: set content (chậm hơn) — chỉ khi paste thất bại
        if not pasted:
            cache_sig = self.load_signature_html()
            read_sig = f'''
tell application "Microsoft Outlook"
{find_draft}
  set sigHtml to ""
  try
    set sigHtml to content of msg as string
  end try
  return sigHtml
end tell
'''
            try:
                live_sig = _run_osascript_file(read_sig, timeout=60)
            except Exception:  # noqa: BLE001
                live_sig = ""
            sig_html = live_sig or cache_sig or ""
            full = merge_body_with_outlook_signature(prepared_marked, sig_html)
            with tempfile.TemporaryDirectory() as tmp:
                full_path = Path(tmp) / "full.html"
                full_path.write_text(full, encoding="utf-8")
                set_script = f'''
set fullPath to "{full_path}"
set fullText to do shell script "cat " & quoted form of fullPath
tell application "Microsoft Outlook"
{find_draft}
  set content of msg to fullText
end tell
'''
                try:
                    _run_osascript_file(set_script, timeout=120)
                except RuntimeError as exc:
                    self._close_vt_send_drafts(token)
                    raise RuntimeError(
                        "Không ghi được body (paste + set content đều lỗi).\n"
                        f"{exc}\n"
                        "Bật Accessibility cho app rồi thử lại."
                    ) from exc

        # 2) Đặt subject thật + kiểm tra To/body + Send
        finish_script = f'''
tell application "Microsoft Outlook"
{find_draft}

  -- Làm sạch To/Cc rồi gắn lại (tránh rác nếu paste lệch)
  try
    delete (every to recipient of msg)
  end try
  try
    delete (every cc recipient of msg)
  end try
{to_block}
{cc_block}
  set subject of msg to "{_esc(subject)}"

  set bodyCheck to ""
  try
    set bodyCheck to content of msg as string
  end try
  if bodyCheck does not contain "{_esc(marker_token)}" then error "Body chưa có nội dung soạn — hủy Send"
  if (length of bodyCheck) < 120 then error "Body quá ngắn — hủy Send"

  set toCount to 0
  try
    set toCount to count of (to recipients of msg)
  end try
  if toCount < 1 then error "To trống — hủy gửi"

  try
    set firstTo to address of email address of item 1 of (to recipients of msg) as string
    if firstTo contains "MicrosoftInternetExplorer" then error "To bị dán nhầm nội dung — hủy"
    if firstTo does not contain "@" then error "To không phải email — hủy"
    if (length of firstTo) > 120 then error "To quá dài (có vẻ dán nhầm body) — hủy"
  end try

  send msg
end tell
'''
        try:
            _run_osascript_file(finish_script, timeout=90)
        except RuntimeError as exc:
            self._close_vt_send_drafts(token)
            raise RuntimeError(
                "Outlook Mac gửi thất bại (đã chặn Outbox trống / To bẩn).\n"
                f"{exc}\n"
                "Legacy Outlook = ON · Accessibility cho paste body."
            ) from exc

        # 3) Chờ Outbox — tránh báo SENT khi Outlook hiện «could not be sent»
        try:
            self._wait_outbox_clear(subject, primary_to, timeout=45.0)
        except RuntimeError:
            self._close_vt_send_drafts(token)
            raise

        time.sleep(0.2)
        self._close_vt_send_drafts(token)
        _activate_app(prev_app)
