from __future__ import annotations

import platform
import subprocess
import tempfile
import time
import uuid
from pathlib import Path
from typing import Any, Optional

from .models import AgencyMail, AppConfig
from .sender_smtp import split_emails
from .template_engine import render_body_html, render_subject

TEMPLATE_SUBJECT_PREFIX = "[VT-TEMPLATE]"


def uuid_hex() -> str:
    return uuid.uuid4().hex


class OutlookDesktopSender:
    """
    Gửi qua Microsoft Outlook đã đăng nhập.
    App chỉ soạn nội dung (Dear / bảng giá / remark).
    Khi gửi, Outlook New Mail tự gắn chữ ký mặc định của account.
    """

    def __init__(self, config: AppConfig):
        self.config = config
        self._ready = False
        self._account_label = ""
        self._template_draft_id: str = ""
        self._win_template_item = None  # Windows COM mail item ref

    @property
    def account_email(self) -> str:
        return self._account_label or self.config.from_email

    @property
    def is_configured(self) -> bool:
        return True

    @property
    def is_ready(self) -> bool:
        return self._ready

    def open_outlook(self) -> str:
        system = platform.system()
        if system == "Darwin":
            return self._open_mac()
        if system == "Windows":
            return self._open_windows()
        raise RuntimeError(f"Hệ điều hành chưa hỗ trợ: {system}")

    def _open_mac(self) -> str:
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
                "Cần Classic/Legacy Outlook + đã đăng nhập.\n"
                f"Chi tiết: {err or 'osascript failed'}"
            )
        self._ready = True
        self._account_label = self.config.from_email or "Outlook (Mac)"
        return f"Outlook đã mở · {self._account_label}"

    def _open_windows(self) -> str:
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise RuntimeError("Thiếu pywin32. Chạy: pip install pywin32") from exc
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        accounts: list[str] = []
        try:
            for i in range(1, namespace.Accounts.Count + 1):
                accounts.append(namespace.Accounts.Item(i).SmtpAddress)
        except Exception:  # noqa: BLE001
            pass
        self._ready = True
        self._account_label = (
            accounts[0] if accounts else (self.config.from_email or "Outlook (Windows)")
        )
        return f"Outlook đã sẵn sàng · {self._account_label}"

    def test_connection(self) -> str:
        return self.open_outlook()

    def clear_credentials(self) -> None:
        self._ready = False

    def resolve_attachment(self, mail: AgencyMail) -> Optional[Path]:
        raw = (mail.attachment or "").strip()
        if not raw:
            return None
        path = Path(raw)
        if path.is_file():
            return path
        if self.config.default_attachment_dir:
            candidate = Path(self.config.default_attachment_dir) / raw
            if candidate.is_file():
                return candidate
        root = Path(__file__).resolve().parent.parent
        try:
            from .paths import user_data_dir

            root = user_data_dir()
        except Exception:  # noqa: BLE001
            pass
        candidate = root / raw
        if candidate.is_file():
            return candidate
        raise FileNotFoundError(f"Không tìm thấy file đính kèm: {raw}")

    def preview(self, mail: AgencyMail) -> dict[str, Any]:
        subject = render_subject(mail)
        body_html = render_body_html(mail, "")
        try:
            att = self.resolve_attachment(mail)
            att_name = att.name if att else "(không có)"
            att_ok = True
            att_error = ""
        except FileNotFoundError as exc:
            att_name = mail.attachment
            att_ok = False
            att_error = str(exc)
        return {
            "to": mail.account_mail,
            "cc": mail.mail_cc,
            "subject": subject,
            "body_html": body_html,
            "attachment": att_name,
            "attachment_ok": att_ok,
            "attachment_error": att_error,
            "agency_company": mail.agency_company,
            "account_name": mail.account_name,
            "from": self.account_email or self.config.from_email or "Outlook default account",
            "signature_note": "Chữ ký mặc định Outlook sẽ tự gắn khi gửi (giống New Mail).",
        }

    # -------------------- Template compose in Outlook New Mail --------------------
    def open_template_composer(self, html_body: str, subject_hint: str = "") -> str:
        """Mở cửa sổ New Mail để soạn/dán bảng + chữ ký. Không gửi."""
        if not self._ready:
            self.open_outlook()
        subj = f"{TEMPLATE_SUBJECT_PREFIX} {subject_hint or 'VT Rate body'}".strip()
        system = platform.system()
        if system == "Darwin":
            return self._open_template_mac(html_body, subj)
        if system == "Windows":
            return self._open_template_windows(html_body, subj)
        raise RuntimeError(f"Hệ điều hành chưa hỗ trợ: {system}")

    def sync_template_from_outlook(self) -> str:
        """Lấy HTML đã soạn từ nháp/cửa sổ template Outlook."""
        if not self._ready:
            self.open_outlook()
        system = platform.system()
        if system == "Darwin":
            return self._sync_template_mac()
        if system == "Windows":
            return self._sync_template_windows()
        raise RuntimeError(f"Hệ điều hành chưa hỗ trợ: {system}")

    def _open_template_mac(self, html_body: str, subject: str) -> str:
        with tempfile.TemporaryDirectory() as tmp:
            html_path = Path(tmp) / "template.html"
            html_path.write_text(html_body or "<div></div>", encoding="utf-8")
            # Use shell to load file → avoid escaping hell
            script = f'''
set htmlPath to "{html_path}"
set htmlText to do shell script "cat " & quoted form of htmlPath
tell application "Microsoft Outlook"
  activate
  set msg to make new outgoing message
  set subject of msg to "{subject.replace('"', '\\"')}"
  try
    set content of msg to htmlText
  on error
    set plain text content of msg to htmlText
  end try
  open msg
  try
    return id of msg as string
  on error
    return ""
  end try
end tell
'''
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                timeout=120,
            )
        if result.returncode != 0:
            raise RuntimeError(
                "Không mở được cửa sổ soạn Outlook.\n"
                + (result.stderr or result.stdout or "")
            )
        self._template_draft_id = (result.stdout or "").strip()
        return (
            "Đã mở New Mail trong Outlook.\n\n"
            "1) Dán bảng giá + chữ ký như soạn tay\n"
            "2) Giữ nguyên subject có [VT-TEMPLATE]\n"
            "3) KHÔNG bấm Send — Save/đóng cửa sổ hoặc giữ mở\n"
            "4) Quay lại app bấm “Đồng bộ từ Outlook”"
        )

    def _sync_template_mac(self) -> str:
        # 1) Try by saved id
        if self._template_draft_id:
            script = f'''
tell application "Microsoft Outlook"
  try
    set msg to message id {self._template_draft_id}
    try
      return content of msg
    on error
      return plain text content of msg
    end try
  end try
end tell
'''
            result = subprocess.run(
                ["osascript", "-e", script], capture_output=True, text=True, timeout=60
            )
            if result.returncode == 0 and (result.stdout or "").strip():
                return result.stdout.strip()

        # 2) Search drafts / outgoing by subject prefix
        script = f'''
tell application "Microsoft Outlook"
  set collected to ""
  try
    set pool to drafts
    repeat with m in pool
      try
        if subject of m contains "{TEMPLATE_SUBJECT_PREFIX}" then
          try
            set collected to content of m
          on error
            set collected to plain text content of m
          end try
          exit repeat
        end if
      end try
    end repeat
  end try
  if collected is "" then
    try
      set pool2 to outgoing messages
      repeat with m in pool2
        try
          if subject of m contains "{TEMPLATE_SUBJECT_PREFIX}" then
            try
              set collected to content of m
            on error
              set collected to plain text content of m
            end try
            exit repeat
          end if
        end try
      end repeat
    end try
  end if
  return collected
end tell
'''
        result = subprocess.run(
            ["osascript", "-e", script], capture_output=True, text=True, timeout=90
        )
        if result.returncode != 0:
            raise RuntimeError(
                "Đồng bộ thất bại.\n" + (result.stderr or result.stdout or "")
            )
        html_body = (result.stdout or "").strip()
        if not html_body:
            raise RuntimeError(
                "Không tìm thấy nháp [VT-TEMPLATE].\n"
                "Hãy mở lại “Soạn trong Outlook”, dán nội dung, Save, rồi Đồng bộ."
            )
        return html_body

    def _open_template_windows(self, html_body: str, subject: str) -> str:
        import win32com.client  # type: ignore

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        if self.config.from_email:
            try:
                namespace = outlook.GetNamespace("MAPI")
                for i in range(1, namespace.Accounts.Count + 1):
                    acc = namespace.Accounts.Item(i)
                    smtp = getattr(acc, "SmtpAddress", "") or ""
                    if smtp.lower() == self.config.from_email.lower():
                        mail.SendUsingAccount = acc
                        break
            except Exception:  # noqa: BLE001
                pass
        mail.Subject = subject
        # Để trống hoặc HTML đơn giản — user dán bảng + chữ ký trong Outlook (giữ nguyên format).
        body = (html_body or "").strip()
        if not body or body in ("<div></div>", "<p></p>", "<p><br></p>"):
            mail.Body = ""
            mail.HTMLBody = "<div><br></div>"
        else:
            mail.HTMLBody = body
        self._win_template_item = mail
        mail.Display(False)
        return (
            "Đã mở New Mail trong Outlook.\n\n"
            "1) Trong Outlook: dán bảng giá + chữ ký như gửi tay (giữ format)\n"
            "2) Giữ subject có [VT-TEMPLATE]\n"
            "3) KHÔNG bấm Send\n"
            "4) Quay lại app → bấm “Đồng bộ từ Outlook”"
        )

    def _windows_html_with_inline_images(self, mail_item: Any) -> str:
        """
        Lấy HTMLBody và nhúng ảnh cid:/file đính kèm → data URI để Preview/gửi giữ đúng chữ ký.
        """
        import base64
        import re
        import tempfile

        from .clipboard_html import embed_local_and_cid_images

        html = str(getattr(mail_item, "HTMLBody", None) or "")
        if not html.strip():
            return html

        cid_map: dict[str, str] = {}
        try:
            count = int(mail_item.Attachments.Count)
        except Exception:  # noqa: BLE001
            count = 0

        prop_w = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
        prop_a = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"

        for i in range(1, count + 1):
            try:
                att = mail_item.Attachments.Item(i)
            except Exception:  # noqa: BLE001
                continue

            filename = str(getattr(att, "FileName", None) or f"image_{i}.png")
            tmp = Path(tempfile.gettempdir()) / f"vt_cid_{uuid_hex()}_{filename}"
            try:
                att.SaveAsFile(str(tmp))
                raw = tmp.read_bytes()
            except Exception:  # noqa: BLE001
                continue
            finally:
                try:
                    tmp.unlink(missing_ok=True)
                except Exception:  # noqa: BLE001
                    pass

            lower = filename.lower()
            if lower.endswith((".jpg", ".jpeg")):
                mime = "image/jpeg"
            elif lower.endswith(".gif"):
                mime = "image/gif"
            elif lower.endswith(".bmp"):
                mime = "image/bmp"
            elif lower.endswith(".webp"):
                mime = "image/webp"
            else:
                mime = "image/png"
            data_uri = f"data:{mime};base64,{base64.b64encode(raw).decode('ascii')}"

            cid = ""
            try:
                pa = att.PropertyAccessor
                try:
                    cid = str(pa.GetProperty(prop_w) or "")
                except Exception:  # noqa: BLE001
                    try:
                        cid = str(pa.GetProperty(prop_a) or "")
                    except Exception:  # noqa: BLE001
                        cid = ""
            except Exception:  # noqa: BLE001
                cid = ""
            cid = cid.strip().strip("<>")
            if cid:
                cid_map[cid.lower()] = data_uri
            # Luôn map theo tên file (Outlook signature_* thường chỉ khớp filename)
            stem = Path(filename).stem.lower()
            cid_map[filename.lower()] = data_uri
            cid_map[stem] = data_uri
            cid_map[f"cid:{stem}"] = data_uri

        return embed_local_and_cid_images(html, cid_map)

    def _sync_template_windows(self) -> str:
        import win32com.client  # type: ignore

        outlook = win32com.client.Dispatch("Outlook.Application")

        # 1) Cửa sổ soạn đang mở (Inspectors) — mới nhất, đủ ảnh chữ ký
        try:
            inspectors = outlook.Inspectors
            for i in range(1, int(inspectors.Count) + 1):
                try:
                    item = inspectors.Item(i).CurrentItem
                    if int(getattr(item, "Class", 0)) != 43:  # olMail
                        continue
                    subj = str(getattr(item, "Subject", "") or "")
                    if TEMPLATE_SUBJECT_PREFIX not in subj:
                        continue
                    html = self._windows_html_with_inline_images(item)
                    if html.strip():
                        self._win_template_item = item
                        return html
                except Exception:  # noqa: BLE001
                    continue
        except Exception:  # noqa: BLE001
            pass

        # 2) Item app đã Display trước đó
        if self._win_template_item is not None:
            try:
                html = self._windows_html_with_inline_images(self._win_template_item)
                if html.strip():
                    return html
            except Exception:  # noqa: BLE001
                pass

        # 3) Drafts
        namespace = outlook.GetNamespace("MAPI")
        drafts = namespace.GetDefaultFolder(16)  # olFolderDrafts
        items = drafts.Items
        items.Sort("[LastModificationTime]", True)
        for i in range(1, min(int(items.Count), 50) + 1):
            it = items.Item(i)
            subj = str(getattr(it, "Subject", "") or "")
            if TEMPLATE_SUBJECT_PREFIX in subj:
                html = self._windows_html_with_inline_images(it)
                if html.strip():
                    return html

        raise RuntimeError(
            "Không tìm thấy thư [VT-TEMPLATE].\n"
            "Hãy bấm “Soạn trong Outlook”, dán nội dung + chữ ký, "
            "giữ cửa sổ mở (hoặc Save Draft), rồi “Đồng bộ từ Outlook”."
        )

    # --------------------------------- send ---------------------------------
    def send_one(self, mail: AgencyMail) -> None:
        if not self._ready:
            self.open_outlook()
        errors = mail.validate()
        if errors:
            raise ValueError("; ".join(errors))

        to_list = split_emails(mail.account_mail)
        if not to_list:
            raise ValueError("Thiếu Account Mail (To).")

        subject = render_subject(mail)
        body_html = render_body_html(mail, "")
        cc_list = split_emails(mail.mail_cc)
        attachment = self.resolve_attachment(mail)

        system = platform.system()
        if system == "Darwin":
            self._send_mac(to_list, cc_list, subject, body_html, attachment)
        elif system == "Windows":
            self._send_windows(to_list, cc_list, subject, body_html, attachment)
        else:
            raise RuntimeError(f"Hệ điều hành chưa hỗ trợ: {system}")

    @staticmethod
    def _mac_set_html_clipboard(html: str) -> None:
        """Đặt HTML lên clipboard Mac để dán vào New Mail (giữ chữ ký UI)."""
        import re

        try:
            from AppKit import NSPasteboard, NSPasteboardTypeHTML, NSPasteboardTypeString  # type: ignore
        except ImportError as exc:
            raise RuntimeError(
                "Thiếu PyObjC (AppKit) để dán HTML vào Outlook Mac.\n"
                "Dùng bản build macOS hoặc: pip install pyobjc-framework-Cocoa"
            ) from exc

        plain = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", html or "")
        plain = re.sub(r"(?is)<br\s*/?>", "\n", plain)
        plain = re.sub(r"(?is)</p>", "\n", plain)
        plain = re.sub(r"<[^>]+>", " ", plain)
        plain = re.sub(r"[ \t]+\n", "\n", plain)
        plain = re.sub(r"\n{3,}", "\n\n", plain)
        plain = re.sub(r"[ \t]{2,}", " ", plain).strip()

        pb = NSPasteboard.generalPasteboard()
        pb.clearContents()
        ok_html = pb.setString_forType_(html, NSPasteboardTypeHTML)
        ok_plain = pb.setString_forType_(plain or " ", NSPasteboardTypeString)
        if not ok_html and not ok_plain:
            raise RuntimeError("Không ghi được clipboard HTML trên macOS.")

    def _send_mac(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Path],
    ) -> None:
        """
        Legacy Outlook Mac — 1 New Mail chuẩn:
        To/Cc/Subject (AppleScript) → open (chữ ký UI) → focus BODY → Cmd+V
        → xác nhận body đã dán → mới send.
        Không set content (xóa chữ ký). Không dán vào Cc/To.
        """
        import re

        from .outlook_html import prepare_body_html_for_outlook

        prepared = prepare_body_html_for_outlook(body_html)
        if not (prepared or "").strip():
            raise ValueError("Nội dung mail trống — không gửi.")
        if not to_list:
            raise ValueError("Thiếu địa chỉ To — không gửi.")

        def esc(s: str) -> str:
            return (s or "").replace("\\", "\\\\").replace('"', '\\"')

        # Chuỗi xác nhận body đã vào content (không dán nhầm Cc).
        probe_raw = self._html_visible_text(prepared)
        probe_raw = re.sub(r"\s+", " ", probe_raw).strip()
        probe = ""
        for chunk in (probe_raw, "Dear", "Good day"):
            c = (chunk or "").strip()
            if len(c) >= 4:
                probe = c[:28].replace('"', "").replace("\\", "")
                break
        if not probe:
            probe = "Dear"

        to_block = "\n".join(
            f'  make new to recipient at msg with properties {{email address:{{address:"{esc(t)}"}}}}'
            for t in to_list
        )
        cc_block = "\n".join(
            f'  make new cc recipient at msg with properties {{email address:{{address:"{esc(cc)}"}}}}'
            for cc in cc_list
        )
        if attachment:
            att = esc(str(attachment.resolve()))
            att_block = (
                f'  make new attachment at msg with properties {{file:POSIX file "{att}"}}'
            )
        else:
            att_block = ""

        account_block = ""
        if self.config.from_email:
            fe = esc(self.config.from_email)
            account_block = (
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

        self._mac_set_html_clipboard(prepared)

        # AppleScript phức tạp → ghi file tạm (handler đệ quy tìm body).
        script = f'''-- VT Rate Mail Sender — Mac Legacy Outlook send
on collectTextAreas(elem)
  set bag to {{}}
  try
    set bag to bag & (every text area of elem)
  end try
  try
    repeat with g in (every group of elem)
      set bag to bag & my collectTextAreas(g)
    end repeat
  end try
  try
    repeat with s in (every scroll area of elem)
      set bag to bag & my collectTextAreas(s)
    end repeat
  end try
  try
    repeat with s in (every splitter group of elem)
      set bag to bag & my collectTextAreas(s)
    end repeat
  end try
  try
    repeat with s in (every splitter of elem)
      set bag to bag & my collectTextAreas(s)
    end repeat
  end try
  return bag
end collectTextAreas

on collectWebAreas(elem)
  set bag to {{}}
  try
    set bag to bag & (every UI element of elem whose role is "AXWebArea")
  end try
  try
    repeat with g in (every group of elem)
      set bag to bag & my collectWebAreas(g)
    end repeat
  end try
  try
    repeat with s in (every scroll area of elem)
      set bag to bag & my collectWebAreas(s)
    end repeat
  end try
  try
    repeat with s in (every splitter group of elem)
      set bag to bag & my collectWebAreas(s)
    end repeat
  end try
  return bag
end collectWebAreas

on focusComposeBody()
  tell application "System Events"
    if not UI elements enabled then
      error "Cần bật Accessibility: System Settings → Privacy & Security → Accessibility → VT Rate Mail Sender / Terminal / osascript"
    end if
    tell process "Microsoft Outlook"
      set frontmost to true
      delay 0.35
      set win to front window
      set focusedBody to false

      -- 1) Ưu tiên AXWebArea (HTML body) — phần CUỐI CÙNG, không lấy first (To/Cc)
      try
        set webs to my collectWebAreas(win)
        if (count of webs) > 0 then
          click item -1 of webs
          set focusedBody to true
        end if
      end try

      -- 2) Text area cuối cùng (header fields = đầu; body = cuối)
      if focusedBody is false then
        try
          set areas to my collectTextAreas(win)
          if (count of areas) > 0 then
            click item -1 of areas
            set focusedBody to true
          end if
        end try
      end if

      -- 3) Scroll area cuối
      if focusedBody is false then
        try
          set scrolls to every scroll area of win
          if (count of scrolls) > 0 then
            click item -1 of scrolls
            set focusedBody to true
          end if
        end try
      end if

      -- 4) Fallback Tab: To → Cc → Subject → Body (~3 lần từ To)
      if focusedBody is false then
        repeat 3 times
          keystroke tab
          delay 0.12
        end repeat
      end if

      delay 0.25
      -- Caret về đầu body (trên chữ ký)
      key code 126 using {{command down}}
      delay 0.12
      keystroke "v" using {{command down}}
      delay 1.2
    end tell
  end tell
end focusComposeBody

tell application "Microsoft Outlook"
  activate
  set msg to make new outgoing message
{account_block}
  set subject of msg to "{esc(subject)}"
{to_block}
{cc_block}
{att_block}
  open msg
  delay 1.7
end tell

my focusComposeBody()

-- Xác nhận body trên UI (phòng content AppleScript chưa kịp sync)
set uiBody to ""
tell application "System Events"
  tell process "Microsoft Outlook"
    try
      set areas to my collectTextAreas(front window)
      if (count of areas) > 0 then
        try
          set uiBody to (value of item -1 of areas) as text
        end try
      end if
    end try
  end tell
end tell

tell application "Microsoft Outlook"
  -- Xác nhận body đã dán TRƯỚC khi send (tránh SENT ảo / Outbox trống)
  set checkText to ""
  try
    set checkText to plain text content of msg
  end try
  if checkText is "" then
    try
      set checkText to content of msg
    end try
  end if
  set okBody to false
  if checkText contains "{esc(probe)}" then set okBody to true
  if uiBody contains "{esc(probe)}" then set okBody to true
  if okBody is false then
    error "Body chưa dán đúng (có thể đang dán nhầm Cc). Không gửi — sửa cửa sổ New Mail rồi thử lại 1 mail."
  end if

  set toCount to 0
  try
    set toCount to count of (to recipients of msg)
  end try
  if toCount < 1 then error "To trống — hủy gửi"

  send msg
end tell
'''
        with tempfile.TemporaryDirectory() as tmp:
            script_path = Path(tmp) / "vt_send_mac.applescript"
            script_path.write_text(script, encoding="utf-8")
            result = subprocess.run(
                ["osascript", str(script_path)],
                capture_output=True,
                text=True,
                timeout=180,
            )
        if result.returncode != 0:
            raise RuntimeError(
                "Outlook Mac gửi thất bại (đã hủy send nếu body chưa đúng).\n"
                + (result.stderr or result.stdout or "")
                + "\nGợi ý:\n"
                "- Xóa Outbox + đóng Untitled bị dán nhầm Cc\n"
                "- Legacy Outlook ON\n"
                "- Accessibility: cho phép VT Rate Mail Sender\n"
                "- Thử lại 1 mail"
            )

    def _windows_pick_account(self, outlook: Any):
        if not self.config.from_email:
            return None
        try:
            namespace = outlook.GetNamespace("MAPI")
            for i in range(1, namespace.Accounts.Count + 1):
                acc = namespace.Accounts.Item(i)
                smtp = getattr(acc, "SmtpAddress", "") or ""
                if smtp.lower() == self.config.from_email.lower():
                    return acc
        except Exception:  # noqa: BLE001
            return None
        return None

    @staticmethod
    def _html_visible_text(inner: str) -> str:
        import re

        text = re.sub(r"(?is)<style[^>]*>.*?</style>", "", inner or "")
        text = re.sub(r"(?is)<!--.*?-->", "", text)
        text = re.sub(r"(?i)<br\s*/?>", "\n", text)
        text = re.sub(r"<[^>]+>", "", text)
        text = (
            text.replace("\xa0", " ")
            .replace("&nbsp;", " ")
            .replace("&#160;", " ")
            .replace("&amp;", "&")
        )
        return text.strip()

    @staticmethod
    def _is_empty_html_block(inner: str) -> bool:
        return OutlookDesktopSender._html_visible_text(inner) == ""

    @staticmethod
    def _inject_style_attr(open_tag: str, css: str) -> str:
        """Thêm/ghi đè style vào thẻ mở <p ...> / <div ...>."""
        import re

        css = css.strip().rstrip(";")
        if re.search(r"\sstyle\s*=", open_tag, flags=re.I):

            def _merge(m: Any) -> str:
                q = m.group(1)
                old = (m.group(2) or "").strip().rstrip(";")
                merged = f"{old};{css}" if old else css
                return f" style={q}{merged}{q}"

            return re.sub(
                r"\sstyle\s*=\s*(['\"])(.*?)\1",
                _merge,
                open_tag,
                count=1,
                flags=re.I,
            )
        if open_tag.endswith(">"):
            return open_tag[:-1] + f' style="{css}"' + ">"
        return open_tag

    @staticmethod
    def _strip_edge_empty_html(html: str, *, leading: bool) -> str:
        """
        Bỏ khối HTML trống ở đầu/cuối (p/div lá rỗng, br, &nbsp;).
        Dùng pattern lá neo biên — tránh regex nested Word HTML.
        """
        import re

        s = html or ""
        leaf_empty = (
            r"<(?:p|div)(?:\s[^>]*)?>\s*(?:"
            r"<br\s*/?>|&nbsp;|\xa0|&#160;|\s|"
            r"<span[^>]*>\s*(?:&nbsp;|\xa0|&#160;|<br\s*/?>|\s|"
            r"<o:p[^>]*>\s*(?:&nbsp;|\xa0|&#160;)?\s*</o:p>)*</span>|"
            r"<o:p[^>]*>\s*(?:&nbsp;|\xa0|&#160;)?\s*</o:p>"
            r")*\s*</(?:p|div)>"
        )
        if leading:
            leaf_re = re.compile(rf"^\s*(?:{leaf_empty})\s*", flags=re.I)
            br_re = re.compile(r"^(?:<br\s*/?>|&nbsp;|\xa0|&#160;|\s)+", flags=re.I)
            wrap_re = re.compile(
                r"^(<div[^>]*(?:WordSection|elementToProof|OutlookMessageBody)[^>]*>)\s*",
                flags=re.I,
            )
        else:
            leaf_re = re.compile(
                rf"(?:{leaf_empty})((?:\s*</div>)*)\s*$",
                flags=re.I,
            )
            br_re = re.compile(r"(?:<br\s*/?>\s*|&nbsp;|\xa0|&#160;|\s)+$", flags=re.I)
            wrap_re = None

        changed = True
        while changed:
            changed = False
            if leading:
                m = wrap_re.match(s) if wrap_re else None
                if m:
                    return m.group(1) + OutlookDesktopSender._strip_edge_empty_html(
                        s[m.end() :], leading=True
                    )
                m = leaf_re.match(s)
                if m:
                    s = s[m.end() :]
                    changed = True
                    continue
                m = br_re.match(s)
                if m:
                    s = s[m.end() :]
                    changed = True
                    continue
            else:
                m = leaf_re.search(s)
                if m:
                    # Bỏ khối trống, giữ các </div> đóng wrapper phía sau
                    s = s[: m.start()] + (m.group(1) or "")
                    changed = True
                    continue
                m = br_re.search(s)
                if m:
                    s = s[: m.start()]
                    changed = True
                    continue
        return s

    @staticmethod
    def _zero_first_block_margin(html: str) -> str:
        import re

        s = html or ""
        m2 = re.match(
            r"^(<div[^>]*(?:WordSection|elementToProof|OutlookMessageBody)[^>]*>)(\s*)",
            s,
            flags=re.I,
        )
        if m2:
            return (
                m2.group(1)
                + m2.group(2)
                + OutlookDesktopSender._zero_first_block_margin(s[m2.end() :])
            )

        m = re.match(r"^(<(?:p|div)(?:\s[^>]*)?)>", s, flags=re.I)
        if not m:
            return s
        open_tag = OutlookDesktopSender._inject_style_attr(
            m.group(1) + ">",
            "margin-top:0;padding-top:0",
        )
        return open_tag + s[m.end() :]

    @staticmethod
    def _zero_last_block_margin(html: str) -> str:
        import re

        s = html or ""
        opens = list(re.finditer(r"<(p|div)(\s[^>]*)?>", s, flags=re.I))
        for m in reversed(opens):
            tag = m.group(1)
            close = re.search(rf"</{tag}\s*>", s[m.end() :], flags=re.I)
            if not close:
                continue
            # Chỉ nhận closing gần nhất và không có block con
            inner = s[m.end() : m.end() + close.start()]
            if re.search(r"<(?:p|div)\b", inner, flags=re.I):
                continue
            if OutlookDesktopSender._is_empty_html_block(inner):
                continue
            # Phải là block lá cuối (sau closing chỉ còn đóng wrapper/whitespace)
            after = s[m.end() + close.end() :]
            if re.search(r"<(?:p|div|table|br)\b", after, flags=re.I):
                continue
            new_open = OutlookDesktopSender._inject_style_attr(
                m.group(0), "margin-bottom:0;padding-bottom:0"
            )
            return s[: m.start()] + new_open + s[m.end() :]
        return s

    @staticmethod
    def _merge_body_with_outlook_signature(body_html: str, signature_html: str) -> str:
        """
        Chèn nội dung app sát chữ ký Outlook (0 khoảng mặc định).
        Khoảng trống chỉ còn nếu user tự Enter trong app.
        """
        import re

        body = OutlookDesktopSender._strip_edge_empty_html(
            (body_html or "").strip(), leading=False
        )
        sig = (signature_html or "").strip()
        if not body:
            return sig
        if not sig:
            return body

        body = OutlookDesktopSender._zero_last_block_margin(body)
        # Wrapper ngoài cùng của app cũng không tạo khoảng dưới
        import re as _re

        _mwrap = _re.match(r"^(<div(?:\s[^>]*)?)>", body, flags=_re.I)
        if _mwrap:
            body = (
                OutlookDesktopSender._inject_style_attr(
                    _mwrap.group(1) + ">",
                    "margin-bottom:0;padding-bottom:0",
                )
                + body[_mwrap.end() :]
            )

        m = re.search(r"<body[^>]*>", sig, flags=re.I)
        if m:
            i = m.end()
            rest = OutlookDesktopSender._strip_edge_empty_html(sig[i:], leading=True)
            # Chèn body vào trong WordSection (cùng vùng soạn) — sát chữ ký hơn
            wm = re.match(
                r"(<div[^>]*(?:WordSection|elementToProof|OutlookMessageBody)[^>]*>)\s*",
                rest,
                flags=re.I,
            )
            if wm:
                inner = OutlookDesktopSender._strip_edge_empty_html(
                    rest[wm.end() :], leading=True
                )
                inner = OutlookDesktopSender._zero_first_block_margin(inner)
                return sig[:i] + wm.group(1) + body + inner

            rest = OutlookDesktopSender._zero_first_block_margin(rest)
            return sig[:i] + body + rest

        rest = OutlookDesktopSender._strip_edge_empty_html(sig, leading=True)
        rest = OutlookDesktopSender._zero_first_block_margin(rest)
        return body + rest

    def _windows_delete_leading_empty_paragraphs(self, word_doc: Any, limit: int = 30) -> int:
        """Xóa đoạn trống đầu New Mail (thường ~2 dòng trước chữ ký)."""
        removed = 0
        for _ in range(limit):
            try:
                if int(word_doc.Paragraphs.Count) < 1:
                    break
                para = word_doc.Paragraphs(1)
                raw = str(para.Range.Text or "")
                clean = (
                    raw.replace("\r", "")
                    .replace("\x07", "")
                    .replace("\xa0", " ")
                    .replace("\u200b", "")
                    .strip()
                )
                if clean:
                    break
                para.Range.Delete()
                removed += 1
            except Exception:  # noqa: BLE001
                break
        return removed

    def _windows_trim_gap_before_signature(self, word_doc: Any) -> int:
        """Xóa đoạn trống ngay trước 'Best Regards' / chữ ký (gap ~2 dòng)."""
        import re

        removed = 0
        try:
            count = int(word_doc.Paragraphs.Count)
        except Exception:  # noqa: BLE001
            return 0

        sig_idx = None
        for i in range(1, count + 1):
            try:
                raw = str(word_doc.Paragraphs(i).Range.Text or "")
            except Exception:  # noqa: BLE001
                continue
            if re.search(r"best\s*regards", raw, flags=re.I):
                sig_idx = i
                break

        if not sig_idx or sig_idx <= 1:
            return 0

        while sig_idx > 1 and removed < 20:
            try:
                prev = word_doc.Paragraphs(sig_idx - 1)
                raw = str(prev.Range.Text or "")
                clean = (
                    raw.replace("\r", "")
                    .replace("\x07", "")
                    .replace("\xa0", " ")
                    .replace("\u200b", "")
                    .strip()
                )
                if clean:
                    break
                prev.Range.Delete()
                removed += 1
                sig_idx -= 1
            except Exception:  # noqa: BLE001
                break
        return removed

    def _windows_apply_body_keep_signature(self, mail_item: Any, body_html: str) -> None:
        """
        Chèn nội dung HTML vào thân mail (không đính file .htm),
        giữ chữ ký Outlook + cắt khoảng trống đầu / trước Best Regards.
        """
        word_doc = None
        try:
            insp = mail_item.GetInspector
            word_doc = insp.WordEditor
        except Exception:  # noqa: BLE001
            word_doc = None

        if word_doc is not None:
            self._windows_delete_leading_empty_paragraphs(word_doc)

        existing = str(getattr(mail_item, "HTMLBody", None) or "")
        mail_item.HTMLBody = self._merge_body_with_outlook_signature(body_html, existing)

        # Cắt gap sau khi merge; đồng thời bỏ mọi attachment tạm vt_body.htm (nếu còn)
        try:
            for i in range(int(mail_item.Attachments.Count), 0, -1):
                att = mail_item.Attachments.Item(i)
                name = str(getattr(att, "FileName", "") or "")
                if name.lower() in {"vt_body.htm", "vt_body.html"}:
                    att.Delete()
        except Exception:  # noqa: BLE001
            pass

        try:
            word_doc = mail_item.GetInspector.WordEditor
            if word_doc is not None:
                self._windows_trim_gap_before_signature(word_doc)
        except Exception:  # noqa: BLE001
            pass

    def _send_windows(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Path],
    ) -> None:
        import win32com.client  # type: ignore

        from .outlook_html import prepare_body_html_for_outlook

        outlook = win32com.client.Dispatch("Outlook.Application")
        account = self._windows_pick_account(outlook)
        prepared = prepare_body_html_for_outlook(body_html)

        mail_item = outlook.CreateItem(0)
        if account is not None:
            try:
                mail_item.SendUsingAccount = account
            except Exception:  # noqa: BLE001
                pass

        try:
            insp = mail_item.GetInspector
            try:
                _ = insp.WordEditor
            except Exception:  # noqa: BLE001
                pass
        except Exception:  # noqa: BLE001
            pass
        try:
            mail_item.Display(False)
            time.sleep(0.45)
        except Exception:  # noqa: BLE001
            pass

        self._windows_apply_body_keep_signature(mail_item, prepared)

        mail_item.To = "; ".join(to_list)
        if cc_list:
            mail_item.CC = "; ".join(cc_list)
        mail_item.Subject = subject
        if attachment:
            mail_item.Attachments.Add(str(attachment.resolve()))
        mail_item.Send()


SmtpSender = OutlookDesktopSender
OutlookSender = OutlookDesktopSender
