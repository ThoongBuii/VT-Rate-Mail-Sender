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
        if len(to_list) != 1:
            raise ValueError(
                f"Mỗi lần chỉ gửi 1 Account Mail. Hiện có {len(to_list)} địa chỉ TO."
            )

        subject = render_subject(mail)
        body_html = render_body_html(mail, "")
        cc_list = split_emails(mail.mail_cc)
        attachment = self.resolve_attachment(mail)

        system = platform.system()
        if system == "Darwin":
            self._send_mac(to_list[0], cc_list, subject, body_html, attachment)
        elif system == "Windows":
            self._send_windows(to_list[0], cc_list, subject, body_html, attachment)
        else:
            raise RuntimeError(f"Hệ điều hành chưa hỗ trợ: {system}")

    def _send_mac(
        self,
        to_addr: str,
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Path],
    ) -> None:
        """
        Mac Classic Outlook: tạo outgoing message (Outlook có thể gắn chữ ký),
        rồi chèn nội dung app phía trên — không thay thế toàn bộ content bằng body thuần.
        """
        with tempfile.TemporaryDirectory() as tmp:
            html_path = Path(tmp) / "body.html"
            html_path.write_text(body_html, encoding="utf-8")

            def esc(s: str) -> str:
                return (s or "").replace("\\", "\\\\").replace('"', '\\"')

            cc_block = "\n".join(
                f'make new cc recipient at msg with properties {{email address:{{address:"{esc(cc)}"}}}}'
                for cc in cc_list
            )
            if attachment:
                att = esc(str(attachment.resolve()))
                att_block = f'make new attachment at msg with properties {{file:POSIX file "{att}"}}'
            else:
                att_block = ""

            script = f'''
set htmlPath to "{html_path}"
set htmlText to do shell script "cat " & quoted form of htmlPath
tell application "Microsoft Outlook"
  activate
  set msg to make new outgoing message
  delay 0.4
  set subject of msg to "{esc(subject)}"
  try
    set existingContent to content of msg
    set content of msg to htmlText & return & return & existingContent
  on error
    try
      set content of msg to htmlText
    on error
      set plain text content of msg to htmlText
    end try
  end try
  make new to recipient at msg with properties {{email address:{{address:"{esc(to_addr)}"}}}}
  {cc_block}
  {att_block}
  send msg
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
                "Outlook Mac gửi thất bại.\n" + (result.stderr or result.stdout or "")
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
    def _merge_body_with_outlook_signature(body_html: str, signature_html: str) -> str:
        """Chèn nội dung app vào đầu New Mail đã có chữ ký Outlook (cùng MailItem)."""
        import re

        body = (body_html or "").strip()
        sig = (signature_html or "").strip()
        if not body:
            return sig
        if not sig:
            return body

        # Outlook hay trả HTML đầy đủ — chèn sau <body ...>
        m = re.search(r"<body[^>]*>", sig, flags=re.I)
        if m:
            i = m.end()
            return sig[:i] + body + "<br><br>" + sig[i:]

        return body + "<br><br>" + sig

    def _send_windows(
        self,
        to_addr: str,
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Path],
    ) -> None:
        import win32com.client  # type: ignore

        outlook = win32com.client.Dispatch("Outlook.Application")
        account = self._windows_pick_account(outlook)

        # Một MailItem: Outlook gắn chữ ký → chèn body → gửi (giữ ảnh cid chữ ký)
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
            time.sleep(0.35)
        except Exception:  # noqa: BLE001
            pass

        existing = str(getattr(mail_item, "HTMLBody", None) or "")
        mail_item.HTMLBody = self._merge_body_with_outlook_signature(body_html, existing)

        mail_item.To = to_addr
        if cc_list:
            mail_item.CC = "; ".join(cc_list)
        mail_item.Subject = subject
        if attachment:
            mail_item.Attachments.Add(str(attachment.resolve()))
        mail_item.Send()


SmtpSender = OutlookDesktopSender
OutlookSender = OutlookDesktopSender
