"""Outlook send facade — tách Windows / macOS."""

from __future__ import annotations

import platform
from pathlib import Path
from typing import Any, Optional

from ..models import AgencyMail, AppConfig
from ..sender_smtp import split_emails
from ..template_engine import render_body_html, render_subject
from .mac import MacOutlookSender
from .windows import WindowsOutlookSender


class OutlookDesktopSender:
    """
    Facade: App soạn Dear/bảng/remark.
    - Windows: COM + merge chữ ký New Mail.
    - macOS: chữ ký đã chụp 1 lần + set content (không Tab/Cmd+V).
    """

    def __init__(self, config: AppConfig):
        self.config = config
        self._ready = False
        self._account_label = ""
        self._mac = MacOutlookSender(config)
        self._win = WindowsOutlookSender(config)

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
            label = self._mac.open_outlook()
            self._ready = True
            self._account_label = label
            return f"Outlook đã mở · {label} · Legacy ON · New Mail + dán body (giữ ảnh chữ ký)"
        if system == "Windows":
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
        raise RuntimeError(f"Hệ điều hành chưa hỗ trợ: {system}")

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
        try:
            from ..paths import user_data_dir

            root = user_data_dir()
        except Exception:  # noqa: BLE001
            root = Path(__file__).resolve().parent.parent.parent
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

        note = "Chữ ký mặc định Outlook sẽ tự gắn khi gửi (giống New Mail)."
        mac_ready = True
        if platform.system() == "Darwin":
            st = self._mac.signature_status()
            note = st["message"]
            mac_ready = True

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
            "signature_note": note,
            "mac_signature_ready": mac_ready,
        }

    def capture_mac_signature(self) -> str:
        if platform.system() != "Darwin":
            raise RuntimeError("Chụp chữ ký chỉ dành cho macOS.")
        msg = self._mac.capture_signature()
        self._ready = True
        return msg

    def mac_signature_status(self) -> dict[str, Any]:
        if platform.system() != "Darwin":
            return {
                "ready": True,
                "message": "Windows không cần chụp chữ ký riêng.",
                "path": "",
            }
        return self._mac.signature_status()

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
            self._mac.send(to_list, cc_list, subject, body_html, attachment)
        elif system == "Windows":
            self._win.send(to_list, cc_list, subject, body_html, attachment)
        else:
            raise RuntimeError(f"Hệ điều hành chưa hỗ trợ: {system}")


SmtpSender = OutlookDesktopSender
OutlookSender = OutlookDesktopSender
