from __future__ import annotations

from dataclasses import dataclass, field, asdict
from enum import Enum
from typing import Any, Optional
import time


class MailStatus(str, Enum):
    PENDING = "pending"
    READY = "ready"
    SENDING = "sending"
    SENT = "sent"
    FAILED = "failed"
    SKIPPED = "skipped"
    PAUSED = "paused"


@dataclass
class AgencyMail:
    agency_company: str = ""
    account_name: str = ""
    account_mail: str = ""
    mail_cc: str = ""
    subject: str = ""
    attachment: str = ""
    template_mail: str = ""
    status: MailStatus = MailStatus.PENDING
    error: str = ""
    sent_at: Optional[float] = None
    row_index: int = 0

    def validate(self) -> list[str]:
        errors: list[str] = []
        if not self.account_mail.strip():
            errors.append("Thiếu Account Mail")
        elif "@" not in self.account_mail:
            errors.append("Account Mail không hợp lệ")
        if not self.subject.strip():
            errors.append("Thiếu Subject")
        if not self.template_mail.strip():
            errors.append("Thiếu nội dung Template Mail")
        return errors

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["status"] = self.status.value
        return data


@dataclass
class SendLogEntry:
    timestamp: float
    row_index: int
    agency_company: str
    account_mail: str
    subject: str
    status: str
    message: str = ""

    def display_time(self) -> str:
        return time.strftime("%H:%M:%S", time.localtime(self.timestamp))


@dataclass
class AppConfig:
    """Config app — Subject/Template/Attachment lưu để dùng lại."""

    smtp_host: str = "mail.vtlogisticsvn.com"
    smtp_port: int = 465
    smtp_ssl: bool = True
    from_name: str = "Amber- VT Logistics"
    from_email: str = "overseas@vtlogisticsvn.com"
    smtp_username: str = "overseas@vtlogisticsvn.com"
    delay_min_seconds: int = 10
    delay_max_seconds: int = 20
    default_attachment_dir: str = ""
    signature_html: str = ""
    # Compose trên UI (HTML — có thể gồm bảng + chữ ký đã dán từ Outlook)
    default_subject: str = "VT Rate Quotation – {{agency_company}}"
    default_attachment: str = ""
    default_template: str = ""

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "AppConfig":
        from .template_engine import SUGGESTED_TEMPLATE, SUGGESTED_SUBJECT

        return cls(
            smtp_host=str(data.get("smtp_host", "mail.vtlogisticsvn.com")),
            smtp_port=int(data.get("smtp_port", 465)),
            smtp_ssl=bool(data.get("smtp_ssl", True)),
            from_name=str(data.get("from_name", "Amber- VT Logistics")),
            from_email=str(data.get("from_email", "overseas@vtlogisticsvn.com")),
            smtp_username=str(
                data.get("smtp_username")
                or data.get("from_email")
                or "overseas@vtlogisticsvn.com"
            ),
            delay_min_seconds=int(data.get("delay_min_seconds", 10)),
            delay_max_seconds=int(data.get("delay_max_seconds", 20)),
            default_attachment_dir=str(data.get("default_attachment_dir", "")),
            signature_html=str(data.get("signature_html", "")),
            default_subject=str(data.get("default_subject") or SUGGESTED_SUBJECT),
            default_attachment=str(data.get("default_attachment", "")),
            default_template=str(data.get("default_template") or SUGGESTED_TEMPLATE),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "smtp_host": self.smtp_host,
            "smtp_port": self.smtp_port,
            "smtp_ssl": self.smtp_ssl,
            "from_name": self.from_name,
            "from_email": self.from_email,
            "smtp_username": self.smtp_username,
            "delay_min_seconds": self.delay_min_seconds,
            "delay_max_seconds": self.delay_max_seconds,
            "default_attachment_dir": self.default_attachment_dir,
            "signature_html": self.signature_html,
            "default_subject": self.default_subject,
            "default_attachment": self.default_attachment,
            "default_template": self.default_template,
        }


@dataclass
class SendProgress:
    total: int = 0
    sent: int = 0
    failed: int = 0
    skipped: int = 0
    current_index: int = -1
    current_agency: str = ""
    is_running: bool = False
    is_paused: bool = False
    wait_remaining: float = 0.0
    logs: list[SendLogEntry] = field(default_factory=list)

    @property
    def done(self) -> int:
        return self.sent + self.failed + self.skipped

    @property
    def percent(self) -> float:
        if self.total <= 0:
            return 0.0
        return min(100.0, (self.done / self.total) * 100.0)
