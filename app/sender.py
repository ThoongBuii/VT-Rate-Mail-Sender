"""Sender facade — ưu tiên Outlook desktop (không nhập mật khẩu)."""

from .outlook_sender import OutlookDesktopSender, OutlookSender, SmtpSender

__all__ = ["OutlookDesktopSender", "OutlookSender", "SmtpSender"]
