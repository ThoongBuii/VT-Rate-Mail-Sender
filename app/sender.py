"""Sender facade — Outlook desktop (không nhập mật khẩu)."""

from .send import OutlookDesktopSender, OutlookSender, SmtpSender

__all__ = ["OutlookDesktopSender", "OutlookSender", "SmtpSender"]
