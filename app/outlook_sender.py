"""Tương thích import cũ — logic gửi nằm ở app.send (Win/Mac tách riêng)."""

from .send import OutlookDesktopSender, OutlookSender, SmtpSender

__all__ = ["OutlookDesktopSender", "OutlookSender", "SmtpSender"]
