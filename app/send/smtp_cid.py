"""Gửi mail SMTP với ảnh nhúng CID (multipart/related) — hết banner external images."""

from __future__ import annotations

import mimetypes
import re
import smtplib
import tempfile
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr, make_msgid
from email import encoders
from pathlib import Path
from typing import Optional, Sequence

from ..models import AppConfig
from ..smtp_creds import get_smtp_password
from .html_merge import html_visible_text
from .mac_inline import prepare_mac_send_html


def _html_to_plain(html: str) -> str:
    s = re.sub(r"(?is)<br\s*/?>", "\n", html or "")
    s = re.sub(r"(?is)</p\s*>", "\n\n", s)
    s = re.sub(r"(?is)<[^>]+>", "", s)
    s = (
        s.replace("&nbsp;", " ")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
    )
    return re.sub(r"\n{3,}", "\n\n", s).strip()


def build_mime_message(
    *,
    config: AppConfig,
    to_list: list[str],
    cc_list: list[str],
    subject: str,
    html_body: str,
    attachments: Optional[Sequence[Path]] = None,
    bcc_self: bool = True,
) -> MIMEMultipart:
    """
    mixed
      └ related
           ├ alternative (plain + html với cid:)
           └ image/* inline (Content-ID)
      └ file attachments (nếu có)
    """
    with tempfile.TemporaryDirectory() as tmp:
        send_html, cid_parts = prepare_mac_send_html(html_body, Path(tmp) / "cid")
        if re.search(r"""(?is)<img\b[^>]*?\bsrc\s*=\s*['"]https?://""", send_html):
            raise RuntimeError(
                "Còn ảnh http ngoài trong HTML — không gửi (tránh banner privacy)."
            )
        if re.search(r"""(?is)<img\b[^>]*?\bsrc\s*=\s*['"]data:""", send_html):
            raise RuntimeError(
                "Còn data:image chưa đổi sang CID — hủy gửi."
            )
        if len(html_visible_text(send_html)) < 40:
            raise RuntimeError("Nội dung mail quá ngắn sau chuẩn bị CID.")

        root = MIMEMultipart("mixed")
        root["Subject"] = Header(subject or "(no subject)", "utf-8")
        root["From"] = formataddr((config.from_name or "", config.from_email))
        root["To"] = ", ".join(to_list)
        if cc_list:
            root["Cc"] = ", ".join(cc_list)
        root["Message-ID"] = make_msgid(domain="vtlogisticsvn.com")
        # Giảm spam heuristic nhẹ
        root["X-Mailer"] = "VT Rate Mail Sender"

        related = MIMEMultipart("related")
        alt = MIMEMultipart("alternative")
        plain = _html_to_plain(send_html)
        alt.attach(MIMEText(plain or " ", "plain", "utf-8"))
        alt.attach(MIMEText(send_html, "html", "utf-8"))
        related.attach(alt)

        for cid, img_path in cid_parts:
            raw = img_path.read_bytes()
            subtype = (img_path.suffix or ".png").lstrip(".").lower()
            if subtype == "jpg":
                subtype = "jpeg"
            img = MIMEImage(raw, _subtype=subtype)
            img.add_header("Content-ID", f"<{cid}>")
            img.add_header(
                "Content-Disposition",
                "inline",
                filename=img_path.name,
            )
            related.attach(img)

        root.attach(related)

        for path in attachments or []:
            if path is None or not Path(path).is_file():
                continue
            p = Path(path)
            ctype, _ = mimetypes.guess_type(str(p))
            if not ctype:
                ctype = "application/octet-stream"
            main, sub = ctype.split("/", 1)
            part = MIMEBase(main, sub)
            part.set_payload(p.read_bytes())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                "attachment",
                filename=p.name,
            )
            root.attach(part)

        # Không gắn header Bcc vào MIME (tránh lộ). Người nhận BCC chỉ nằm trong
        # envelope recipients lúc sendmail().
        del bcc_self

        return root


def test_smtp_login(config: AppConfig, *, timeout: float = 25.0) -> str:
    password = get_smtp_password()
    if not password:
        raise RuntimeError(
            "Chưa có mật khẩu SMTP.\n"
            "Bấm «Chữ ký» → nhập mật khẩu mail (overseas@…) → Lưu."
        )
    user = (config.smtp_username or config.from_email or "").strip()
    host = (config.smtp_host or "").strip()
    port = int(config.smtp_port or 465)
    if not host or not user:
        raise RuntimeError("Thiếu smtp_host / smtp_username trong config.")

    if config.smtp_ssl:
        with smtplib.SMTP_SSL(host, port, timeout=timeout) as smtp:
            smtp.login(user, password)
    else:
        with smtplib.SMTP(host, port, timeout=timeout) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(user, password)
    return f"{user} @ {host}:{port}"


def send_via_smtp(
    config: AppConfig,
    *,
    to_list: list[str],
    cc_list: list[str],
    subject: str,
    html_body: str,
    attachments: Optional[Sequence[Path]] = None,
    timeout: float = 90.0,
) -> None:
    if not to_list:
        raise ValueError("Thiếu địa chỉ To.")
    password = get_smtp_password()
    if not password:
        raise RuntimeError(
            "Chưa có mật khẩu SMTP.\n"
            "Bấm «Chữ ký» → nhập mật khẩu mail → Lưu, rồi Semi-Auto."
        )

    msg = build_mime_message(
        config=config,
        to_list=to_list,
        cc_list=cc_list or [],
        subject=subject,
        html_body=html_body,
        attachments=attachments,
        bcc_self=True,
    )

    user = (config.smtp_username or config.from_email or "").strip()
    host = (config.smtp_host or "").strip()
    port = int(config.smtp_port or 465)
    recipients = list(to_list) + list(cc_list or [])
    if config.from_email and config.from_email.lower() not in {
        r.lower() for r in recipients
    }:
        recipients.append(config.from_email)

    try:
        if config.smtp_ssl:
            with smtplib.SMTP_SSL(host, port, timeout=timeout) as smtp:
                smtp.login(user, password)
                smtp.sendmail(config.from_email, recipients, msg.as_string())
        else:
            with smtplib.SMTP(host, port, timeout=timeout) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.ehlo()
                smtp.login(user, password)
                smtp.sendmail(config.from_email, recipients, msg.as_string())
    except smtplib.SMTPAuthenticationError as exc:
        raise RuntimeError(
            "SMTP đăng nhập thất bại (sai mật khẩu / user).\n"
            f"User: {user} · Host: {host}:{port}\n"
            f"{exc}"
        ) from exc
    except smtplib.SMTPException as exc:
        raise RuntimeError(f"SMTP gửi lỗi: {exc}") from exc
