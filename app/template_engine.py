from __future__ import annotations

import html
import re
from pathlib import Path

from .models import AgencyMail

PLACEHOLDER_RE = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}")

VAR_MAP = {
    "agency_company": "agency_company",
    "agency": "agency_company",
    "company": "agency_company",
    "account_name": "account_name",
    "name": "account_name",
    "account_mail": "account_mail",
    "email": "account_mail",
    "mail_cc": "mail_cc",
    "cc": "mail_cc",
    "subject": "subject",
}

SUGGESTED_SUBJECT = "VT Rate Quotation – {{agency_company}}"

# HTML gợi ý — user sẽ dán bảng giá + chữ ký bằng Outlook New Mail
SUGGESTED_TEMPLATE_HTML = """
<div style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#222;line-height:1.45;">
  <p>Dear {{account_name}},</p>
  <p>Good day!</p>
  <p>Please find our latest VT Rate Quotation for <b>{{agency_company}}</b>.</p>
  <p style="color:#c0392b;"><b><i>*** Remark: Due to current market changes, Carrier may issue unforeseen notices. Rate &amp; space are subject to availability at the time of booking. ***</i></b></p>
  <p>Thank you for your kind support.</p>
  <p>Best regards,</p>
</div>
""".strip()

# Backward alias used by models
SUGGESTED_TEMPLATE = SUGGESTED_TEMPLATE_HTML


def is_html_content(text: str) -> bool:
    if not text:
        return False
    return bool(re.search(r"<(table|div|p|br|img|span|html|body|font|b|i|u|strong)\b", text, re.I))


def _values_for(mail: AgencyMail) -> dict[str, str]:
    return {
        "agency_company": mail.agency_company,
        "account_name": mail.account_name or "Sir/Madam",
        "account_mail": mail.account_mail,
        "mail_cc": mail.mail_cc,
        "subject": mail.subject,
    }


def render_text(template: str, mail: AgencyMail, *, escape_html_values: bool = False) -> str:
    values = _values_for(mail)

    def repl(match: re.Match[str]) -> str:
        key = match.group(1).lower()
        field = VAR_MAP.get(key, key)
        val = values.get(field, match.group(0))
        if escape_html_values:
            val = html.escape(val)
        return val

    return PLACEHOLDER_RE.sub(repl, template or "")


def text_to_html(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if is_html_content(text):
        return text
    escaped = html.escape(text)
    paragraphs = escaped.split("\n\n")
    parts: list[str] = []
    for para in paragraphs:
        lines = "<br>\n".join(para.split("\n"))
        parts.append(
            f"<p style='margin:0 0 12px 0;font-family:Calibri,Arial,sans-serif;"
            f"font-size:14px;color:#222;line-height:1.45;'>{lines}</p>"
        )
    return "<div>" + "".join(parts) + "</div>"


def render_subject(mail: AgencyMail) -> str:
    return render_text(mail.subject, mail).strip()


def render_body_html(mail: AgencyMail, signature_html: str = "") -> str:
    """Render body. Template có thể là HTML đầy đủ (bảng + chữ ký đã dán)."""
    raw = mail.template_mail or ""
    if is_html_content(raw):
        body = render_text(raw, mail, escape_html_values=True)
    else:
        body = text_to_html(render_text(raw, mail))
    if signature_html.strip():
        body = body + "<br>" + signature_html
    return body


def load_template_file(path: str | Path) -> str:
    return Path(path).read_text(encoding="utf-8")


def wrap_preview_document(inner_html: str) -> str:
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  body {{ margin: 12px; background: #fff; color: #222;
         font-family: Calibri, Arial, sans-serif; font-size: 14px; }}
  table {{ border-collapse: collapse; }}
  img {{ max-width: 100%; height: auto; }}
</style></head><body>{inner_html}</body></html>"""
