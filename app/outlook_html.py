"""Chuẩn hóa HTML body app → gần với Word/Outlook New Mail (font, bảng, px→pt)."""

from __future__ import annotations

import re

DEFAULT_FONT_STACK = "Aptos, Calibri, Arial, sans-serif"
DEFAULT_SIZE_PT = "12pt"
TABLE_CELL_SIZE_PT = "10pt"


def _has_css(style: str, prop: str) -> bool:
    return bool(re.search(rf"(?i)(?:^|;)\s*{re.escape(prop)}\s*:", style or ""))


def _merge_style(existing: str, additions: str) -> str:
    parts: list[str] = []
    seen: set[str] = set()
    for chunk in f"{existing or ''};{additions or ''}".split(";"):
        chunk = chunk.strip()
        if not chunk or ":" not in chunk:
            continue
        key = chunk.split(":", 1)[0].strip().lower()
        if key in seen:
            continue
        seen.add(key)
        parts.append(chunk)
    return ";".join(parts)


def convert_px_font_sizes_to_pt(html: str) -> str:
    def _repl(m: re.Match[str]) -> str:
        px = float(m.group(1))
        pt = round(px * 72 / 96, 1)
        if pt == int(pt):
            pt_s = str(int(pt))
        else:
            pt_s = str(pt)
        return f"font-size:{pt_s}pt"

    return re.sub(r"font-size\s*:\s*(\d+(?:\.\d+)?)px", _repl, html or "", flags=re.I)


def convert_px_widths_to_pt(html: str) -> str:
    """
    Đổi width/height px → pt (96dpi cố định).

    Laptop Windows hay scale 125%/150%: CF_HTML / contenteditable mang px đã
    «phóng» theo DPI → Word InsertFile hiểu lệch → bảng nhảy cột so với desktop 100%.
    Chuẩn hóa về pt (đơn vị Word) giúp đồng bộ giữa máy.
    """

    def _style_repl(m: re.Match[str]) -> str:
        prop, num = m.group(1), float(m.group(2))
        pt = round(num * 72 / 96, 1)
        pt_s = str(int(pt)) if pt == int(pt) else str(pt)
        return f"{prop}:{pt_s}pt"

    return re.sub(
        r"(?i)(?<![-\w])(width|height|min-width|max-width|min-height|max-height)"
        r"\s*:\s*(\d+(?:\.\d+)?)px",
        _style_repl,
        html or "",
    )


def harden_tables_for_outlook(html: str) -> str:
    """Giữ layout bảng khi Outlook Word render lại HTMLBody — không ép font."""

    def _table_repl(m: re.Match[str]) -> str:
        tag = m.group(0)
        style = ""
        sm = re.search(r"""\sstyle\s*=\s*["']([^"']*)["']""", tag, flags=re.I)
        if sm:
            style = sm.group(1)
        # Bỏ max-width:% / width:100% — Word hay co bảng khác nhau theo DPI cửa sổ
        style = re.sub(r"(?i)(?:^|;)\s*max-width\s*:\s*100%\s*;?", ";", style or "")
        style = re.sub(r"(?i)(?:^|;)\s*width\s*:\s*100%\s*;?", ";", style)
        style = re.sub(r";{2,}", ";", style).strip("; ").strip()
        add = (
            "border-collapse:collapse;"
            "border-spacing:0;"
            "mso-table-lspace:0pt;"
            "mso-table-rspace:0pt;"
        )
        new_style = _merge_style(style, add)

        if sm:
            tag = tag[: sm.start()] + f' style="{new_style}"' + tag[sm.end() :]
        else:
            tag = re.sub(r"^<table\b", f'<table style="{new_style}"', tag, count=1, flags=re.I)

        if not re.search(r"\sborder\s*=", tag, flags=re.I):
            tag = re.sub(r"^<table\b", '<table border="1"', tag, count=1, flags=re.I)
        if not re.search(r"\scellspacing\s*=", tag, flags=re.I):
            tag = re.sub(r"^<table\b", '<table cellspacing="0"', tag, count=1, flags=re.I)
        if not re.search(r"\scellpadding\s*=", tag, flags=re.I):
            tag = re.sub(r"^<table\b", '<table cellpadding="4"', tag, count=1, flags=re.I)
        return tag

    html = re.sub(r"<table\b[^>]*>", _table_repl, html or "", flags=re.I)

    def _cell_repl(m: re.Match[str]) -> str:
        full = m.group(0)
        style = ""
        sm = re.search(r"""\sstyle\s*=\s*["']([^"']*)["']""", full, flags=re.I)
        if sm:
            style = sm.group(1)
        if _has_css(style, "vertical-align"):
            return full
        new_style = _merge_style(style, "vertical-align:middle")
        if sm:
            return full[: sm.start()] + f' style="{new_style}"' + full[sm.end() :]
        return re.sub(
            r"^<(td|th)\b",
            rf'<\1 style="{new_style}"',
            full,
            count=1,
            flags=re.I,
        )

    return re.sub(r"<(?:td|th)\b[^>]*>", _cell_repl, html, flags=re.I)


def ensure_inline_fonts(html: str, *, in_table: bool = False) -> str:
    """Không ép font cố định — font do user chọn trên thanh công cụ."""
    return html or ""


def wrap_root_body(html: str) -> str:
    s = (html or "").strip()
    if not s:
        return s
    if re.search(
        r"""(?is)^\s*<div\b[^>]*\bdata-vt-body\s*=\s*["']1["']""",
        s,
    ):
        return s
    return (
        f'<div data-vt-body="1" style="line-height:1.35;margin:0;padding:0;color:#222;">'
        f"{s}</div>"
    )


def strip_trailing_empty_blocks(html: str) -> str:
    """Bỏ p/div trống / <br> cuối (thường từ contenteditable)."""
    s = html or ""
    leaf_empty = (
        r"<(?:p|div)(?:\s[^>]*)?>\s*(?:"
        r"<br\s*/?>|&nbsp;|\xa0|&#160;|\s|"
        r"<span[^>]*>\s*(?:&nbsp;|\xa0|&#160;|<br\s*/?>|\s|"
        r"<o:p[^>]*>\s*(?:&nbsp;|\xa0|&#160;)?\s*</o:p>)*</span>|"
        r"<o:p[^>]*>\s*(?:&nbsp;|\xa0|&#160;)?\s*</o:p>"
        r")*\s*</(?:p|div)>"
    )
    leaf_re = re.compile(rf"(?:{leaf_empty})((?:\s*</div>)*)\s*$", flags=re.I)
    br_re = re.compile(r"(?:<br\s*/?>\s*|&nbsp;|\xa0|&#160;|\s)+$", flags=re.I)
    changed = True
    while changed:
        changed = False
        m = leaf_re.search(s)
        if m:
            s = s[: m.start()] + (m.group(1) or "")
            changed = True
            continue
        m = br_re.search(s)
        if m:
            s = s[: m.start()]
            changed = True
            continue
    return s


def harden_images_for_outlook(html: str) -> str:
    """Bỏ max-width:100% trên img — Outlook hay thu nhỏ ảnh vì CSS này."""

    def _img_repl(m: re.Match[str]) -> str:
        open_tag, attrs, close = m.group(1), m.group(2), m.group(3)
        style_m = re.search(r"""(?is)\sstyle\s*=\s*["']([^"']*)["']""", attrs)
        if style_m:
            style = style_m.group(1)
            style = re.sub(r"(?i)(?:^|;)\s*max-width\s*:\s*100%\s*;?", ";", style)
            style = re.sub(r";{2,}", ";", style).strip("; ").strip()
            if style:
                attrs = attrs[: style_m.start()] + f' style="{style}"' + attrs[style_m.end() :]
            else:
                attrs = attrs[: style_m.start()] + attrs[style_m.end() :]
        return f"{open_tag}{attrs}{close}"

    return re.sub(r"(?is)(<img\b)([^>]*?)(/?>)", _img_repl, html or "")


def prepare_body_html_for_outlook(html: str) -> str:
    """Pipeline trước khi chèn vào New Mail Outlook (đồng bộ desktop/laptop DPI)."""
    s = (html or "").strip()
    if not s:
        return s
    s = convert_px_font_sizes_to_pt(s)
    s = convert_px_widths_to_pt(s)
    s = harden_tables_for_outlook(s)
    s = harden_images_for_outlook(s)
    s = ensure_inline_fonts(s)
    s = strip_trailing_empty_blocks(s)
    s = wrap_root_body(s)
    return s
