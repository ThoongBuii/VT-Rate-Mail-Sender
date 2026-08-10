"""Merge body HTML with Outlook signature HTML (shared Win/Mac)."""
from __future__ import annotations

from typing import Any

def html_visible_text(inner: str) -> str:
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

def is_empty_html_block(inner: str) -> bool:
    return html_visible_text(inner) == ""

def inject_style_attr(open_tag: str, css: str) -> str:
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

def strip_edge_empty_html(html: str, *, leading: bool) -> str:
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
                return m.group(1) + strip_edge_empty_html(
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

def zero_first_block_margin(html: str) -> str:
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
            + zero_first_block_margin(s[m2.end() :])
        )

    m = re.match(r"^(<(?:p|div)(?:\s[^>]*)?)>", s, flags=re.I)
    if not m:
        return s
    open_tag = inject_style_attr(
        m.group(1) + ">",
        "margin-top:0;padding-top:0",
    )
    return open_tag + s[m.end() :]

def zero_last_block_margin(html: str) -> str:
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
        if is_empty_html_block(inner):
            continue
        # Phải là block lá cuối (sau closing chỉ còn đóng wrapper/whitespace)
        after = s[m.end() + close.end() :]
        if re.search(r"<(?:p|div|table|br)\b", after, flags=re.I):
            continue
        new_open = inject_style_attr(
            m.group(0), "margin-bottom:0;padding-bottom:0"
        )
        return s[: m.start()] + new_open + s[m.end() :]
    return s

def merge_body_with_outlook_signature(body_html: str, signature_html: str) -> str:
    """
    Chèn nội dung app sát chữ ký Outlook (0 khoảng mặc định).
    Khoảng trống chỉ còn nếu user tự Enter trong app.
    """
    import re

    body = strip_edge_empty_html(
        (body_html or "").strip(), leading=False
    )
    sig = (signature_html or "").strip()
    if not body:
        return sig
    if not sig:
        return body

    body = zero_last_block_margin(body)
    # Wrapper ngoài cùng của app cũng không tạo khoảng dưới
    import re as _re

    _mwrap = _re.match(r"^(<div(?:\s[^>]*)?)>", body, flags=_re.I)
    if _mwrap:
        body = (
            inject_style_attr(
                _mwrap.group(1) + ">",
                "margin-bottom:0;padding-bottom:0",
            )
            + body[_mwrap.end() :]
        )

    m = re.search(r"<body[^>]*>", sig, flags=re.I)
    if m:
        i = m.end()
        rest = strip_edge_empty_html(sig[i:], leading=True)
        # Chèn body vào trong WordSection (cùng vùng soạn) — sát chữ ký hơn
        wm = re.match(
            r"(<div[^>]*(?:WordSection|elementToProof|OutlookMessageBody)[^>]*>)\s*",
            rest,
            flags=re.I,
        )
        if wm:
            inner = strip_edge_empty_html(
                rest[wm.end() :], leading=True
            )
            inner = zero_first_block_margin(inner)
            return sig[:i] + wm.group(1) + body + inner

        rest = zero_first_block_margin(rest)
        return sig[:i] + body + rest

    rest = strip_edge_empty_html(sig, leading=True)
    rest = zero_first_block_margin(rest)
    return body + rest

