from __future__ import annotations

import re
from typing import Optional


# email thô trong chuỗi (kể cả "Name <a@b.com>")
_EMAIL_RE = re.compile(
    r"[a-zA-Z0-9](?:[a-zA-Z0-9._%+\-]*[a-zA-Z0-9])?@"
    r"[a-zA-Z0-9](?:[a-zA-Z0-9.\-]*[a-zA-Z0-9])?\.[a-zA-Z]{2,}",
)


def is_valid_email(addr: str) -> bool:
    a = (addr or "").strip()
    if not a or "@" not in a or " " in a or len(a) > 120:
        return False
    if a.lower().startswith("mailto:"):
        return False
    if "microsoftinternetexplorer" in a.lower():
        return False
    local, _, domain = a.rpartition("@")
    if not local or not domain or "." not in domain:
        return False
    if domain.startswith(".") or domain.endswith("."):
        return False
    return bool(_EMAIL_RE.fullmatch(a))


def extract_email(raw: str) -> str:
    """Lấy một địa chỉ email sạch từ ô Excel / 'Name <mail@x.com>'."""
    s = (raw or "").strip()
    if not s:
        return ""
    # Name <email@domain>
    m = re.search(r"<([^>]+)>", s)
    if m:
        cand = m.group(1).strip().strip('"').strip("'")
        if is_valid_email(cand):
            return cand
    s = s.strip().strip('"').strip("'")
    if is_valid_email(s):
        return s
    found = _EMAIL_RE.findall(s)
    for cand in found:
        if is_valid_email(cand):
            return cand
    return ""


def split_emails(raw: str) -> list[str]:
    """
    Tách To/Cc: ; hoặc , · bỏ trùng · chỉ giữ email hợp lệ.
    Ví dụ: 'a@x.com; Name <b@y.com>, c@z.com'
    """
    text = (raw or "").strip()
    if not text:
        return []
    parts = re.split(r"[;,]+", text)
    out: list[str] = []
    seen: set[str] = set()
    for part in parts:
        em = extract_email(part)
        if not em:
            # Phần không tách được nhưng cả chuỗi là một email
            continue
        key = em.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(em)
    # Fallback: cả ô chỉ có một email không có ;,
    if not out:
        em = extract_email(text)
        if em:
            out.append(em)
    return out


def normalize_to_cc(
    to_raw: str, cc_raw: str = ""
) -> tuple[list[str], list[str]]:
    """
    To ưu tiên; địa chỉ trùng giữa To và Cc chỉ giữ ở To.
    """
    to_list = split_emails(to_raw)
    cc_list = split_emails(cc_raw)
    to_keys = {e.lower() for e in to_list}
    cc_clean = [e for e in cc_list if e.lower() not in to_keys]
    return to_list, cc_clean
