"""Chuỗi attachment nhiều file — ngăn cách ||| (an toàn với đường dẫn Windows)."""

from __future__ import annotations

ATTACH_SEP = "|||"


def split_attachments(raw: str | None) -> list[str]:
    text = (raw or "").strip()
    if not text:
        return []
    if ATTACH_SEP in text:
        parts = text.split(ATTACH_SEP)
    elif "\n" in text:
        parts = text.splitlines()
    else:
        parts = [text]
    return [p.strip() for p in parts if p.strip()]


def join_attachments(paths: list[str] | None) -> str:
    cleaned = [p.strip() for p in (paths or []) if p and str(p).strip()]
    return ATTACH_SEP.join(cleaned)
