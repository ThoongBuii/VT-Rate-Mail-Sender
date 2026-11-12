"""Giữ cứng nội dung khung Soạn trong một lượt gửi.

Lấy một lần lúc bắt đầu list. Chỉ thả khi list gửi xong,
hoặc khi Pause rồi sửa nội dung thư.
"""

from __future__ import annotations

_held_template: str | None = None
_clips: dict[str, list[tuple[str, bytes]]] = {}


def hold(template_html: str) -> None:
    """Ghim nội dung app cho cả list. Không đọc lại khung Soạn mỗi mail."""
    global _held_template
    _held_template = template_html or ""


def held_template() -> str | None:
    return _held_template


def release() -> None:
    """Bỏ bản đang giữ và mọi bản copy đã lưu theo nội dung đó."""
    global _held_template
    _held_template = None
    _clips.clear()


def replace_if_edited(template_html: str) -> bool:
    """Pause rồi sửa thư: thả bản cũ. Không sửa thì giữ nguyên."""
    new = template_html or ""
    if _held_template is None:
        hold(new)
        return False
    if new == _held_template:
        return False
    release()
    hold(new)
    return True


def cached_clip(key: str) -> list[tuple[str, bytes]] | None:
    return _clips.get(key)


def store_clip(key: str, clip: list[tuple[str, bytes]]) -> None:
    _clips[key] = clip
