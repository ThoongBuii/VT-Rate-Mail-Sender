"""Windows Outlook: chuyển data:image trong HTML → attachment CID (ảnh mới hiện khi gửi)."""

from __future__ import annotations

import base64
import re
import tempfile
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any

# PR_ATTACH_CONTENT_ID / PR_ATTACHMENT_HIDDEN (MAPI)
_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
_PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

_DATA_IMG_RE = re.compile(
    r"""(?is)(src\s*=\s*["'])(data:image/([a-z0-9.+-]+);base64,([A-Za-z0-9+/=\s]+))(["'])"""
)


@dataclass
class InlineCidImage:
    path: Path
    cid: str


def _ext_for_mime(subtype: str) -> str:
    sub = (subtype or "png").split("+", 1)[0].lower().strip()
    if sub in {"jpeg", "jpg"}:
        return "jpg"
    if sub in {"png", "gif", "bmp", "webp"}:
        return sub
    return "png"


def replace_data_images_with_cid(html: str) -> tuple[str, list[InlineCidImage]]:
    """Thay data:image bằng cid:…; ghi file tạm (chưa Attach vào Outlook)."""
    if not html or "data:image/" not in html.lower():
        return html or "", []

    items: list[InlineCidImage] = []
    tmp_dir = Path(tempfile.mkdtemp(prefix="vt_inline_"))

    def _repl(m: re.Match[str]) -> str:
        prefix, _full, subtype, b64, suffix = (
            m.group(1),
            m.group(2),
            m.group(3),
            m.group(4),
            m.group(5),
        )
        try:
            raw = base64.b64decode(re.sub(r"\s+", "", b64), validate=False)
        except Exception:  # noqa: BLE001
            return m.group(0)
        if not raw:
            return m.group(0)

        ext = _ext_for_mime(subtype)
        cid = f"vt.img.{uuid.uuid4().hex[:16]}@vt.local"
        path = tmp_dir / f"{cid.split('@', 1)[0]}.{ext}"
        try:
            path.write_bytes(raw)
        except Exception:  # noqa: BLE001
            return m.group(0)

        items.append(InlineCidImage(path=path, cid=cid))
        return f"{prefix}cid:{cid}{suffix}"

    return _DATA_IMG_RE.sub(_repl, html), items


def attach_cid_images(mail_item: Any, items: list[InlineCidImage]) -> None:
    """Đính ảnh ẩn sau khi đã set HTMLBody (cid khớp Content-ID)."""
    for item in items:
        try:
            att = mail_item.Attachments.Add(str(item.path), 1)  # olByValue
            pa = att.PropertyAccessor
            # Outlook thường cần Content-ID dạng <cid>
            pa.SetProperty(_PR_ATTACH_CONTENT_ID, f"<{item.cid}>")
            try:
                pa.SetProperty(_PR_ATTACHMENT_HIDDEN, True)
            except Exception:  # noqa: BLE001
                pass
        except Exception:  # noqa: BLE001
            continue


def cleanup_temp_files(items: list[InlineCidImage]) -> None:
    paths = [i.path for i in items]
    for p in paths:
        try:
            p.unlink(missing_ok=True)
        except Exception:  # noqa: BLE001
            pass
    parents = {p.parent for p in paths}
    for d in parents:
        try:
            if d.is_dir() and not any(d.iterdir()):
                d.rmdir()
        except Exception:  # noqa: BLE001
            pass
