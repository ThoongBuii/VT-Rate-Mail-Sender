"""Windows Outlook: chuyển data:image trong HTML → attachment CID (ảnh mới hiện khi gửi).

Giữ kích thước ảnh (width/height px) để Outlook không thu nhỏ lung tung.
"""

from __future__ import annotations

import base64
import io
import re
import tempfile
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

# PR_ATTACH_CONTENT_ID / PR_ATTACHMENT_HIDDEN (MAPI)
_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
_PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

# Match cả thẻ <img ...> có src=data:image
_IMG_TAG_RE = re.compile(
    r"""(?is)(<img\b)([^>]*?)(/?>)"""
)
_SRC_DATA_RE = re.compile(
    r"""(?is)src\s*=\s*["'](data:image/([a-z0-9.+-]+);base64,([A-Za-z0-9+/=\s]+))["']"""
)

# Giới hạn chiều rộng hợp lý trong mail (Outlook hay scale nếu quá lớn)
_MAX_EMAIL_WIDTH = 650


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


def _parse_px(value: str) -> Optional[int]:
    m = re.search(r"(\d+(?:\.\d+)?)\s*px", value or "", flags=re.I)
    if m:
        return max(1, int(round(float(m.group(1)))))
    m = re.fullmatch(r"\s*(\d+)\s*", value or "")
    if m:
        return max(1, int(m.group(1)))
    return None


def _attr(attrs: str, name: str) -> str:
    m = re.search(rf"""(?is)\b{name}\s*=\s*["']([^"']*)["']""", attrs)
    return m.group(1) if m else ""


def _style_prop(style: str, prop: str) -> str:
    m = re.search(rf"(?i)(?:^|;)\s*{re.escape(prop)}\s*:\s*([^;]+)", style or "")
    return (m.group(1) or "").strip() if m else ""


def _image_natural_size(raw: bytes) -> tuple[Optional[int], Optional[int]]:
    try:
        from PIL import Image

        with Image.open(io.BytesIO(raw)) as im:
            return int(im.width), int(im.height)
    except Exception:  # noqa: BLE001
        return None, None


def _resolve_display_size(
    attrs: str, nat_w: Optional[int], nat_h: Optional[int]
) -> tuple[Optional[int], Optional[int]]:
    """Ưu tiên width/height đã set trên thẻ; không thì dùng kích thước thật (cap max)."""
    style = _attr(attrs, "style")
    w = _parse_px(_attr(attrs, "width")) or _parse_px(_style_prop(style, "width"))
    h = _parse_px(_attr(attrs, "height")) or _parse_px(_style_prop(style, "height"))

    if w and h:
        return w, h
    if w and nat_w and nat_h and nat_w > 0:
        return w, max(1, int(round(w * nat_h / nat_w)))
    if h and nat_w and nat_h and nat_h > 0:
        return max(1, int(round(h * nat_w / nat_h))), h
    if nat_w and nat_h:
        if nat_w > _MAX_EMAIL_WIDTH:
            return _MAX_EMAIL_WIDTH, max(1, int(round(_MAX_EMAIL_WIDTH * nat_h / nat_w)))
        return nat_w, nat_h
    return w, h


def _rebuild_img_attrs(attrs: str, cid: str, width: Optional[int], height: Optional[int]) -> str:
    # Bỏ src cũ + max-width gây Outlook thu nhỏ
    attrs = re.sub(r"""(?is)\s*src\s*=\s*["'][^"']*["']""", "", attrs)
    attrs = re.sub(r"""(?is)\s*width\s*=\s*["'][^"']*["']""", "", attrs)
    attrs = re.sub(r"""(?is)\s*height\s*=\s*["'][^"']*["']""", "", attrs)

    style = _attr(attrs, "style")
    if style:
        style = re.sub(r"(?i)(?:^|;)\s*max-width\s*:[^;]*;?", ";", style)
        style = re.sub(r"(?i)(?:^|;)\s*width\s*:[^;]*;?", ";", style)
        style = re.sub(r"(?i)(?:^|;)\s*height\s*:[^;]*;?", ";", style)
        style = re.sub(r";{2,}", ";", style).strip("; ").strip()
        attrs = re.sub(r"""(?is)\s*style\s*=\s*["'][^"']*["']""", "", attrs)

    parts = [f'src="cid:{cid}"']
    style_bits: list[str] = []
    if width:
        parts.append(f'width="{width}"')
        style_bits.append(f"width:{width}px")
    if height:
        parts.append(f'height="{height}"')
        style_bits.append(f"height:{height}px")
    # Không dùng max-width:100% — Outlook hay thu nhỏ ảnh vì CSS này
    if style:
        style_bits.insert(0, style)
    if style_bits:
        parts.append(f'style="{";".join(style_bits)}"')

    rest = attrs.strip()
    if rest:
        return " " + " ".join(parts) + " " + rest
    return " " + " ".join(parts)


def replace_data_images_with_cid(html: str) -> tuple[str, list[InlineCidImage]]:
    """Thay data:image bằng cid:…; ghi file tạm; giữ width/height px."""
    if not html or "data:image/" not in html.lower():
        return html or "", []

    items: list[InlineCidImage] = []
    tmp_dir = Path(tempfile.mkdtemp(prefix="vt_inline_"))

    def _img_repl(m: re.Match[str]) -> str:
        open_tag, attrs, close = m.group(1), m.group(2), m.group(3)
        sm = _SRC_DATA_RE.search(attrs)
        if not sm:
            return m.group(0)

        subtype, b64 = sm.group(2), sm.group(3)
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

        nat_w, nat_h = _image_natural_size(raw)
        width, height = _resolve_display_size(attrs, nat_w, nat_h)
        items.append(InlineCidImage(path=path, cid=cid))
        new_attrs = _rebuild_img_attrs(attrs, cid, width, height)
        return f"{open_tag}{new_attrs}{close}"

    return _IMG_TAG_RE.sub(_img_repl, html), items


def attach_cid_images(mail_item: Any, items: list[InlineCidImage]) -> None:
    """Đính ảnh ẩn sau khi đã set HTMLBody (cid khớp Content-ID)."""
    for item in items:
        try:
            att = mail_item.Attachments.Add(str(item.path), 1)  # olByValue
            pa = att.PropertyAccessor
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
