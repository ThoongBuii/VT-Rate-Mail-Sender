"""Windows Outlook — ảnh body: tách file tạm + placeholder, chèn bằng WordEditor (giống paste tay).

Không dùng CID Attachments (Outlook hay hiện khung trống + file lộ Attachment bar).
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

_IMG_TAG_RE = re.compile(r"""(?is)(<img\b)([^>]*?)(/?>)""")
_SRC_DATA_RE = re.compile(
    r"""(?is)src\s*=\s*["'](data:image/([a-z0-9.+-]+);base64,([A-Za-z0-9+/=\s]+))["']"""
)

_MAX_EMAIL_WIDTH = 650


@dataclass
class InlinePicture:
    """Ảnh sẽ chèn InlineShapes — marker tìm trong Word sau khi set HTMLBody."""

    path: Path
    marker: str
    width_px: Optional[int] = None
    height_px: Optional[int] = None


# Alias cũ (windows.py / cleanup) — cùng dataclass
InlineCidImage = InlinePicture


def _ext_for_mime(subtype: str) -> str:
    sub = (subtype or "png").split("+", 1)[0].lower().strip()
    if sub in {"jpeg", "jpg"}:
        return "jpg"
    if sub in {"png", "gif", "bmp"}:
        return sub
    # webp / khác → png (Word/Outlook ổn định hơn)
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


def _write_image_file(raw: bytes, path: Path, *, want_ext: str) -> Path:
    """Ghi file ảnh; webp/lạ → PNG qua Pillow."""
    try:
        from PIL import Image

        with Image.open(io.BytesIO(raw)) as im:
            if want_ext == "jpg":
                if im.mode in ("RGBA", "P", "LA"):
                    im = im.convert("RGB")
                im.save(path, format="JPEG", quality=92)
            elif want_ext == "png" or want_ext not in {"gif", "bmp"}:
                if want_ext != "png":
                    path = path.with_suffix(".png")
                im.save(path, format="PNG")
            else:
                im.save(path)
            return path
    except Exception:  # noqa: BLE001
        path.write_bytes(raw)
        return path


def extract_data_images_as_placeholders(html: str) -> tuple[str, list[InlinePicture]]:
    """
    Thay <img src=data:image...> bằng marker text ##VTIMG_n##.
    WordEditor sẽ Find marker → InlineShapes.AddPicture (giống paste tay).
    """
    if not html or "data:image/" not in html.lower():
        return html or "", []

    items: list[InlinePicture] = []
    tmp_dir = Path(tempfile.mkdtemp(prefix="vt_inline_"))
    counter = {"n": 0}

    def _img_repl(m: re.Match[str]) -> str:
        attrs = m.group(2)
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
        idx = counter["n"]
        counter["n"] = idx + 1
        marker = f"##VTIMG_{idx}_{uuid.uuid4().hex[:8]}##"
        path = tmp_dir / f"vt_img_{idx}.{ext}"
        path = _write_image_file(raw, path, want_ext=ext)

        nat_w, nat_h = _image_natural_size(raw)
        width, height = _resolve_display_size(attrs, nat_w, nat_h)
        items.append(
            InlinePicture(path=path, marker=marker, width_px=width, height_px=height)
        )
        # Marker plain text — Word Find ổn định; ảnh thường đứng riêng block như paste tay
        return f'<p style="margin:0 0 0.6em;">{marker}</p>'

    return _IMG_TAG_RE.sub(_img_repl, html), items


# Tên cũ — một số chỗ còn gọi
def replace_data_images_with_cid(html: str) -> tuple[str, list[InlinePicture]]:
    return extract_data_images_as_placeholders(html)


def _find_marker_range(word_doc: Any, marker: str) -> Any:
    """Tìm Range chứa marker; None nếu không thấy."""
    try:
        rng = word_doc.Content
        find = rng.Find
        find.ClearFormatting()
        try:
            find.Replacement.ClearFormatting()
        except Exception:  # noqa: BLE001
            pass
        find.Text = marker
        find.Forward = True
        find.Wrap = 0  # wdFindStop
        find.Format = False
        find.MatchCase = True
        find.MatchWholeWord = False
        find.MatchWildcards = False
        if bool(find.Execute()):
            return rng
    except Exception:  # noqa: BLE001
        return None
    return None


def insert_pictures_via_word(word_doc: Any, pictures: list[InlinePicture]) -> int:
    """
    Tìm marker trong WordEditor → xóa → AddPicture (inline, giống paste).
    Trả số ảnh chèn thành công.
    """
    if not word_doc or not pictures:
        return 0

    inserted = 0
    for pic in pictures:
        if not pic.path.is_file():
            continue
        try:
            rng = _find_marker_range(word_doc, pic.marker)
            if rng is None:
                continue

            rng.Text = ""
            shape = rng.InlineShapes.AddPicture(
                str(pic.path.resolve()),
                False,  # LinkToFile
                True,  # SaveWithDocument
            )
            try:
                # msoTrue = -1 — giữ tỉ lệ như paste tay
                shape.LockAspectRatio = -1
            except Exception:  # noqa: BLE001
                pass

            # Word dùng point (px → pt: *72/96) — ưu tiên Width, giữ aspect
            if pic.width_px and pic.width_px > 0:
                shape.Width = float(pic.width_px) * 72.0 / 96.0
            elif pic.height_px and pic.height_px > 0:
                shape.Height = float(pic.height_px) * 72.0 / 96.0

            inserted += 1
        except Exception:  # noqa: BLE001
            continue
    return inserted


def attach_cid_images(mail_item: Any, items: list[InlinePicture]) -> None:
    """Deprecated — giữ stub để import cũ không gãy; Plan A không dùng."""
    return


def cleanup_temp_files(items: list[InlinePicture]) -> None:
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
