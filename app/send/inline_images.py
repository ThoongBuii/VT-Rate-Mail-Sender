"""Windows Outlook — ảnh body qua WordEditor AddPicture (giống paste tay).

Chiến lược (tránh khung trống + ảnh / paste 2 lần):
1) Gỡ hết <img data:image> khỏi HTML, lưu file tạm
2) Chèn 1 neo ##VT_IMG_END## cuối body (trước chữ ký)
3) Sau HTMLBody: xóa neo → AddPicture lần lượt (1 ảnh = 1 shape)
4) Clear clipboard; xóa shape link/rỗng/trùng
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

_IMG_TAG_RE = re.compile(r"""(?is)<img\b[^>]*?/?>""")
_SRC_DATA_RE = re.compile(
    r"""(?is)src\s*=\s*["'](data:image/([a-z0-9.+-]+);base64,([A-Za-z0-9+/=\s]+))["']"""
)

IMG_END_MARKER = "##VT_IMG_END##"
_MAX_EMAIL_WIDTH = 650

# Word WdInlineShape types
_WD_INLINE_SHAPE_LINKED_PICTURE = 4
_WD_INLINE_SHAPE_LINKED_OLE = 2


@dataclass
class InlinePicture:
    path: Path
    width_px: Optional[int] = None
    height_px: Optional[int] = None
    marker: str = ""  # legacy unused


InlineCidImage = InlinePicture


def _ext_for_mime(subtype: str) -> str:
    sub = (subtype or "png").split("+", 1)[0].lower().strip()
    if sub in {"jpeg", "jpg"}:
        return "jpg"
    if sub in {"png", "gif", "bmp"}:
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


def _attr_from_tag(tag: str, name: str) -> str:
    m = re.search(rf"""(?is)\b{name}\s*=\s*["']([^"']*)["']""", tag)
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
    tag: str, nat_w: Optional[int], nat_h: Optional[int]
) -> tuple[Optional[int], Optional[int]]:
    style = _attr_from_tag(tag, "style")
    w = _parse_px(_attr_from_tag(tag, "width")) or _parse_px(_style_prop(style, "width"))
    h = _parse_px(_attr_from_tag(tag, "height")) or _parse_px(_style_prop(style, "height"))

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


def clear_windows_clipboard() -> None:
    try:
        import win32clipboard  # type: ignore

        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
        finally:
            win32clipboard.CloseClipboard()
        return
    except Exception:  # noqa: BLE001
        pass
    try:
        import ctypes

        user32 = ctypes.windll.user32  # type: ignore[attr-defined]
        if user32.OpenClipboard(None):
            try:
                user32.EmptyClipboard()
            finally:
                user32.CloseClipboard()
    except Exception:  # noqa: BLE001
        pass


def extract_data_images_as_placeholders(html: str) -> tuple[str, list[InlinePicture]]:
    """
    Gỡ mọi <img data:image> khỏi HTML → file tạm.
    Thêm 1 neo ##VT_IMG_END## cuối body (nếu có ảnh).
    """
    if not html or "data:image/" not in html.lower():
        return html or "", []

    items: list[InlinePicture] = []
    tmp_dir = Path(tempfile.mkdtemp(prefix="vt_inline_"))

    def _img_repl(tag_match: re.Match[str]) -> str:
        tag = tag_match.group(0)
        sm = _SRC_DATA_RE.search(tag)
        if not sm:
            # img không phải data URI — bỏ luôn để tránh khung vỡ
            return ""

        subtype, b64 = sm.group(2), sm.group(3)
        try:
            raw = base64.b64decode(re.sub(r"\s+", "", b64), validate=False)
        except Exception:  # noqa: BLE001
            return ""
        if not raw:
            return ""

        ext = _ext_for_mime(subtype)
        idx = len(items)
        path = tmp_dir / f"vt_img_{idx}_{uuid.uuid4().hex[:8]}.{ext}"
        path = _write_image_file(raw, path, want_ext=ext)
        nat_w, nat_h = _image_natural_size(raw)
        width, height = _resolve_display_size(tag, nat_w, nat_h)
        items.append(InlinePicture(path=path, width_px=width, height_px=height))
        return ""  # không để marker/img trong HTML → hết khung dashed trống

    new_html = _IMG_TAG_RE.sub(_img_repl, html)
    if items:
        new_html = (new_html or "").rstrip() + f'<p style="margin:0;">{IMG_END_MARKER}</p>'
    return new_html, items


def replace_data_images_with_cid(html: str) -> tuple[str, list[InlinePicture]]:
    return extract_data_images_as_placeholders(html)


def _find_text_range(word_doc: Any, text: str) -> Any:
    try:
        rng = word_doc.Content
        find = rng.Find
        find.ClearFormatting()
        try:
            find.Replacement.ClearFormatting()
        except Exception:  # noqa: BLE001
            pass
        find.Text = text
        find.Forward = True
        find.Wrap = 0
        find.Format = False
        find.MatchCase = True
        find.MatchWholeWord = False
        find.MatchWildcards = False
        if bool(find.Execute()):
            return rng
    except Exception:  # noqa: BLE001
        return None
    return None


def _shape_count(word_doc: Any) -> int:
    try:
        return int(word_doc.InlineShapes.Count)
    except Exception:  # noqa: BLE001
        return 0


def _apply_size(shape: Any, pic: InlinePicture) -> None:
    try:
        shape.LockAspectRatio = -1
    except Exception:  # noqa: BLE001
        pass
    try:
        if pic.width_px and pic.width_px > 0:
            shape.Width = float(pic.width_px) * 72.0 / 96.0
        elif pic.height_px and pic.height_px > 0:
            shape.Height = float(pic.height_px) * 72.0 / 96.0
    except Exception:  # noqa: BLE001
        pass


def _purge_bad_shapes(word_doc: Any, *, keep_min: int = 0) -> None:
    """
    Xóa shape gây khung trống:
    - Linked picture / linked OLE (hay thành dashed box)
    - Width/Height ~ 0
    Và gộp ảnh trùng kích thước liền kề (giữ 1).
    """
    try:
        n = _shape_count(word_doc)
    except Exception:  # noqa: BLE001
        return

    # Xóa linked / rỗng (từ cuối)
    for i in range(n, 0, -1):
        try:
            shape = word_doc.InlineShapes(i)
            stype = int(getattr(shape, "Type", 0) or 0)
            w = float(getattr(shape, "Width", 0) or 0)
            h = float(getattr(shape, "Height", 0) or 0)
            if stype in {_WD_INLINE_SHAPE_LINKED_PICTURE, _WD_INLINE_SHAPE_LINKED_OLE}:
                shape.Delete()
                continue
            if w < 2 or h < 2:
                shape.Delete()
                continue
        except Exception:  # noqa: BLE001
            continue

    # Ảnh trùng kích thước liền kề → xóa bản sau (paste đôi)
    try:
        n = _shape_count(word_doc)
    except Exception:  # noqa: BLE001
        return
    i = n
    while i >= 2:
        try:
            a = word_doc.InlineShapes(i - 1)
            b = word_doc.InlineShapes(i)
            aw, ah = float(a.Width or 0), float(a.Height or 0)
            bw, bh = float(b.Width or 0), float(b.Height or 0)
            if abs(aw - bw) < 1.5 and abs(ah - bh) < 1.5 and aw > 5:
                b.Delete()
                # không giảm i quá — đã xóa i, xét lại cặp mới
                i = min(i, _shape_count(word_doc))
                continue
        except Exception:  # noqa: BLE001
            pass
        i -= 1


def insert_pictures_via_word(word_doc: Any, pictures: list[InlinePicture]) -> int:
    """Xóa neo ##VT_IMG_END## → chèn lần lượt từng ảnh (không để img/marker trong HTML)."""
    if not word_doc or not pictures:
        return 0

    clear_windows_clipboard()
    inserted = 0

    rng = _find_text_range(word_doc, IMG_END_MARKER)
    if rng is None:
        # Fallback: chèn trước chữ ký / cuối document
        try:
            rng = word_doc.Content
            rng.Collapse(0)  # wdCollapseStart = 1, end = 0 — dùng end
        except Exception:  # noqa: BLE001
            return 0
        try:
            # Collapse to end of document content before signature is harder;
            # try Best Regards
            sig = _find_text_range(word_doc, "Best Regards")
            if sig is None:
                sig = _find_text_range(word_doc, "Best regards")
            if sig is not None:
                rng = sig
                rng.Collapse(1)  # wdCollapseStart — chèn ngay trước chữ ký
            else:
                rng = word_doc.Content
                rng.Collapse(0)  # end
        except Exception:  # noqa: BLE001
            return 0
    else:
        try:
            rng.Delete()
        except Exception:  # noqa: BLE001
            try:
                rng.Text = ""
            except Exception:  # noqa: BLE001
                return 0

    # Điểm neo cố định — chèn ngược để thứ tự ảnh đúng như Soạn
    try:
        anchor_pos = int(rng.Start)
    except Exception:  # noqa: BLE001
        anchor_pos = 0

    for pic in reversed(pictures):
        if not pic.path.is_file():
            continue
        try:
            clear_windows_clipboard()
            r = word_doc.Range(anchor_pos, anchor_pos)
            before = _shape_count(word_doc)
            shape = r.InlineShapes.AddPicture(
                str(pic.path.resolve()),
                False,  # LinkToFile=False — tránh khung dashed (linked)
                True,  # SaveWithDocument
            )
            after = _shape_count(word_doc)
            for _ in range(max(0, after - before - 1)):
                try:
                    word_doc.InlineShapes(_shape_count(word_doc)).Delete()
                except Exception:  # noqa: BLE001
                    break

            try:
                _apply_size(shape, pic)
            except Exception:  # noqa: BLE001
                try:
                    _apply_size(word_doc.InlineShapes(_shape_count(word_doc)), pic)
                except Exception:  # noqa: BLE001
                    pass

            try:
                shape.Range.InsertParagraphAfter()
            except Exception:  # noqa: BLE001
                pass

            inserted += 1
            clear_windows_clipboard()
        except Exception:  # noqa: BLE001
            continue

    # Dọn neo sót + shape xấu/trùng
    leftover = _find_text_range(word_doc, IMG_END_MARKER)
    if leftover is not None:
        try:
            leftover.Delete()
        except Exception:  # noqa: BLE001
            try:
                leftover.Text = ""
            except Exception:  # noqa: BLE001
                pass

    _purge_bad_shapes(word_doc)
    clear_windows_clipboard()
    return inserted


def attach_cid_images(mail_item: Any, items: list[InlinePicture]) -> None:
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
