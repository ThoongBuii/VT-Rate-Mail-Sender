"""Windows Outlook — ảnh body: A (CID MAPI) + C (AddPicture/BreakLink repair).

A — Chính (thứ tự quan trọng):
  1) data:image → <img src="cid:..." alt="VTIMG:...">
  2) Attachments.Add + PR_ATTACH_CONTENT_ID + ATT_MHTML_REF
  3) Mới gán HTMLBody (để Outlook bind cid, tránh khung trống sớm)

C — Repair nếu còn khung hỏng / linked trong body (trước Best Regards):
  BreakLink → nếu vẫn hỏng: xóa shape → AddPicture embed + BreakLink
"""

from __future__ import annotations

import base64
import io
import re
import time
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

_IMG_TAG_RE = re.compile(r"""(?is)<img\b[^>]*?/?>""")
_SRC_DATA_RE = re.compile(
    r"""(?is)src\s*=\s*["'](data:image/([a-z0-9.+-]+);base64,([A-Za-z0-9+/=\s]+))["']"""
)

_MAX_EMAIL_WIDTH = 650

_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
_PR_ATTACH_MIME_TAG = "http://schemas.microsoft.com/mapi/proptag/0x370E001F"
_PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003"  # 4 = ATT_MHTML_REF

_WD_INLINE_PICTURE = 3
_WD_INLINE_LINKED_PICTURE = 4

_ALT_PREFIX = "VTIMG:"


@dataclass
class InlinePicture:
    path: Path
    cid: str
    width_px: Optional[int] = None
    height_px: Optional[int] = None

    @property
    def marker(self) -> str:
        return self.cid

    @property
    def alt_key(self) -> str:
        return f"{_ALT_PREFIX}{self.cid}"


InlineCidImage = InlinePicture


def _inline_store_dir() -> Path:
    try:
        from ..paths import user_data_dir

        d = user_data_dir() / "inline_send"
    except Exception:  # noqa: BLE001
        import tempfile

        d = Path(tempfile.gettempdir()) / "vt_inline_send"
    d.mkdir(parents=True, exist_ok=True)
    return d


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


def _save_png(raw: bytes, path: Path) -> Path:
    from PIL import Image

    path = path.with_suffix(".png")
    with Image.open(io.BytesIO(raw)) as im:
        if im.mode == "RGBA":
            pass
        elif im.mode != "RGB":
            im = im.convert("RGBA") if "A" in im.getbands() else im.convert("RGB")
        im.save(path, format="PNG", optimize=True)
    if not path.is_file() or path.stat().st_size < 32:
        raise RuntimeError(f"Không ghi được ảnh: {path}")
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
    """Thay data:image bằng <img src=\"cid:...\"> đúng vị trí (A)."""
    if not html or "data:image/" not in html.lower():
        return html or "", []

    items: list[InlinePicture] = []
    store = _inline_store_dir()

    def _img_repl(tag_match: re.Match[str]) -> str:
        tag = tag_match.group(0)
        sm = _SRC_DATA_RE.search(tag)
        if not sm:
            return ""
        try:
            raw = base64.b64decode(re.sub(r"\s+", "", sm.group(3)), validate=False)
        except Exception:  # noqa: BLE001
            return ""
        if not raw:
            return ""
        try:
            path = _save_png(raw, store / f"vt_{uuid.uuid4().hex}.png")
        except Exception:  # noqa: BLE001
            return ""

        idx = len(items)
        cid = f"vt.img.{idx}.{uuid.uuid4().hex[:10]}"
        nat_w, nat_h = _image_natural_size(raw)
        width, height = _resolve_display_size(tag, nat_w, nat_h)
        pic = InlinePicture(path=path, cid=cid, width_px=width, height_px=height)
        items.append(pic)

        dims = ""
        style_bits: list[str] = ["border:0"]
        if width:
            dims += f' width="{width}"'
            style_bits.append(f"width:{width}px")
        if height:
            dims += f' height="{height}"'
            style_bits.append(f"height:{height}px")
        style = ";".join(style_bits)
        # alt = khóa để Word C tìm đúng shape
        return (
            f'<img src="cid:{cid}"{dims} alt="{pic.alt_key}" '
            f'border="0" style="{style}" />'
        )

    return _IMG_TAG_RE.sub(_img_repl, html), items


def replace_data_images_with_cid(html: str) -> tuple[str, list[InlinePicture]]:
    return extract_data_images_as_placeholders(html)


def _normalize_cid(value: str) -> str:
    return (value or "").strip().strip("<>").lower()


def _existing_attach_cids(mail_item: Any) -> set[str]:
    found: set[str] = set()
    try:
        n = int(mail_item.Attachments.Count)
    except Exception:  # noqa: BLE001
        return found
    for i in range(1, n + 1):
        try:
            att = mail_item.Attachments.Item(i)
            pa = getattr(att, "PropertyAccessor", None)
            if pa is None:
                continue
            raw = str(pa.GetProperty(_PR_ATTACH_CONTENT_ID) or "")
            key = _normalize_cid(raw)
            if key:
                found.add(key)
        except Exception:  # noqa: BLE001
            continue
    return found


def _set_inline_attach_props(att: Any, cid: str) -> bool:
    pa = getattr(att, "PropertyAccessor", None)
    if pa is None:
        return False
    cid_ok = False
    for cid_val in (cid, f"<{cid}>"):
        try:
            pa.SetProperty(_PR_ATTACH_CONTENT_ID, cid_val)
            cid_ok = True
            break
        except Exception:  # noqa: BLE001
            continue
    if not cid_ok:
        return False
    try:
        pa.SetProperty(_PR_ATTACH_FLAGS, 4)
    except Exception:  # noqa: BLE001
        pass
    try:
        pa.SetProperty(_PR_ATTACH_MIME_TAG, "image/png")
    except Exception:  # noqa: BLE001
        pass
    return True


def attach_cid_images(mail_item: Any, items: list[InlinePicture]) -> int:
    """A: đính PNG + Content-ID + ATT_MHTML_REF (không đính trùng cid)."""
    if not mail_item or not items:
        return 0

    have = _existing_attach_cids(mail_item)
    attached = 0
    for pic in items:
        if not pic.path.is_file() or pic.path.stat().st_size < 32:
            continue
        key = _normalize_cid(pic.cid)
        if key in have:
            # Đã có — chỉ khẳng định lại flags
            try:
                for i in range(1, int(mail_item.Attachments.Count) + 1):
                    att = mail_item.Attachments.Item(i)
                    pa = getattr(att, "PropertyAccessor", None)
                    if pa is None:
                        continue
                    raw = _normalize_cid(str(pa.GetProperty(_PR_ATTACH_CONTENT_ID) or ""))
                    if raw == key:
                        _set_inline_attach_props(att, pic.cid)
                        break
            except Exception:  # noqa: BLE001
                pass
            continue
        try:
            att = mail_item.Attachments.Add(str(pic.path.resolve()), 1)  # olByValue
        except Exception:  # noqa: BLE001
            continue
        if _set_inline_attach_props(att, pic.cid):
            have.add(key)
            attached += 1

    return attached


def _shape_count(word_doc: Any) -> int:
    try:
        return int(word_doc.InlineShapes.Count)
    except Exception:  # noqa: BLE001
        return 0


def _shape_type(shape: Any) -> int:
    try:
        return int(shape.Type)
    except Exception:  # noqa: BLE001
        return -1


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
    try:
        shape.AlternativeText = pic.alt_key
    except Exception:  # noqa: BLE001
        pass


def _force_embed(shape: Any) -> None:
    """C: ép lưu trong document, cắt link file ngoài."""
    try:
        lf = shape.LinkFormat
        try:
            lf.SavePictureWithDocument = True
        except Exception:  # noqa: BLE001
            pass
        try:
            lf.BreakLink()
        except Exception:  # noqa: BLE001
            pass
    except Exception:  # noqa: BLE001
        pass


def _link_source(shape: Any) -> str:
    try:
        return str(shape.LinkFormat.SourceFullName or "").strip()
    except Exception:  # noqa: BLE001
        return ""


def _looks_broken_or_linked(shape: Any) -> bool:
    t = _shape_type(shape)
    if t == _WD_INLINE_LINKED_PICTURE:
        return True
    if _link_source(shape):
        return True
    try:
        w = float(shape.Width)
        h = float(shape.Height)
        if w < 4 or h < 4:
            return True
    except Exception:  # noqa: BLE001
        return True
    try:
        sh = float(shape.ScaleHeight)
        sw = float(shape.ScaleWidth)
        if sh == 0 and sw == 0:
            return True
    except Exception:  # noqa: BLE001
        return True
    try:
        _ = shape.PictureFormat.Brightness
    except Exception:  # noqa: BLE001
        return True
    return False


def _signature_start(word_doc: Any) -> Optional[int]:
    try:
        rng = word_doc.Content
        find = rng.Find
        find.ClearFormatting()
        find.Text = "Best Regards"
        find.Forward = True
        find.Wrap = 0
        find.MatchCase = False
        if bool(find.Execute()):
            return int(rng.Start)
    except Exception:  # noqa: BLE001
        pass
    try:
        end = int(word_doc.Content.End)
        return max(0, int(end * 0.65))
    except Exception:  # noqa: BLE001
        return None


def _shape_alt(shape: Any) -> str:
    try:
        return str(shape.AlternativeText or "")
    except Exception:  # noqa: BLE001
        return ""


def _find_shape_for_pic(word_doc: Any, pic: InlinePicture) -> Any:
    """Tìm shape theo alt VTIMG:cid — fallback shape body thứ tự."""
    alt = pic.alt_key
    sig = _signature_start(word_doc)
    body: list[Any] = []
    for i in range(1, _shape_count(word_doc) + 1):
        try:
            s = word_doc.InlineShapes(i)
            st = int(s.Range.Start)
            if sig is not None and st >= sig:
                continue
            if _shape_alt(s) == alt or alt in _shape_alt(s):
                return s
            body.append(s)
        except Exception:  # noqa: BLE001
            continue
    return None


def _body_shapes(word_doc: Any) -> list[Any]:
    sig = _signature_start(word_doc)
    out: list[tuple[int, Any]] = []
    for i in range(1, _shape_count(word_doc) + 1):
        try:
            s = word_doc.InlineShapes(i)
            st = int(s.Range.Start)
            if sig is not None and st >= sig:
                continue
            out.append((st, s))
        except Exception:  # noqa: BLE001
            continue
    out.sort(key=lambda x: x[0])
    return [s for _st, s in out]


def _add_picture_embedded(word_doc: Any, pos: int, pic: InlinePicture) -> bool:
    if not pic.path.is_file() or pic.path.stat().st_size < 32:
        return False
    clear_windows_clipboard()
    try:
        r = word_doc.Range(pos, pos)
        shape = r.InlineShapes.AddPicture(str(pic.path.resolve()), False, True)
        _force_embed(shape)
        _apply_size(shape, pic)
        clear_windows_clipboard()
        return True
    except Exception:  # noqa: BLE001
        clear_windows_clipboard()
        return False


def _replace_shape_with_file(word_doc: Any, shape: Any, pic: InlinePicture) -> bool:
    try:
        pos = int(shape.Range.Start)
    except Exception:  # noqa: BLE001
        return False
    try:
        shape.Delete()
    except Exception:  # noqa: BLE001
        return False
    return _add_picture_embedded(word_doc, pos, pic)


def repair_broken_images_via_word(
    word_doc: Any, pictures: list[InlinePicture]
) -> int:
    """
    C: với mỗi ảnh body —
    1) BreakLink nếu đang link
    2) Nếu vẫn broken/linked → xóa + AddPicture embed
    3) Nếu không thấy shape (CID không tạo shape) → bỏ qua (tránh dán nhầm chữ ký)
       trừ khi có đúng 1 picture và 1 body shape hỏng theo thứ tự
    """
    if not word_doc or not pictures:
        return 0

    clear_windows_clipboard()
    repaired = 0

    # Pass 1: BreakLink trên shape khớp từng pic (hoặc mọi linked trong body)
    for pic in pictures:
        shape = _find_shape_for_pic(word_doc, pic)
        if shape is not None:
            _force_embed(shape)
            try:
                shape.AlternativeText = pic.alt_key
            except Exception:  # noqa: BLE001
                pass

    for s in _body_shapes(word_doc):
        if _link_source(s) or _shape_type(s) == _WD_INLINE_LINKED_PICTURE:
            _force_embed(s)

    time.sleep(0.12)

    # Pass 2: thay shape vẫn hỏng
    used_shape_ids: set[int] = set()
    for idx, pic in enumerate(pictures):
        shape = _find_shape_for_pic(word_doc, pic)
        if shape is None:
            # Fallback theo thứ tự body
            body = _body_shapes(word_doc)
            if idx < len(body):
                shape = body[idx]
            else:
                continue

        try:
            sid = int(shape.Range.Start)
        except Exception:  # noqa: BLE001
            sid = id(shape)
        if sid in used_shape_ids:
            continue

        if not _looks_broken_or_linked(shape):
            # Có thể vẫn là khung đỏ X nhưng Type=Picture — ép embed thêm lần
            _force_embed(shape)
            if not _looks_broken_or_linked(shape):
                used_shape_ids.add(sid)
                continue

        if _replace_shape_with_file(word_doc, shape, pic):
            repaired += 1
            used_shape_ids.add(sid)
        time.sleep(0.05)

    # Pass 3: body còn shape linked/broken → thay 1-1 theo thứ tự pictures
    still = [s for s in _body_shapes(word_doc) if _looks_broken_or_linked(s)]
    if still:
        for s, pic in zip(still, pictures):
            if _replace_shape_with_file(word_doc, s, pic):
                repaired += 1
            time.sleep(0.05)

    clear_windows_clipboard()
    return repaired


def insert_pictures_via_word(word_doc: Any, pictures: list[InlinePicture]) -> int:
    return repair_broken_images_via_word(word_doc, pictures)


def cleanup_temp_files(items: list[InlinePicture]) -> None:
    try:
        store = _inline_store_dir()
        now = time.time()
        for p in store.glob("vt_*.png"):
            try:
                if now - p.stat().st_mtime > 2 * 3600:
                    p.unlink(missing_ok=True)
            except Exception:  # noqa: BLE001
                pass
    except Exception:  # noqa: BLE001
        pass
