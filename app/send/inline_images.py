"""Windows Outlook — ảnh body đúng vị trí như gửi tay.

Cách làm:
1) Thay mỗi <img data:image> bằng marker text ##VTIMG_n_xxxx## (đúng chỗ ảnh)
2) Merge body + chữ ký Outlook (HTMLBody) — chữ ký không đụng
3) Xóa marker → InlineShapes.AddPicture (embed, không link) tại đúng chỗ
4) Xóa shape rỗng / shape cạnh cạnh (khung dashed kế ảnh) nếu có
Không dùng Clipboard Paste — Paste hay để lại khung trống cạnh ảnh.
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

# Word: wdInlineShapePicture=3, wdInlineShapeLinkedPicture=4
_WD_INLINE_PICTURE = 3
_WD_INLINE_LINKED_PICTURE = 4


@dataclass
class InlinePicture:
    path: Path
    marker: str
    width_px: Optional[int] = None
    height_px: Optional[int] = None


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
    """Thay mỗi ảnh bằng marker **đúng vị trí** trong HTML."""
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
        marker = f"##VTIMG_{idx}_{uuid.uuid4().hex[:10]}##"
        nat_w, nat_h = _image_natural_size(raw)
        width, height = _resolve_display_size(tag, nat_w, nat_h)
        items.append(
            InlinePicture(path=path, marker=marker, width_px=width, height_px=height)
        )
        return marker

    return _IMG_TAG_RE.sub(_img_repl, html), items


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


def _shape_area(shape: Any) -> float:
    try:
        return float(shape.Width) * float(shape.Height)
    except Exception:  # noqa: BLE001
        return 0.0


def _shape_type(shape: Any) -> int:
    try:
        return int(shape.Type)
    except Exception:  # noqa: BLE001
        return -1


def _looks_empty_frame(shape: Any) -> bool:
    """Khung dashed trống / linked hỏng — hay nằm kế ảnh thật."""
    t = _shape_type(shape)
    if t == _WD_INLINE_LINKED_PICTURE:
        return True
    try:
        w = float(shape.Width)
        h = float(shape.Height)
    except Exception:  # noqa: BLE001
        return False
    if w < 4 or h < 4:
        return True
    # Khung placeholder Outlook đôi khi có size nhưng không có ảnh thật
    try:
        # ScaleHeight/ScaleWidth bất thường hoặc PictureFormat lỗi
        _ = shape.PictureFormat
    except Exception:  # noqa: BLE001
        return True
    return False


def _delete_neighbor_shapes(
    word_doc: Any, keep_start: int, *, window: int = 120
) -> None:
    """Giữ 1 ảnh tốt tại chỗ chèn; xóa khung trống / bản trùng kế bên."""
    candidates: list[tuple[int, int, float, bool, Any]] = []
    # (index, start, area, is_empty, shape)
    n = _shape_count(word_doc)
    for i in range(1, n + 1):
        try:
            s = word_doc.InlineShapes(i)
            st = int(s.Range.Start)
            if abs(st - keep_start) > window:
                continue
            candidates.append(
                (i, st, _shape_area(s), _looks_empty_frame(s), s)
            )
        except Exception:  # noqa: BLE001
            continue

    if len(candidates) <= 1:
        # Vẫn xóa nếu đúng 1 mà là empty/linked
        if len(candidates) == 1 and candidates[0][3]:
            try:
                candidates[0][4].Delete()
            except Exception:  # noqa: BLE001
                pass
        return

    # Chọn shape giữ lại: không empty, Type picture, area lớn nhất, gần keep_start
    keep = None
    best_score = None
    for item in candidates:
        _i, st, area, is_empty, s = item
        if is_empty:
            continue
        t = _shape_type(s)
        type_bonus = 100000.0 if t == _WD_INLINE_PICTURE else 0.0
        score = (type_bonus + area, -abs(st - keep_start))
        if best_score is None or score > best_score:
            best_score = score
            keep = item

    if keep is None:
        # Không có shape tốt — xóa hết empty gần đó
        for _i, _st, _a, is_empty, s in sorted(candidates, key=lambda x: -x[0]):
            if is_empty:
                try:
                    s.Delete()
                except Exception:  # noqa: BLE001
                    pass
        return

    keep_idx = keep[0]
    # Xóa từ index cao → thấp
    for i, _st, _a, _e, s in sorted(candidates, key=lambda x: -x[0]):
        if i == keep_idx:
            continue
        try:
            s.Delete()
        except Exception:  # noqa: BLE001
            pass


def _replace_marker_with_image(word_doc: Any, pic: InlinePicture) -> bool:
    """
    Tìm marker → xóa text → AddPicture embed tại chỗ.
    Không Paste clipboard (tránh khung dashed + ảnh kế bên).
    """
    if not pic.path.is_file() or pic.path.stat().st_size < 32:
        return False

    clear_windows_clipboard()

    rng = _find_text_range(word_doc, pic.marker)
    if rng is None:
        return False

    try:
        pos = int(rng.Start)
    except Exception:  # noqa: BLE001
        return False

    # Xóa marker trước — không để text lẫn với ảnh
    try:
        rng.Text = ""
    except Exception:  # noqa: BLE001
        try:
            rng.Delete()
        except Exception:  # noqa: BLE001
            return False

    clear_windows_clipboard()
    time.sleep(0.05)

    abs_path = str(pic.path.resolve())
    before_n = _shape_count(word_doc)

    try:
        r = word_doc.Range(pos, pos)
        # LinkToFile=False, SaveWithDocument=True → embed, không khung linked trống
        shape = r.InlineShapes.AddPicture(abs_path, False, True)
    except Exception:  # noqa: BLE001
        return False

    clear_windows_clipboard()
    time.sleep(0.05)

    try:
        keep_start = int(shape.Range.Start)
    except Exception:  # noqa: BLE001
        keep_start = pos

    _apply_size(shape, pic)

    # Xóa khung trống / shape trùng kế bên (đúng bug screenshot: box | ảnh)
    _delete_neighbor_shapes(word_doc, keep_start, window=150)

    # Marker sót
    leftover = _find_text_range(word_doc, pic.marker)
    if leftover is not None:
        try:
            leftover.Text = ""
        except Exception:  # noqa: BLE001
            try:
                leftover.Delete()
            except Exception:  # noqa: BLE001
                pass

    after_n = _shape_count(word_doc)
    return after_n >= before_n


def insert_pictures_via_word(word_doc: Any, pictures: list[InlinePicture]) -> int:
    if not word_doc or not pictures:
        return 0

    clear_windows_clipboard()
    inserted = 0
    insert_positions: list[int] = []

    for pic in pictures:
        try:
            rng = _find_text_range(word_doc, pic.marker)
            approx = int(rng.Start) if rng is not None else -1
            if _replace_marker_with_image(word_doc, pic):
                inserted += 1
                if approx >= 0:
                    insert_positions.append(approx)
        except Exception:  # noqa: BLE001
            continue

    # Pass cuối: dọn linked/empty còn sót sát chỗ vừa chèn
    for pos in insert_positions:
        try:
            _delete_neighbor_shapes(word_doc, pos, window=180)
        except Exception:  # noqa: BLE001
            pass

    clear_windows_clipboard()
    return inserted


def attach_cid_images(mail_item: Any, items: list[InlinePicture]) -> None:
    return


def cleanup_temp_files(items: list[InlinePicture]) -> None:
    """Không xóa file vừa gửi; chỉ dọn bản cũ > 2 giờ."""
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
