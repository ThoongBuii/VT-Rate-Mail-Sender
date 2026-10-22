"""Windows Outlook — ảnh body: paste clipboard vào Word (giống gửi tay).

1) Gỡ <img data:image> → lưu PNG vào thư mục app (không xóa ngay sau Send)
2) Neo ##VT_IMG_END## cuối body (trước chữ ký)
3) Xóa neo → đưa ảnh vào Clipboard (CF_DIB) → Range.Paste (đúng 1 lần / ảnh)
4) Clear clipboard; không đụng InlineShapes chữ ký
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

IMG_END_MARKER = "##VT_IMG_END##"
_MAX_EMAIL_WIDTH = 650


@dataclass
class InlinePicture:
    path: Path
    width_px: Optional[int] = None
    height_px: Optional[int] = None
    marker: str = ""


InlineCidImage = InlinePicture


def _inline_store_dir() -> Path:
    try:
        from ..paths import user_data_dir

        d = user_data_dir() / "inline_send"
    except Exception:  # noqa: BLE001
        d = Path(tempfile_fallback()) / "vt_inline_send"
    d.mkdir(parents=True, exist_ok=True)
    return d


def tempfile_fallback() -> str:
    import tempfile

    return tempfile.gettempdir()


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
    """Luôn lưu PNG RGB/RGBA ổn định cho Word paste."""
    from PIL import Image

    path = path.with_suffix(".png")
    with Image.open(io.BytesIO(raw)) as im:
        if im.mode not in ("RGB", "RGBA"):
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


def _set_clipboard_dib(path: Path) -> None:
    """Đưa ảnh vào clipboard dạng CF_DIB — Word Paste giống Ctrl+V thủ công."""
    from PIL import Image
    import win32clipboard  # type: ignore

    with Image.open(path) as im:
        if im.mode == "RGBA":
            bg = Image.new("RGB", im.size, (255, 255, 255))
            bg.paste(im, mask=im.split()[3])
            im = bg
        elif im.mode != "RGB":
            im = im.convert("RGB")
        buf = io.BytesIO()
        im.save(buf, "BMP")
        dib = buf.getvalue()[14:]  # bỏ BITMAPFILEHEADER

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, dib)
    finally:
        win32clipboard.CloseClipboard()


def extract_data_images_as_placeholders(html: str) -> tuple[str, list[InlinePicture]]:
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
            path = store / f"vt_{uuid.uuid4().hex}.png"
            path = _save_png(raw, path)
        except Exception:  # noqa: BLE001
            return ""
        nat_w, nat_h = _image_natural_size(raw)
        width, height = _resolve_display_size(tag, nat_w, nat_h)
        items.append(InlinePicture(path=path, width_px=width, height_px=height))
        return ""

    new_html = _IMG_TAG_RE.sub(_img_repl, html)
    if items:
        new_html = (new_html or "").rstrip() + f"<p>{IMG_END_MARKER}</p>"
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


def _find_new_shape_near(word_doc: Any, pos: int, before_starts: set[int]) -> Any:
    """Shape mới gần điểm paste (không đụng logo chữ ký phía sau)."""
    best = None
    best_dist = 10_000_000
    try:
        n = _shape_count(word_doc)
    except Exception:  # noqa: BLE001
        return None
    for i in range(1, n + 1):
        try:
            shape = word_doc.InlineShapes(i)
            st = int(shape.Range.Start)
            # Bỏ shape đã có trước (xấp xỉ theo Start cũ)
            if any(abs(st - old) <= 2 for old in before_starts):
                continue
            dist = abs(st - pos)
            if dist < best_dist:
                best_dist = dist
                best = shape
        except Exception:  # noqa: BLE001
            continue
    return best


def _collect_shape_starts(word_doc: Any) -> set[int]:
    starts: set[int] = set()
    try:
        for i in range(1, _shape_count(word_doc) + 1):
            try:
                starts.add(int(word_doc.InlineShapes(i).Range.Start))
            except Exception:  # noqa: BLE001
                continue
    except Exception:  # noqa: BLE001
        pass
    return starts


def _paste_one_image(word_doc: Any, anchor_pos: int, pic: InlinePicture) -> bool:
    """Clipboard CF_DIB + Range.Paste — giống Ctrl+V ảnh vào New Mail."""
    if not pic.path.is_file() or pic.path.stat().st_size < 32:
        return False

    clear_windows_clipboard()
    try:
        _set_clipboard_dib(pic.path)
    except Exception:  # noqa: BLE001
        clear_windows_clipboard()
        return False

    before_starts = _collect_shape_starts(word_doc)
    before_n = _shape_count(word_doc)
    try:
        r = word_doc.Range(anchor_pos, anchor_pos)
        r.Paste()
    except Exception:  # noqa: BLE001
        clear_windows_clipboard()
        # Fallback AddPicture nếu Paste fail
        try:
            r = word_doc.Range(anchor_pos, anchor_pos)
            shape = r.InlineShapes.AddPicture(str(pic.path.resolve()), False, True)
            _apply_size(shape, pic)
            clear_windows_clipboard()
            return True
        except Exception:  # noqa: BLE001
            clear_windows_clipboard()
            return False

    time.sleep(0.05)
    after_n = _shape_count(word_doc)
    shape = _find_new_shape_near(word_doc, anchor_pos, before_starts)

    # Paste đôi từ clip cũ → xóa shape thừa gần neo (chỉ shape mới, không đụng chữ ký xa)
    if after_n > before_n + 1 and shape is not None:
        try:
            # Xóa các shape mới khác (Start gần neo) trừ shape đã chọn
            keep_start = int(shape.Range.Start)
            for i in range(_shape_count(word_doc), 0, -1):
                try:
                    s = word_doc.InlineShapes(i)
                    st = int(s.Range.Start)
                    if abs(st - anchor_pos) < 80 and abs(st - keep_start) > 2:
                        # chỉ xóa nếu là shape mới
                        if not any(abs(st - old) <= 2 for old in before_starts):
                            s.Delete()
                except Exception:  # noqa: BLE001
                    continue
            shape = _find_new_shape_near(word_doc, anchor_pos, before_starts)
        except Exception:  # noqa: BLE001
            pass

    if shape is not None:
        _apply_size(shape, pic)
        try:
            shape.Range.InsertParagraphAfter()
        except Exception:  # noqa: BLE001
            pass
        clear_windows_clipboard()
        return True

    clear_windows_clipboard()
    return after_n > before_n


def insert_pictures_via_word(word_doc: Any, pictures: list[InlinePicture]) -> int:
    if not word_doc or not pictures:
        return 0

    clear_windows_clipboard()
    inserted = 0

    rng = _find_text_range(word_doc, IMG_END_MARKER)
    if rng is None:
        sig = _find_text_range(word_doc, "Best Regards") or _find_text_range(
            word_doc, "Best regards"
        )
        if sig is not None:
            rng = sig
            try:
                rng.Collapse(1)  # trước chữ ký
            except Exception:  # noqa: BLE001
                pass
        else:
            return 0
    else:
        try:
            rng.Delete()
        except Exception:  # noqa: BLE001
            try:
                rng.Text = ""
            except Exception:  # noqa: BLE001
                return 0

    try:
        anchor_pos = int(rng.Start)
    except Exception:  # noqa: BLE001
        return 0

    # Chèn ngược tại cùng neo → thứ tự cuối đúng như Soạn
    for pic in reversed(pictures):
        if _paste_one_image(word_doc, anchor_pos, pic):
            inserted += 1

    leftover = _find_text_range(word_doc, IMG_END_MARKER)
    if leftover is not None:
        try:
            leftover.Delete()
        except Exception:  # noqa: BLE001
            try:
                leftover.Text = ""
            except Exception:  # noqa: BLE001
                pass

    clear_windows_clipboard()
    return inserted


def attach_cid_images(mail_item: Any, items: list[InlinePicture]) -> None:
    return


def cleanup_temp_files(items: list[InlinePicture]) -> None:
    """
    Không xóa ngay file vừa Send (Outlook có thể vẫn đọc).
    Chỉ dọn file inline_send cũ hơn 2 giờ.
    """
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
