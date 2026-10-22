"""Windows Outlook — ảnh body qua Word InsertFile (cách khác A/C).

Không gán HTMLBody cho ảnh (tránh khung cid trống / paste lệch).
1) data:image → PNG trên đĩa + <img src=\"file:///ABS_PATH\">
2) Ghi file .htm tạm
3) WordEditor.Range(0,0).InsertFile(htm) — chèn phía trên chữ ký New Mail
4) BreakLink mọi ảnh body (trước Best Regards) để embed trong mail
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
_WD_INLINE_LINKED_PICTURE = 4


@dataclass
class InlinePicture:
    path: Path
    cid: str = ""
    width_px: Optional[int] = None
    height_px: Optional[int] = None

    @property
    def marker(self) -> str:
        return self.cid or self.path.name


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


def _path_to_file_uri(path: Path) -> str:
    """file:///C:/... — Word InsertFile hiểu được ảnh local."""
    abs_path = path.resolve()
    # Windows: C:\a\b.png → file:///C:/a/b.png
    as_posix = abs_path.as_posix()
    if re.match(r"^[A-Za-z]:/", as_posix):
        return "file:///" + as_posix
    return abs_path.as_uri()


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
    """Thay data:image bằng <img src=file:///ABS> + danh sách file."""
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
        items.append(
            InlinePicture(path=path, cid=cid, width_px=width, height_px=height)
        )

        uri = _path_to_file_uri(path)
        dims = ""
        style_bits = ["border:0"]
        if width:
            dims += f' width="{width}"'
            style_bits.append(f"width:{width}px")
        if height:
            dims += f' height="{height}"'
            style_bits.append(f"height:{height}px")
        style = ";".join(style_bits)
        return (
            f'<img src="{uri}"{dims} alt="" border="0" style="{style}" />'
        )

    return _IMG_TAG_RE.sub(_img_repl, html), items


def replace_data_images_with_cid(html: str) -> tuple[str, list[InlinePicture]]:
    return extract_data_images_as_placeholders(html)


def write_body_htm_for_insert(html: str, pictures: list[InlinePicture]) -> Path:
    """Ghi .htm đầy đủ để Word InsertFile (ảnh đã là file:/// tuyệt đối)."""
    store = _inline_store_dir()
    name = f"vt_body_{uuid.uuid4().hex[:12]}.htm"
    path = store / name
    body = (html or "").strip()
    # Fragment đủ cho Word — tránh wrapper phức tạp làm lệch layout
    doc = (
        "<!DOCTYPE html>\n"
        '<html><head><meta charset="utf-8"></head>'
        f"<body>{body}</body></html>\n"
    )
    path.write_text(doc, encoding="utf-8", errors="replace")
    return path


def _shape_count(word_doc: Any) -> int:
    try:
        return int(word_doc.InlineShapes.Count)
    except Exception:  # noqa: BLE001
        return 0


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
        return max(0, int(int(word_doc.Content.End) * 0.65))
    except Exception:  # noqa: BLE001
        return None


def _force_embed(shape: Any) -> None:
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


def embed_body_pictures_after_insert(
    word_doc: Any, pictures: list[InlinePicture]
) -> int:
    """Sau InsertFile: BreakLink + chỉnh size ảnh body (không đụng chữ ký)."""
    if not word_doc:
        return 0
    sig = _signature_start(word_doc)
    embedded = 0
    body_shapes: list[Any] = []
    for i in range(1, _shape_count(word_doc) + 1):
        try:
            s = word_doc.InlineShapes(i)
            st = int(s.Range.Start)
            if sig is not None and st >= sig:
                continue
            body_shapes.append(s)
            _force_embed(s)
            embedded += 1
        except Exception:  # noqa: BLE001
            continue

    # Size theo thứ tự pictures
    for s, pic in zip(body_shapes, pictures):
        _apply_size(s, pic)
        _force_embed(s)

    return embedded


def insert_body_html_via_word_file(
    word_doc: Any, body_html: str
) -> tuple[list[InlinePicture], Optional[Path]]:
    """
    Chèn body bằng InsertFile tại đầu document (phía trên chữ ký).
    Trả về (pictures, htm_path).
    """
    if not word_doc:
        return [], None

    ready, pictures = extract_data_images_as_placeholders(body_html)
    htm_path = write_body_htm_for_insert(ready, pictures)

    clear_windows_clipboard()
    abs_htm = str(htm_path.resolve())
    try:
        # Chèn tại đầu — chữ ký New Mail bị đẩy xuống dưới
        # InsertFile(FileName, Range/Bookmark, ConfirmConversions, Link, Attachment)
        rng = word_doc.Range(0, 0)
        rng.InsertFile(abs_htm, "", False, False, False)
    except Exception:  # noqa: BLE001
        try:
            rng = word_doc.Range(0, 0)
            rng.InsertFile(abs_htm)
        except Exception:  # noqa: BLE001
            try:
                sel = word_doc.Application.Selection
                sel.HomeKey(Unit=6)  # wdStory
                sel.InsertFile(abs_htm)
            except Exception:  # noqa: BLE001
                clear_windows_clipboard()
                return pictures, htm_path

    time.sleep(0.2)
    if pictures:
        embed_body_pictures_after_insert(word_doc, pictures)
    clear_windows_clipboard()
    return pictures, htm_path


# --- API tương thích cũ (không dùng CID path nữa) ---


def attach_cid_images(mail_item: Any, items: list[InlinePicture]) -> int:
    return 0


def repair_broken_images_via_word(
    word_doc: Any, pictures: list[InlinePicture]
) -> int:
    return embed_body_pictures_after_insert(word_doc, pictures)


def insert_pictures_via_word(word_doc: Any, pictures: list[InlinePicture]) -> int:
    return embed_body_pictures_after_insert(word_doc, pictures)


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
        for p in store.glob("vt_body_*.htm"):
            try:
                if now - p.stat().st_mtime > 2 * 3600:
                    p.unlink(missing_ok=True)
            except Exception:  # noqa: BLE001
                pass
    except Exception:  # noqa: BLE001
        pass
