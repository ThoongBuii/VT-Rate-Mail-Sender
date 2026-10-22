"""Windows Outlook — ảnh body: placeholder + WordEditor AddPicture (giống paste tay).

Tránh paste 2 lần: xóa Clipboard trước khi AddPicture; nếu Word chèn thừa shape thì xóa bản dư.
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
_MARKER_LEFT_RE = re.compile(r"##VTIMG_\d+_[0-9a-f]+##", flags=re.I)

_MAX_EMAIL_WIDTH = 650


@dataclass
class InlinePicture:
    path: Path
    marker: str
    width_px: Optional[int] = None
    height_px: Optional[int] = None


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
    """Tránh AddPicture kích hoạt paste clipboard (ảnh Soạn còn trên clip → bị 2 lần)."""
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
    """Thay <img data:image> bằng marker ##VTIMG_n_xxxx##."""
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
        return f'<p style="margin:0 0 0.6em;">{marker}</p>'

    return _IMG_TAG_RE.sub(_img_repl, html), items


def replace_data_images_with_cid(html: str) -> tuple[str, list[InlinePicture]]:
    return extract_data_images_as_placeholders(html)


def _find_marker_range(word_doc: Any, marker: str) -> Any:
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


def _remove_leftover_markers(word_doc: Any) -> None:
    """Xóa mọi ##VTIMG_…## còn sót (tránh chữ marker / paste lệch)."""
    try:
        text = str(word_doc.Content.Text or "")
    except Exception:  # noqa: BLE001
        return
    for marker in _MARKER_LEFT_RE.findall(text):
        while True:
            rng = _find_marker_range(word_doc, marker)
            if rng is None:
                break
            try:
                rng.Delete()
            except Exception:  # noqa: BLE001
                try:
                    rng.Text = ""
                except Exception:  # noqa: BLE001
                    break


def _shape_count(word_doc: Any) -> int:
    try:
        return int(word_doc.InlineShapes.Count)
    except Exception:  # noqa: BLE001
        return 0


def insert_pictures_via_word(word_doc: Any, pictures: list[InlinePicture]) -> int:
    """
    Tìm marker → xóa → AddPicture đúng 1 lần.
    Clear clipboard trước mỗi lần chèn; nếu Word nhét thêm shape (paste clip) thì xóa bản dư.
    """
    if not word_doc or not pictures:
        return 0

    clear_windows_clipboard()
    inserted = 0

    for pic in pictures:
        if not pic.path.is_file():
            continue
        try:
            clear_windows_clipboard()
            rng = _find_marker_range(word_doc, pic.marker)
            if rng is None:
                continue

            before = _shape_count(word_doc)
            try:
                rng.Delete()
            except Exception:  # noqa: BLE001
                rng.Text = ""

            # Range sau Delete thường collapsed tại điểm chèn
            shape = rng.InlineShapes.AddPicture(
                str(pic.path.resolve()),
                False,  # LinkToFile
                True,  # SaveWithDocument
            )

            after = _shape_count(word_doc)
            # Paste clipboard thừa → xóa shape mới nhất (giữ shape vừa AddPicture)
            extra = after - before - 1
            for _ in range(max(0, extra)):
                try:
                    word_doc.InlineShapes(after).Delete()
                    after -= 1
                except Exception:  # noqa: BLE001
                    break

            try:
                shape.LockAspectRatio = -1  # msoTrue
            except Exception:  # noqa: BLE001
                pass

            if pic.width_px and pic.width_px > 0:
                try:
                    shape.Width = float(pic.width_px) * 72.0 / 96.0
                except Exception:  # noqa: BLE001
                    # shape có thể bị invalid sau dedupe — lấy shape cuối
                    try:
                        shape = word_doc.InlineShapes(_shape_count(word_doc))
                        shape.LockAspectRatio = -1
                        shape.Width = float(pic.width_px) * 72.0 / 96.0
                    except Exception:  # noqa: BLE001
                        pass
            elif pic.height_px and pic.height_px > 0:
                try:
                    shape.Height = float(pic.height_px) * 72.0 / 96.0
                except Exception:  # noqa: BLE001
                    pass

            inserted += 1
            clear_windows_clipboard()
        except Exception:  # noqa: BLE001
            continue

    _remove_leftover_markers(word_doc)
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
