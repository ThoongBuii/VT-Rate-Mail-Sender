"""Đọc HTML clipboard (Windows CF_HTML / macOS WebArchive) và nhúng ảnh thành data URI."""

from __future__ import annotations

import base64
import html as html_lib
import platform
import re
import tempfile
import time
from pathlib import Path
from typing import Optional
from urllib.parse import unquote, urlparse


def _mime_for_name(name: str) -> str:
    lower = name.lower()
    if lower.endswith((".jpg", ".jpeg")):
        return "image/jpeg"
    if lower.endswith(".gif"):
        return "image/gif"
    if lower.endswith(".bmp"):
        return "image/bmp"
    if lower.endswith(".webp"):
        return "image/webp"
    if lower.endswith(".emf"):
        return "image/emf"
    if lower.endswith(".wmf"):
        return "image/wmf"
    return "image/png"


def _file_to_data_uri(path: Path) -> Optional[str]:
    try:
        if not path.is_file():
            return None
        raw = path.read_bytes()
        if not raw:
            return None
        mime = _mime_for_name(path.name)
        return f"data:{mime};base64,{base64.b64encode(raw).decode('ascii')}"
    except OSError:
        return None


def _bytes_to_data_uri(raw: bytes, mime: str = "image/png") -> str:
    return f"data:{mime};base64,{base64.b64encode(raw).decode('ascii')}"


def _extract_cf_html_fragment(raw: str) -> str:
    """Parse Windows 'HTML Format' clipboard (CF_HTML)."""
    text = raw
    start = text.find("<!--StartFragment-->")
    end = text.find("<!--EndFragment-->")
    if start != -1 and end != -1 and end > start:
        return text[start + len("<!--StartFragment-->") : end].strip()

    def _offset(key: str) -> Optional[int]:
        m = re.search(rf"{key}:(\d+)", text)
        return int(m.group(1)) if m else None

    s = _offset("StartHTML")
    e = _offset("EndHTML")
    if s is not None and e is not None and 0 <= s < e <= len(text):
        return text[s:e].strip()

    s = _offset("StartFragment")
    e = _offset("EndFragment")
    if s is not None and e is not None and 0 <= s < e <= len(text):
        return text[s:e].strip()

    return text.strip()


def read_windows_cf_html() -> str:
    """Đọc CF_HTML từ clipboard Windows. Trả về HTML fragment hoặc ''."""
    if platform.system() != "Windows":
        return ""
    try:
        import win32clipboard  # type: ignore
        import win32con  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Thiếu pywin32 để đọc clipboard Outlook.") from exc

    win32clipboard.OpenClipboard()
    try:
        fmt = win32clipboard.RegisterClipboardFormat("HTML Format")
        if not win32clipboard.IsClipboardFormatAvailable(fmt):
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                return html_lib.escape(win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT) or "")
            return ""
        data = win32clipboard.GetClipboardData(fmt)
        if isinstance(data, bytes):
            for enc in ("utf-8", "utf-16-le", "mbcs", "latin-1"):
                try:
                    raw = data.decode(enc)
                    break
                except UnicodeDecodeError:
                    raw = ""
            else:
                raw = data.decode("utf-8", errors="ignore")
        else:
            raw = str(data or "")
        return _extract_cf_html_fragment(raw)
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:  # noqa: BLE001
            pass


def _parse_mac_webarchive(data) -> tuple[str, dict[str, str]]:
    """Apple Web Archive → (html, cid_map url→dataURI). Giữ ảnh chữ ký khi Copy từ Outlook Mac."""
    try:
        from Foundation import NSPropertyListSerialization, NSPropertyListImmutable  # type: ignore
    except ImportError:
        return "", {}

    plist, _fmt, error = NSPropertyListSerialization.propertyListWithData_options_format_error_(
        data, NSPropertyListImmutable, None, None
    )
    if error or not plist:
        return "", {}

    main = plist.get("WebMainResource") or {}
    main_data = main.get("WebResourceData")
    html = ""
    if main_data is not None:
        try:
            html = bytes(main_data).decode("utf-8", errors="ignore")
        except Exception:  # noqa: BLE001
            html = str(main_data)

    cid_map: dict[str, str] = {}
    for sub in plist.get("WebSubresources") or []:
        try:
            url = str(sub.get("WebResourceURL") or "").strip()
            raw = sub.get("WebResourceData")
            mime = str(sub.get("WebResourceMIMEType") or "image/png")
            if not url or raw is None:
                continue
            if not mime.lower().startswith("image/"):
                continue
            blob = bytes(raw)
            if not blob:
                continue
            uri = _bytes_to_data_uri(blob, mime)
            cid_map[url.lower()] = uri
            if url.lower().startswith("cid:"):
                key = url[4:].strip().strip("<>").lower()
                cid_map[key] = uri
                base = key.split("@")[0]
                if base:
                    cid_map[base] = uri
            name = Path(unquote(urlparse(url).path or "")).name.lower()
            if name:
                cid_map[name] = uri
        except Exception:  # noqa: BLE001
            continue
    return html, cid_map


def _mac_recent_clip_image_map() -> dict[str, str]:
    """Tìm ảnh clip gần đây trong temp (Outlook đôi khi ghi file khi Copy)."""
    out: dict[str, str] = {}
    root = Path(tempfile.gettempdir())
    now = time.time()
    exts = {".png", ".jpg", ".jpeg", ".gif", ".webp"}

    candidates: list[Path] = []
    try:
        for p in root.rglob("*"):
            if not p.is_file():
                continue
            low = p.name.lower()
            path_l = str(p).lower()
            if p.suffix.lower() not in exts:
                continue
            if not (
                low.startswith("image00")
                or "signature_" in low
                or "msohtml" in path_l
                or "outlook" in path_l
            ):
                continue
            try:
                if now - p.stat().st_mtime > 600:
                    continue
            except OSError:
                continue
            candidates.append(p)
    except Exception:  # noqa: BLE001
        return out

    candidates.sort(key=lambda p: p.stat().st_mtime if p.exists() else 0, reverse=True)
    for p in candidates[:40]:
        uri = _file_to_data_uri(p)
        if not uri:
            continue
        out[p.name.lower()] = uri
        out[p.stem.lower()] = uri
    return out


def read_mac_clipboard_html() -> tuple[str, dict[str, str]]:
    """Đọc HTML clipboard macOS + map ảnh (WebArchive / file / temp)."""
    if platform.system() != "Darwin":
        return "", {}
    try:
        from AppKit import NSPasteboard, NSPasteboardTypeHTML, NSPasteboardTypeString  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Thiếu PyObjC (AppKit). Cài: pip install pyobjc-framework-Cocoa"
        ) from exc

    pb = NSPasteboard.generalPasteboard()
    cid_map: dict[str, str] = {}
    html = ""

    for wtype in (
        "com.apple.WebArchive",
        "Apple Web Archive pasteboard type",
    ):
        try:
            data = pb.dataForType_(wtype)
        except Exception:  # noqa: BLE001
            data = None
        if not data:
            continue
        h, m = _parse_mac_webarchive(data)
        if h.strip():
            html = h
            cid_map.update(m)
            break

    if not html.strip():
        html = pb.stringForType_(NSPasteboardTypeHTML) or ""
        if not html.strip():
            for t in ("public.html", "Apple HTML pasteboard type"):
                try:
                    html = pb.stringForType_(t) or ""
                except Exception:  # noqa: BLE001
                    html = ""
                if html.strip():
                    break

    try:
        from AppKit import (  # type: ignore
            NSBitmapImageRep,
            NSImage,
            NSPasteboardTypeTIFF,
            NSPNGFileType,
        )

        imgs = pb.readObjectsForClasses_options_([NSImage], None) or []
        for i, img in enumerate(imgs):
            try:
                tiff = img.TIFFRepresentation()
                if not tiff:
                    continue
                rep = NSBitmapImageRep.imageRepWithData_(tiff)
                if rep is None:
                    continue
                png = rep.representationUsingType_properties_(NSPNGFileType, None)
                if not png:
                    continue
                uri = _bytes_to_data_uri(bytes(png), "image/png")
                cid_map[f"clipboard_image_{i}"] = uri
                cid_map[f"image00{i + 1}.png"] = uri
                cid_map[f"image00{i + 1}.jpg"] = uri
                cid_map[f"image00{i + 1}"] = uri
            except Exception:  # noqa: BLE001
                continue
        tiff_data = pb.dataForType_(NSPasteboardTypeTIFF)
        if tiff_data and "clipboard_image_0" not in cid_map:
            try:
                img = NSImage.alloc().initWithData_(tiff_data)
                tiff = img.TIFFRepresentation() if img else None
                rep = NSBitmapImageRep.imageRepWithData_(tiff) if tiff else None
                png = (
                    rep.representationUsingType_properties_(NSPNGFileType, None)
                    if rep
                    else None
                )
                if png:
                    cid_map["clipboard_image_0"] = _bytes_to_data_uri(bytes(png), "image/png")
            except Exception:  # noqa: BLE001
                pass
    except Exception:  # noqa: BLE001
        pass

    cid_map.update(_mac_recent_clip_image_map())

    if not html.strip():
        plain = pb.stringForType_(NSPasteboardTypeString) or ""
        if plain.strip():
            safe = (
                plain.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\n", "<br>\n")
            )
            html = f"<div>{safe}</div>"

    return html, cid_map


def embed_local_and_cid_images(html: str, cid_map: Optional[dict[str, str]] = None) -> str:
    """
    Chuyển src=file:// / cid: / tên file local thành data URI.
    Outlook copy hay để ảnh tại %TEMP%\\msohtmlclip*\\... hoặc WebArchive Mac.
    """
    if not html:
        return html
    cid_map = {k.lower(): v for k, v in (cid_map or {}).items()}

    def _resolve_src(src: str) -> str:
        original = src
        s = src.strip().strip('"').strip("'")
        low = s.lower()

        if low.startswith("data:"):
            return original

        if low.startswith("cid:"):
            key = s[4:].strip().strip("<>").lower()
            if key in cid_map:
                return cid_map[key]
            base = key.split("@")[0]
            if base in cid_map:
                return cid_map[base]
            for k, uri in cid_map.items():
                if base and (base in k or k in base):
                    return uri
            return original

        if low in cid_map:
            return cid_map[low]

        path: Optional[Path] = None
        if low.startswith("file:"):
            parsed = urlparse(s)
            path_str = unquote(parsed.path or "")
            if re.match(r"^/[A-Za-z]:/", path_str):
                path_str = path_str[1:]
            path_str = path_str.replace("/", "\\") if platform.system() == "Windows" else path_str
            path = Path(path_str)
        else:
            candidate = Path(unquote(s))
            if candidate.is_file():
                path = candidate
            else:
                name = candidate.name
                if name and platform.system() == "Windows":
                    temp = Path(tempfile.gettempdir())
                    matches: list[Path] = []
                    for clip_dir in temp.glob("msohtmlclip*"):
                        matches.extend(clip_dir.rglob(name))
                    if not matches:
                        matches = list(temp.glob(name))
                    matches = sorted(
                        matches,
                        key=lambda p: p.stat().st_mtime if p.exists() else 0,
                        reverse=True,
                    )
                    if matches:
                        path = matches[0]
                elif name and name.lower() in cid_map:
                    return cid_map[name.lower()]

        if path is not None:
            uri = _file_to_data_uri(path)
            if uri:
                return uri

        base = Path(unquote(s.split("?")[0])).name.lower()
        if base:
            if base in cid_map:
                return cid_map[base]
            for key, uri in cid_map.items():
                if base in key or key in base:
                    return uri
        return original

    def _repl(match: re.Match[str]) -> str:
        src = match.group(1)
        resolved = _resolve_src(src)
        return f'src="{resolved}"'

    return re.sub(r"""src\s*=\s*["']([^"']+)["']""", _repl, html, flags=re.I)


def clipboard_html_for_compose(browser_html: str = "") -> str:
    """
    Windows: CF_HTML. macOS: WebArchive/HTML AppKit.
    Nhúng ảnh → data URI; giữ font/style gốc tối đa (không ép Aptos).
    """
    from .outlook_html import convert_px_font_sizes_to_pt, harden_tables_for_outlook

    html = ""
    cid_map: dict[str, str] = {}

    if platform.system() == "Windows":
        try:
            html = read_windows_cf_html()
        except Exception:  # noqa: BLE001
            html = ""
    elif platform.system() == "Darwin":
        try:
            html, cid_map = read_mac_clipboard_html()
        except Exception:  # noqa: BLE001
            html, cid_map = "", {}

    if not (html or "").strip():
        html = browser_html or ""

    html = embed_local_and_cid_images(html, cid_map)
    html = convert_px_font_sizes_to_pt(html)
    html = harden_tables_for_outlook(html)
    return html.strip()
