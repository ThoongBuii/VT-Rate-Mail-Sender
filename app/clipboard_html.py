"""Đọc HTML clipboard Windows (Outlook CF_HTML) và nhúng ảnh local/cid thành data URI."""

from __future__ import annotations

import base64
import html as html_lib
import re
import platform
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


def _extract_cf_html_fragment(raw: str) -> str:
    """Parse Windows 'HTML Format' clipboard (CF_HTML)."""
    text = raw
    # CF_HTML thường bắt đầu bằng Version/StartHTML/... rồi payload
    start = text.find("<!--StartFragment-->")
    end = text.find("<!--EndFragment-->")
    if start != -1 and end != -1 and end > start:
        return text[start + len("<!--StartFragment-->") : end].strip()

    # Fallback theo offset StartHTML/EndHTML
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
            # Fallback text
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


def embed_local_and_cid_images(html: str, cid_map: Optional[dict[str, str]] = None) -> str:
    """
    Chuyển src=file:// / cid: / tên file local thành data URI.
    Outlook copy hay để ảnh tại %TEMP%\\msohtmlclip*\\...
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
            return cid_map.get(key, original)

        # file:///C:/... hoặc file://localhost/C:/...
        path: Optional[Path] = None
        if low.startswith("file:"):
            parsed = urlparse(s)
            path_str = unquote(parsed.path or "")
            # Windows: /C:/Users/... → C:/Users/...
            if re.match(r"^/[A-Za-z]:/", path_str):
                path_str = path_str[1:]
            path_str = path_str.replace("/", "\\") if platform.system() == "Windows" else path_str
            path = Path(path_str)
        else:
            # Đường dẫn tuyệt đối hoặc tên file trong temp Outlook
            candidate = Path(unquote(s))
            if candidate.is_file():
                path = candidate
            else:
                # Tìm trong %TEMP%/msohtmlclip*
                name = candidate.name
                if name and platform.system() == "Windows":
                    import tempfile

                    temp = Path(tempfile.gettempdir())
                    matches: list[Path] = []
                    for clip_dir in temp.glob("msohtmlclip*"):
                        matches.extend(clip_dir.rglob(name))
                    if not matches:
                        # Một số bản Outlook để file trực tiếp trong Temp
                        matches = list(temp.glob(name))
                    matches = sorted(
                        matches,
                        key=lambda p: p.stat().st_mtime if p.exists() else 0,
                        reverse=True,
                    )
                    if matches:
                        path = matches[0]

        if path is not None:
            uri = _file_to_data_uri(path)
            if uri:
                return uri

        # khớp theo basename với cid_map keys chứa tên file
        base = Path(unquote(s.split("?")[0])).name.lower()
        if base:
            for key, uri in cid_map.items():
                if base in key or key in base:
                    return uri
        return original

    def _repl(match: re.Match[str]) -> str:
        src = match.group(1)
        resolved = _resolve_src(src)
        quote = '"' if '"' in match.group(0) or True else "'"
        # giữ nguyên attribute name/case gần đúng
        return f'src={quote}{resolved}{quote}'

    return re.sub(r"""src\s*=\s*["']([^"']+)["']""", _repl, html, flags=re.I)


def clipboard_html_for_compose(browser_html: str = "") -> str:
    """
    Ưu tiên CF_HTML Windows (Outlook), fallback HTML từ browser paste event.
    Luôn cố nhúng ảnh file:// tạm của Outlook.
    """
    html = ""
    if platform.system() == "Windows":
        try:
            html = read_windows_cf_html()
        except Exception:  # noqa: BLE001
            html = ""
    if not (html or "").strip():
        html = browser_html or ""
    html = embed_local_and_cid_images(html)
    return html.strip()
