"""Mac: data:image → file + cid (Outlook set content với data: hay thành external → banner privacy)."""

from __future__ import annotations

import base64
import re
import tempfile
from pathlib import Path
from typing import Optional


_DATA_IMG_SRC = re.compile(
    r"""(?is)src\s*=\s*(['"])(data:image/([a-z0-9.+-]+);base64,([A-Za-z0-9+/=\s]+))\1"""
)


def strip_vml_keep_html_imgs(html: str) -> str:
    """Bỏ khối VML Word (trùng ảnh); giữ nhánh <!--[if !vml]> <img>."""
    s = html or ""
    # Xóa khối if gte vml / if vml
    s = re.sub(
        r"(?is)<!--\[if[^\]]*?\bvml\b[^\]]*?\]>.*?<!\[endif\]-->",
        "",
        s,
    )
    # Gỡ wrapper if !vml
    s = re.sub(r"(?is)<!--\[if\s*!vml\]>", "", s)
    s = re.sub(r"(?is)<!\[endif\]-->", "", s)
    # VML còn sót
    s = re.sub(r"(?is)</?v:[^>]*>", "", s)
    s = re.sub(r"(?is)<v:[^>]+/>", "", s)
    return s


def strip_external_image_hooks(html: str) -> str:
    """Gỡ background=/url(http) còn sót — không để Outlook coi là external image."""
    s = html or ""

    def _bg_attr(m: re.Match[str]) -> str:
        val = (m.group(2) or "").strip()
        if val.lower().startswith("data:"):
            return m.group(0)
        return ""

    s = re.sub(
        r"""(?is)\sbackground\s*=\s*(['"])([^'"]*)\1""",
        _bg_attr,
        s,
    )

    def _css_url(m: re.Match[str]) -> str:
        inner = (m.group(1) or "").strip().strip("'\"")
        if inner.lower().startswith("data:"):
            return m.group(0)
        if inner.startswith("#"):  # VML #default#VML
            return m.group(0)
        return "url(about:blank)"

    s = re.sub(r"""(?is)url\(\s*([^)]+?)\s*\)""", _css_url, s)
    return s


def data_uris_to_cid_files(
    html: str, tmp_dir: Path
) -> tuple[str, list[tuple[str, Path]]]:
    """
    Đổi mọi data:image → cid:vtN@vt.local + file trên đĩa.
    Trả (html_mới, [(cid, path), ...]).
    """
    tmp_dir.mkdir(parents=True, exist_ok=True)
    parts: list[tuple[str, Path]] = []
    # dedupe identical base64 → cùng cid
    seen: dict[str, str] = {}

    def _repl(m: re.Match[str]) -> str:
        q = m.group(1)
        full_data = m.group(2)
        subtype = (m.group(3) or "png").split("+")[0].lower()
        b64 = re.sub(r"\s+", "", m.group(4) or "")
        if full_data in seen:
            return f"src={q}cid:{seen[full_data]}{q}"
        try:
            raw = base64.b64decode(b64, validate=False)
        except Exception:  # noqa: BLE001
            return m.group(0)
        if not raw:
            return m.group(0)
        ext = "jpg" if subtype in {"jpeg", "jpg"} else ("gif" if subtype == "gif" else "png")
        idx = len(parts)
        cid = f"vt{idx}@vt.local"
        path = tmp_dir / f"vt{idx}.{ext}"
        path.write_bytes(raw)
        parts.append((cid, path))
        seen[full_data] = cid
        return f"src={q}cid:{cid}{q}"

    out = _DATA_IMG_SRC.sub(_repl, html or "")
    return out, parts


def prepare_mac_send_html(html: str, tmp_dir: Path) -> tuple[str, list[tuple[str, Path]]]:
    """HTML gửi Mac: bỏ VML → data→cid files → gỡ hook external."""
    s = strip_vml_keep_html_imgs(html or "")
    s = strip_external_image_hooks(s)
    s, parts = data_uris_to_cid_files(s, tmp_dir)
    # Không còn data: / http img
    leftover = re.findall(
        r"""(?is)<img\b[^>]*?\bsrc\s*=\s*['"](?!cid:)([^'"]+)['"]""",
        s,
    )
    for src in leftover:
        if src.lower().startswith("data:"):
            continue
        # xóa img ngoài còn sót (không gửi external)
        s = re.sub(
            re.escape(f'src="{src}"'),
            'src=""',
            s,
            flags=re.I,
        )
        s = re.sub(
            re.escape(f"src='{src}'"),
            "src=''",
            s,
            flags=re.I,
        )
    return s, parts
