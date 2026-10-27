"""Mật khẩu SMTP — lưu local (không đưa vào config.json / git)."""

from __future__ import annotations

import os
from pathlib import Path

from .paths import user_data_dir


def _pass_path() -> Path:
    return user_data_dir() / "smtp.pass"


def get_smtp_password() -> str:
    env = (os.environ.get("VT_SMTP_PASSWORD") or "").strip()
    if env:
        return env
    p = _pass_path()
    if not p.is_file():
        return ""
    try:
        return p.read_text(encoding="utf-8").strip()
    except Exception:  # noqa: BLE001
        return ""


def set_smtp_password(password: str) -> None:
    raw = (password or "").strip()
    p = _pass_path()
    if not raw:
        try:
            p.unlink(missing_ok=True)
        except Exception:  # noqa: BLE001
            pass
        return
    p.write_text(raw, encoding="utf-8")
    try:
        os.chmod(p, 0o600)
    except Exception:  # noqa: BLE001
        pass


def has_smtp_password() -> bool:
    return bool(get_smtp_password())
