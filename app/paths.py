from __future__ import annotations

import os
import sys
from pathlib import Path


def bundle_dir() -> Path:
    """Thư mục resource đóng gói (PyInstaller) hoặc project root khi dev."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent.parent


def user_data_dir() -> Path:
    """
    Dữ liệu người dùng (config, logs, attachments).
    Dev: project root · App đóng gói: Application Support / AppData.
    """
    if getattr(sys, "frozen", False):
        if sys.platform == "darwin":
            base = Path.home() / "Library" / "Application Support" / "VTRateMailSender"
        elif sys.platform == "win32":
            base = Path(os.environ.get("APPDATA", str(Path.home()))) / "VTRateMailSender"
        else:
            base = Path.home() / ".vt-rate-mail-sender"
    else:
        base = Path(__file__).resolve().parent.parent

    base.mkdir(parents=True, exist_ok=True)
    (base / "logs").mkdir(exist_ok=True)
    (base / "attachments").mkdir(exist_ok=True)
    return base


def web_dir() -> Path:
    return bundle_dir() / "app" / "web"
