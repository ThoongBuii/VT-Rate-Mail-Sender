#!/usr/bin/env python3
"""VT Rate Mail Sender — cửa sổ desktop (pywebview hoặc Edge/Chrome --app)."""

from __future__ import annotations

import os
import shutil
import socket
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
from pathlib import Path


def _fix_stdio_encoding() -> None:
    if sys.platform != "win32":
        return
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except Exception:  # noqa: BLE001
            pass


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return int(s.getsockname()[1])


def _safe_print(msg: str) -> None:
    try:
        print(msg)
    except UnicodeEncodeError:
        print(msg.encode("ascii", errors="replace").decode("ascii"))


def _force_exit(code: int = 0) -> None:
    try:
        os._exit(code)
    except Exception:  # noqa: BLE001
        sys.exit(code)


def _log_startup(message: str) -> None:
    try:
        from app.paths import user_data_dir

        path = user_data_dir() / "logs" / "startup.log"
        with path.open("a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
    except Exception:  # noqa: BLE001
        pass


def _win_message(title: str, text: str, error: bool = True) -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        flags = 0x10 if error else 0x40
        ctypes.windll.user32.MessageBoxW(0, text, title, flags)
    except Exception:  # noqa: BLE001
        pass


def _find_edge_or_chrome() -> str | None:
    """Ưu tiên Microsoft Edge, rồi Google Chrome."""
    if sys.platform != "win32":
        return None

    candidates: list[str] = []
    local = os.environ.get("LOCALAPPDATA", "")
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")

    for base in (pf, pf86, local):
        candidates.extend(
            [
                str(Path(base) / "Microsoft" / "Edge" / "Application" / "msedge.exe"),
                str(Path(base) / "Google" / "Chrome" / "Application" / "chrome.exe"),
            ]
        )

    for name in ("msedge", "chrome"):
        found = shutil.which(name)
        if found:
            candidates.append(found)

    for path in candidates:
        if path and Path(path).is_file():
            return path
    return None


def _open_browser_app_window(url: str, width: int = 1320, height: int = 860) -> bool:
    """
    Mở Edge/Chrome kiểu ứng dụng (--app) — không hiện thanh địa chỉ.
    Chờ cửa sổ đóng rồi return True. False nếu không tìm thấy trình duyệt.
    """
    browser = _find_edge_or_chrome()
    if not browser:
        return False

    from app.paths import user_data_dir

    profile = user_data_dir() / "browser-app-profile"
    profile.mkdir(parents=True, exist_ok=True)

    cmd = [
        browser,
        f"--user-data-dir={profile}",
        f"--app={url}",
        f"--window-size={width},{height}",
        "--disable-features=TranslateUI",
        "--no-first-run",
        "--no-default-browser-check",
    ]
    _log_startup(f"browser-app: {' '.join(cmd)}")
    proc = subprocess.Popen(cmd)  # noqa: S603
    proc.wait()
    return True


def _start_webview(url: str) -> None:
    import webview

    start_kwargs: dict = {}
    if sys.platform == "win32":
        # Tránh WinForms/pythonnet (thường lỗi khi đóng gói + tải ZIP).
        start_kwargs["gui"] = "edgechromium"

    window = webview.create_window(
        "VT Rate Mail Sender",
        url,
        width=1320,
        height=860,
        min_size=(1050, 700),
    )

    def _on_closed() -> None:
        _force_exit(0)

    try:
        window.events.closed += _on_closed
    except Exception:  # noqa: BLE001
        pass

    webview.start(**start_kwargs)
    _force_exit(0)


def main() -> None:
    _fix_stdio_encoding()
    from app.webapp import flask_app

    port = _free_port()
    url = f"http://127.0.0.1:{port}/"

    def run_server() -> None:
        flask_app.run(host="127.0.0.1", port=port, threaded=True, use_reloader=False)

    threading.Thread(target=run_server, daemon=True).start()
    time.sleep(0.6)

    # 1) pywebview (Mac ổn; Windows cần WebView2 runtime)
    try:
        _start_webview(url)
        return
    except Exception as exc:  # noqa: BLE001
        detail = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        _log_startup(f"webview failed:\n{detail}")
        _safe_print(f"pywebview failed: {exc}")

    # 2) Windows: Edge/Chrome --app (cửa sổ giống desktop, không cần WebView2 COM)
    if sys.platform == "win32":
        try:
            if _open_browser_app_window(url):
                _force_exit(0)
                return
        except Exception as exc:  # noqa: BLE001
            _log_startup(f"browser-app failed: {exc}")
            _safe_print(f"browser-app failed: {exc}")

    # 3) Fallback cuối: trình duyệt thường
    webbrowser.open(url)
    _win_message(
        "VT Rate Mail Sender",
        "Khong mo duoc cua so desktop (WebView2).\n"
        "Da mo trinh duyet tam thoi.\n\n"
        "Cai Microsoft Edge WebView2 Runtime neu muon cua so app:\n"
        "https://developer.microsoft.com/microsoft-edge/webview2/\n\n"
        "Tat app: Task Manager → VTRateMailSender.exe\n"
        "hoac chay stop-app-windows.bat\n\n"
        "Log: %APPDATA%\\VTRateMailSender\\logs\\startup.log",
    )
    if getattr(sys, "frozen", False):
        # Giữ server tới khi user tắt process (để dùng được tạm trong browser)
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass
        _force_exit(0)
        return

    _safe_print(f"Server: {url}")
    _safe_print("Nhan Ctrl+C de dung.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
