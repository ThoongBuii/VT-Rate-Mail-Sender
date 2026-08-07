#!/usr/bin/env python3
"""VT Rate Mail Sender — Rich editor UI (TinyMCE) trong cửa sổ app."""

from __future__ import annotations

import atexit
import os
import sys
import socket
import threading
import time
import webbrowser


def _fix_stdio_encoding() -> None:
    """Windows console thường cp1252 — tránh crash khi in tiếng Việt."""
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
    """Flask/WebView2 đôi khi giữ process sống sau khi đóng cửa sổ → khóa file trên Windows."""
    try:
        os._exit(code)
    except Exception:  # noqa: BLE001
        sys.exit(code)


def _win_message(title: str, text: str) -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(0, text, title, 0x10)
    except Exception:  # noqa: BLE001
        pass


def main() -> None:
    _fix_stdio_encoding()
    # Đảm bảo process thoát khi cửa sổ đóng (kể cả thread Flask còn treo)
    atexit.register(lambda: None)

    from app.webapp import flask_app

    port = _free_port()
    url = f"http://127.0.0.1:{port}/"

    def run_server() -> None:
        flask_app.run(host="127.0.0.1", port=port, threaded=True, use_reloader=False)

    threading.Thread(target=run_server, daemon=True).start()
    time.sleep(0.6)

    try:
        import webview

        # WinForms/pythonnet hay lỗi với PyInstaller + ZIP tải từ web.
        # Edge WebView2 ổn định hơn trên Windows 10/11 (có Microsoft Edge).
        start_kwargs: dict = {}
        if sys.platform == "win32":
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
    except Exception as exc:  # noqa: BLE001
        _safe_print(f"Khong mo duoc pywebview ({exc}). Mo trinh duyet...")
        webbrowser.open(url)
        _win_message(
            "VT Rate Mail Sender",
            "Khong mo duoc cua so desktop.\n"
            "Da mo trinh duyet tam thoi.\n\n"
            "De tat app (moi xoa/ghi de duoc file):\n"
            "Task Manager → End task → VTRateMailSender.exe\n"
            "hoac chay stop-app-windows.bat",
        )
        # Không giữ process vô hạn khi đóng gói — tránh khóa thư mục cài đặt
        if getattr(sys, "frozen", False):
            time.sleep(2)
            _force_exit(1)
        _safe_print(f"Server: {url}")
        _safe_print("Nhan Ctrl+C de dung.")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass


if __name__ == "__main__":
    main()
