#!/usr/bin/env python3
"""VT Rate Mail Sender — desktop window (pywebview / Edge WebView2)."""

from __future__ import annotations

import os
import socket
import sys
import threading
import time
import traceback
import webbrowser


APP_USER_MODEL_ID = "VTLogistics.VTRateMailSender"


def _fix_stdio_encoding() -> None:
    if sys.platform != "win32":
        return
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except Exception:  # noqa: BLE001
            pass


def _set_windows_app_id() -> None:
    """Giúp Windows coi đây là app riêng (pin taskbar / icon), không gộp với browser."""
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
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


def _start_webview(url: str) -> None:
    # Ép Edge WebView2 — không dùng WinForms/pythonnet, không mở Chrome.
    if sys.platform == "win32":
        os.environ.setdefault("PYWEBVIEW_GUI", "edgechromium")

    import webview

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


def main() -> None:
    _fix_stdio_encoding()
    _set_windows_app_id()

    from app.webapp import flask_app

    port = _free_port()
    url = f"http://127.0.0.1:{port}/"

    def run_server() -> None:
        flask_app.run(host="127.0.0.1", port=port, threaded=True, use_reloader=False)

    threading.Thread(target=run_server, daemon=True).start()
    time.sleep(0.6)

    try:
        _start_webview(url)
        return
    except Exception as exc:  # noqa: BLE001
        detail = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        _log_startup(f"webview failed:\n{detail}")
        _safe_print(f"pywebview failed: {exc}")

    # Không fallback Chrome/website — đó là lý do pin ra icon Chrome.
    msg = (
        "Khong mo duoc cua so desktop (Microsoft Edge WebView2).\n\n"
        "Day la ung dung desktop, khong chay bang Chrome.\n"
        "Hay cai WebView2 Runtime (mien phi) rồi mo lai app:\n"
        "https://go.microsoft.com/fwlink/p/?LinkId=2124703\n\n"
        "Log: %APPDATA%\\VTRateMailSender\\logs\\startup.log"
    )
    _win_message("VT Rate Mail Sender", msg)
    try:
        webbrowser.open("https://go.microsoft.com/fwlink/p/?LinkId=2124703")
    except Exception:  # noqa: BLE001
        pass
    _force_exit(1)


if __name__ == "__main__":
    main()
