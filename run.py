#!/usr/bin/env python3
"""VT Rate Mail Sender — Rich editor UI (TinyMCE) trong cửa sổ app."""

from __future__ import annotations

import socket
import threading
import time
import webbrowser


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return int(s.getsockname()[1])


def main() -> None:
    from app.webapp import flask_app

    port = _free_port()
    url = f"http://127.0.0.1:{port}/"

    def run_server() -> None:
        flask_app.run(host="127.0.0.1", port=port, threaded=True, use_reloader=False)

    threading.Thread(target=run_server, daemon=True).start()
    time.sleep(0.6)

    try:
        import webview

        webview.create_window(
            "VT Rate Mail Sender",
            url,
            width=1320,
            height=860,
            min_size=(1050, 700),
        )
        webview.start()
    except Exception as exc:  # noqa: BLE001
        print(f"Không mở được pywebview ({exc}). Mở trình duyệt…")
        webbrowser.open(url)
        print(f"Server: {url}")
        print("Nhấn Ctrl+C để dừng.")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass


if __name__ == "__main__":
    main()
