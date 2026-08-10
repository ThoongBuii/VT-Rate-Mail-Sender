"""macOS Legacy Outlook — tối giản, hiệu quả.

Không rewrite chữ ký (tránh gãy ảnh CID).
Không UI paste kiểu Windows.
Luồng: New Mail → To/Cc/Subject → content=body → gắn signature Outlook → Send.
"""

from __future__ import annotations

import subprocess
import tempfile
from pathlib import Path
from typing import Optional

from ..models import AppConfig
from ..outlook_html import prepare_body_html_for_outlook


def _esc(s: str) -> str:
    return (s or "").replace("\\", "\\\\").replace('"', '\\"')


def _account_block(from_email: str) -> str:
    if not from_email:
        return ""
    fe = _esc(from_email)
    return (
        "\n"
        "  try\n"
        "    repeat with acc in (get exchange accounts)\n"
        "      try\n"
        f'        if (email address of acc as string) contains "{fe}" then\n'
        "          set account of msg to acc\n"
        "          exit repeat\n"
        "        end if\n"
        "      end try\n"
        "    end repeat\n"
        "  end try\n"
        "  try\n"
        "    repeat with acc in (get imap accounts)\n"
        "      try\n"
        f'        if (email address of acc as string) contains "{fe}" then\n'
        "          set account of msg to acc\n"
        "          exit repeat\n"
        "        end if\n"
        "      end try\n"
        "    end repeat\n"
        "  end try\n"
    )


def _run_osascript(script: str, timeout: int = 180) -> str:
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "vt_mac.applescript"
        path.write_text(script, encoding="utf-8")
        result = subprocess.run(
            ["osascript", str(path)],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    if result.returncode != 0:
        raise RuntimeError((result.stderr or result.stdout or "osascript failed").strip())
    return (result.stdout or "").strip()


class MacOutlookSender:
    def __init__(self, config: AppConfig):
        self.config = config

    def open_outlook(self) -> str:
        subprocess.run(["open", "-a", "Microsoft Outlook"], check=False)
        result = subprocess.run(
            ["osascript", "-l", "JavaScript", "-e", 'Application("Microsoft Outlook").name();'],
            capture_output=True,
            text=True,
            timeout=60,
        )
        if result.returncode != 0:
            err = (result.stderr or result.stdout or "").strip()
            raise RuntimeError(
                "Không điều khiển được Microsoft Outlook trên Mac.\n"
                "Cần Legacy Outlook ON + đã đăng nhập.\n"
                f"Chi tiết: {err or 'osascript failed'}"
            )
        return self.config.from_email or "Outlook (Mac)"

    def signature_status(self) -> dict:
        return {
            "ready": True,
            "path": "",
            "message": (
                "macOS: New Mail → điền To/Cc/Subject/Body → gắn chữ ký Outlook → Send "
                "(không sửa HTML chữ ký)."
            ),
        }

    def capture_signature(self) -> str:
        return "Không cần chụp chữ ký — Outlook tự gắn khi gửi."

    def send(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Path],
    ) -> None:
        """
        1 cửa sổ / 1 lần gửi, tối thiểu chỉnh sửa:
        body HTML do app; chữ ký do Outlook gắn (giữ ảnh).
        """
        prepared = prepare_body_html_for_outlook(body_html)
        if not (prepared or "").strip():
            raise ValueError("Nội dung mail trống — không gửi.")
        if not to_list:
            raise ValueError("Thiếu địa chỉ To — không gửi.")

        to_block = "\n".join(
            f'  make new to recipient at msg with properties {{email address:{{address:"{_esc(t)}"}}}}'
            for t in to_list
        )
        cc_block = "\n".join(
            f'  make new cc recipient at msg with properties {{email address:{{address:"{_esc(cc)}"}}}}'
            for cc in cc_list
        )
        if attachment:
            att = _esc(str(attachment.resolve()))
            att_block = (
                f'  make new attachment at msg with properties {{file:POSIX file "{att}"}}'
            )
        else:
            att_block = ""

        with tempfile.TemporaryDirectory() as tmp:
            body_path = Path(tmp) / "body.html"
            body_path.write_text(prepared, encoding="utf-8")

            # Một script duy nhất — không phase 2, không merge/rewrite chữ ký.
            script = f'''
set bodyPath to "{body_path}"
set bodyText to do shell script "cat " & quoted form of bodyPath
if bodyText is "" then error "Body HTML trống"

tell application "Microsoft Outlook"
  activate

  set msg to make new outgoing message
{_account_block(self.config.from_email)}

  set subject of msg to "{_esc(subject)}"
{to_block}
{cc_block}

  -- Body trước (app), chữ ký sau (Outlook giữ ảnh)
  set content of msg to bodyText

  try
    set accObj to account of msg
    set sigs to signatures of accObj
    if (count of sigs) > 0 then set signature of msg to item 1 of sigs
  end try
  try
    if (count of signatures) > 0 then set signature of msg to item 1 of signatures
  end try

{att_block}

  set toCount to 0
  try
    set toCount to count of (to recipients of msg)
  end try
  if toCount < 1 then error "To trống — hủy gửi"

  -- Gửi thẳng, không mở cửa sổ soạn (tránh sót Untitled + không rewrite HTML)
  send msg
end tell
'''
            try:
                _run_osascript(script, timeout=180)
            except RuntimeError as exc:
                raise RuntimeError(
                    "Outlook Mac gửi thất bại.\n"
                    f"{exc}\n"
                    "Gợi ý:\n"
                    "- Bật Legacy Outlook = ON (góc phải Outlook)\n"
                    "- Account overseas đã login, chữ ký New Mail tay OK\n"
                    "- Xóa Outbox + đóng hết Untitled rồi thử 1 mail"
                ) from exc
