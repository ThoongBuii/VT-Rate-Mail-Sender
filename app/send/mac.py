"""macOS Legacy Outlook — mở New Mail (chữ ký sẵn) → điền To/Cc/Subject/Body → Send."""

from __future__ import annotations

import re
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

from ..models import AppConfig
from ..outlook_html import prepare_body_html_for_outlook
from .html_merge import html_visible_text


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
            "message": "macOS: New Mail Legacy tự có chữ ký — app điền To/Cc/Subject/Body rồi Send.",
        }

    def capture_signature(self) -> str:
        return (
            "macOS không cần chụp chữ ký riêng — dùng New Mail Legacy "
            "(chữ ký sẵn) khi Semi-Auto."
        )

    def send(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Path],
    ) -> None:
        """
        1 New Mail Legacy:
        - set account + To/Cc/Subject (+ gắn signature property nếu có)
        - open → chờ chữ ký hiện trong content
        - chèn body phía trên chữ ký (set content = body & sig)
        - send
        """
        prepared = prepare_body_html_for_outlook(body_html)
        if not (prepared or "").strip():
            raise ValueError("Nội dung mail trống — không gửi.")
        if not to_list:
            raise ValueError("Thiếu địa chỉ To — không gửi.")

        probe_raw = re.sub(r"\s+", " ", html_visible_text(prepared)).strip()
        probe = (probe_raw[:28] if len(probe_raw) >= 4 else "Dear").replace('"', "").replace(
            "\\", ""
        )

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
            script = f'''
set bodyPath to "{body_path}"
set bodyText to do shell script "cat " & quoted form of bodyPath
if bodyText is "" then error "Body HTML trống"

tell application "Microsoft Outlook"
  activate

  -- (1) Một New Mail
  set msg to make new outgoing message
{_account_block(self.config.from_email)}

  -- (2) Meta trước khi mở
  set subject of msg to "{_esc(subject)}"
{to_block}
{cc_block}
{att_block}

  -- Gắn chữ ký account nếu Outlook hỗ trợ
  try
    set accObj to account of msg
    set sigs to signatures of accObj
    if (count of sigs) > 0 then set signature of msg to item 1 of sigs
  end try
  try
    if (count of signatures) > 0 then set signature of msg to item 1 of signatures
  end try

  -- (3) Mở → Legacy render chữ ký vào content (giống New Mail tay)
  open msg
  delay 1.2

  set sigText to ""
  repeat with i from 1 to 24
    try
      set sigText to content of msg
    end try
    if (length of sigText) > 40 then exit repeat
    delay 0.25
  end repeat

  if (length of sigText) ≤ 40 then
    try
      if (count of signatures) > 0 then set signature of msg to item 1 of signatures
    end try
    delay 1.0
    try
      set sigText to content of msg
    end try
  end if

  -- (4) Body phía trên chữ ký — 1 cửa sổ, không tạo mail 2
  try
    set content of msg to bodyText & sigText
  on error
    try
      set content of msg to bodyText & return & sigText
    on error
      set content of msg to bodyText
    end try
  end try

  delay 0.45

  set finalContent to ""
  try
    set finalContent to content of msg
  end try
  if finalContent is "" then
    try
      set finalContent to plain text content of msg
    end try
  end if
  if finalContent does not contain "{_esc(probe)}" then
    error "Body chưa vào New Mail — hủy Send. Đóng Untitled, Legacy ON, thử 1 mail."
  end if

  set toCount to 0
  try
    set toCount to count of (to recipients of msg)
  end try
  if toCount < 1 then error "To trống — hủy gửi"

  -- (5) Gửi
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
                    "- Legacy Outlook ON\n"
                    "- Account overseas login, New Mail tay vẫn có chữ ký\n"
                    "- Xóa Outbox + đóng Untitled rồi thử 1 mail"
                ) from exc
