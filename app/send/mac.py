"""macOS Legacy Outlook — hiệu quả: New Mail + chữ ký (CID→base64) + body → Send.
Không automation UI kiểu Windows (Tab/Cmd+V).
"""

from __future__ import annotations

import base64
import mimetypes
import re
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

from ..models import AppConfig
from ..outlook_html import prepare_body_html_for_outlook
from .html_merge import html_visible_text, merge_body_with_outlook_signature


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


def _embed_cid_images(html: str, att_dir: Path) -> str:
    """Thay cid:... bằng data: URI từ file đã export — giữ logo chữ ký khi set content."""
    if not html or not att_dir.is_dir():
        return html or ""

    files: dict[str, Path] = {}
    for f in att_dir.iterdir():
        if f.is_file():
            files[f.name] = f
            files[f.name.lower()] = f
            files[f.stem] = f
            files[f.stem.lower()] = f

    def _lookup(cid: str) -> Optional[Path]:
        name = (cid or "").strip().strip("<>").split("@")[0].strip()
        if not name:
            return None
        for key in (
            name,
            name.lower(),
            Path(name).name,
            Path(name).stem,
            Path(name).stem.lower(),
        ):
            if key in files:
                return files[key]
        token = re.sub(r"[^a-zA-Z0-9_]+", "", name).lower()
        if len(token) >= 6:
            for f in att_dir.iterdir():
                if not f.is_file():
                    continue
                compact = f.name.lower().replace("-", "").replace(".", "")
                if token in compact:
                    return f
        return None

    def _repl(m: re.Match[str]) -> str:
        quote = m.group(1)
        path = _lookup(m.group(2))
        if path is None:
            return m.group(0)
        raw = path.read_bytes()
        mime = mimetypes.guess_type(path.name)[0] or "image/png"
        b64 = base64.b64encode(raw).decode("ascii")
        return f"src={quote}data:{mime};base64,{b64}{quote}"

    return re.sub(r"""(?is)src\s*=\s*(['"])cid:([^'"]+)\1""", _repl, html)


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
            "message": "macOS: New Mail + chữ ký (ảnh nhúng) → điền To/Cc/Subject/Body → Send.",
        }

    def capture_signature(self) -> str:
        return "macOS không cần chụp chữ ký riêng — Semi-Auto tự lấy từ New Mail Legacy."

    def send(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Path],
    ) -> None:
        prepared = prepare_body_html_for_outlook(body_html)
        if not (prepared or "").strip():
            raise ValueError("Nội dung mail trống — không gửi.")
        if not to_list:
            raise ValueError("Thiếu địa chỉ To — không gửi.")

        # Probe ngắn — tránh fail vì &nbsp; / xuống dòng trong HTML
        probe_raw = re.sub(r"\s+", " ", html_visible_text(prepared)).strip()
        probe = "Dear"
        for token in re.findall(r"[A-Za-zÀ-ỹ]{3,}", probe_raw or ""):
            probe = token[:12]
            break

        to_block = "\n".join(
            f'  make new to recipient at msg with properties {{email address:{{address:"{_esc(t)}"}}}}'
            for t in to_list
        )
        cc_block = "\n".join(
            f'  make new cc recipient at msg with properties {{email address:{{address:"{_esc(cc)}"}}}}'
            for cc in cc_list
        )

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            body_path = tmp_path / "body.html"
            sig_path = tmp_path / "sig.html"
            full_path = tmp_path / "full.html"
            att_dir = tmp_path / "sig_atts"
            att_dir.mkdir(parents=True, exist_ok=True)
            body_path.write_text(prepared, encoding="utf-8")

            # —— Phase 1: New Mail + chữ ký → ghi sig.html + export ảnh ——
            # (chưa đính file rate — tránh lẫn với ảnh chữ ký)
            phase1 = f'''
on lowerCase(t)
  set s to t as string
  set lowers to "abcdefghijklmnopqrstuvwxyz"
  set uppers to "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  set out to ""
  repeat with c in characters of s
    set ch to c as string
    set x to offset of ch in uppers
    if x > 0 then
      set out to out & character x of lowers
    else
      set out to out & ch
    end if
  end repeat
  return out
end lowerCase

set sigPath to "{sig_path}"
set attDir to "{att_dir}"
do shell script "mkdir -p " & quoted form of attDir

tell application "Microsoft Outlook"
  activate
  set msg to make new outgoing message
{_account_block(self.config.from_email)}
  set subject of msg to "{_esc(subject)}"
{to_block}
{cc_block}

  try
    set accObj to account of msg
    set sigs to signatures of accObj
    if (count of sigs) > 0 then set signature of msg to item 1 of sigs
  end try
  try
    if (count of signatures) > 0 then set signature of msg to item 1 of signatures
  end try

  open msg
  delay 1.5

  set sigText to ""
  repeat with i from 1 to 24
    try
      set sigText to content of msg
    end try
    if (length of sigText) > 40 then exit repeat
    delay 0.25
  end repeat

  try
    set nAtt to count of attachments of msg
    repeat with i from 1 to nAtt
      try
        set a to attachment i of msg
        set aname to name of a as string
        set low to my lowerCase(aname)
        if low ends with ".png" or low ends with ".jpg" or low ends with ".jpeg" or low ends with ".gif" or low ends with ".webp" or low ends with ".tif" or low ends with ".tiff" then
          set dest to attDir & "/" & aname
          try
            save a in (POSIX file dest)
          end try
        end if
      end try
    end repeat
  end try

  set mid to ""
  try
    set mid to (id of msg) as string
  end try
end tell

set fRef to open for access (POSIX file sigPath) with write permission
set eof of fRef to 0
write sigText to fRef as «class utf8»
close access fRef

return mid
'''
            mid = _run_osascript(phase1, timeout=180)

            sig_html = ""
            if sig_path.is_file():
                sig_html = sig_path.read_text(encoding="utf-8", errors="ignore")
            sig_html = _embed_cid_images(sig_html, att_dir)
            full_html = merge_body_with_outlook_signature(prepared, sig_html)
            if len(html_visible_text(full_html)) < 8:
                raise RuntimeError(
                    "Không lấy được chữ ký/body từ New Mail.\n"
                    "Kiểm tra Legacy Outlook ON và New Mail tay vẫn có chữ ký."
                )
            full_path.write_text(full_html, encoding="utf-8")

            att_block = ""
            if attachment:
                att = _esc(str(attachment.resolve()))
                att_block = (
                    f'  make new attachment at msg with properties '
                    f'{{file:POSIX file "{att}"}}\n'
                )

            # —— Phase 2: set content đã embed ảnh + attachment rate + Send ——
            find_msg = ""
            if mid:
                find_msg = f'''
  set msg to missing value
  try
    set msg to message id "{_esc(mid)}"
  end try
  if msg is missing value then
    try
      set msg to first outgoing message whose subject is "{_esc(subject)}"
    end try
  end if
'''
            else:
                find_msg = f'''
  set msg to first outgoing message whose subject is "{_esc(subject)}"
'''

            phase2 = f'''
set fullPath to "{full_path}"
set fullText to do shell script "cat " & quoted form of fullPath
if fullText is "" then error "HTML full trống"

tell application "Microsoft Outlook"
  activate
{find_msg}
  if msg is missing value then error "Không tìm lại được New Mail"

  set content of msg to fullText
  delay 0.5

{att_block}

  -- Verify mềm: plain text hoặc content có probe ngắn
  set ok to false
  set checkText to ""
  try
    set checkText to plain text content of msg
  end try
  if checkText is "" then
    try
      set checkText to content of msg
    end try
  end if
  if checkText contains "{_esc(probe)}" then set ok to true
  if checkText contains "Dear" then set ok to true
  if checkText contains "dear" then set ok to true
  if ok is false then
    -- Vẫn gửi nếu content đủ dài (draft đã đúng trước đây — verify cũ quá chặt)
    if (length of checkText) < 80 then
      error "Body quá ngắn sau khi ghi — hủy Send"
    end if
  end if

  set toCount to 0
  try
    set toCount to count of (to recipients of msg)
  end try
  if toCount < 1 then error "To trống — hủy gửi"

  send msg
end tell
'''
            try:
                _run_osascript(phase2, timeout=180)
            except RuntimeError as exc:
                raise RuntimeError(
                    "Outlook Mac gửi thất bại.\n"
                    f"{exc}\n"
                    "Gợi ý: Legacy Outlook ON · đóng Untitled/Draft lỗi · thử 1 mail."
                ) from exc
