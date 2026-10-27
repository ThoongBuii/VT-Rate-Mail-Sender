"""Windows Outlook send via COM — độc lập hoàn toàn với macOS.

Luồng gửi (giữ nguyên nội dung/ảnh/chữ ký):
1) CreateItem + Display = New Mail (chữ ký Outlook)
2) A–D: ẩn Inspector autosend + trả focus về cửa sổ đang làm
3) To / Cc / Subject → InsertFile body+ảnh → BreakLink
4) Disable input cửa sổ autosend → Save → Send

Không đổi đường gửi mail — chỉ tách UI để không cướp chat/soạn tay.
"""

from __future__ import annotations

import re
import time
from pathlib import Path
from typing import Any, Optional, Sequence, Union

from ..models import AppConfig
from ..outlook_html import prepare_body_html_for_outlook
from .inline_images import (
    InlinePicture,
    cleanup_temp_files,
    clear_windows_clipboard,
    insert_body_html_via_word_file,
)
from .win_isolate import (
    IsolateState,
    get_foreground_hwnd,
    isolate_after_display,
    lock_input_after_content,
    reassert_user_focus,
    unlock_before_send,
)


class WindowsOutlookSender:
    def __init__(self, config: AppConfig):
        self.config = config

    def _windows_pick_account(self, outlook: Any):
        if not self.config.from_email:
            return None
        try:
            namespace = outlook.GetNamespace("MAPI")
            for i in range(1, namespace.Accounts.Count + 1):
                acc = namespace.Accounts.Item(i)
                smtp = getattr(acc, "SmtpAddress", "") or ""
                if smtp.lower() == self.config.from_email.lower():
                    return acc
        except Exception:  # noqa: BLE001
            return None
        return None

    def _windows_delete_leading_empty_paragraphs(self, word_doc: Any, limit: int = 30) -> int:
        removed = 0
        for _ in range(limit):
            try:
                if int(word_doc.Paragraphs.Count) < 1:
                    break
                para = word_doc.Paragraphs(1)
                raw = str(para.Range.Text or "")
                clean = (
                    raw.replace("\r", "")
                    .replace("\x07", "")
                    .replace("\xa0", " ")
                    .replace("\u200b", "")
                    .strip()
                )
                if clean:
                    break
                para.Range.Delete()
                removed += 1
            except Exception:  # noqa: BLE001
                break
        return removed

    def _windows_trim_gap_before_signature(self, word_doc: Any) -> int:
        import re

        removed = 0
        try:
            count = int(word_doc.Paragraphs.Count)
        except Exception:  # noqa: BLE001
            return 0

        sig_idx = None
        for i in range(1, count + 1):
            try:
                raw = str(word_doc.Paragraphs(i).Range.Text or "")
            except Exception:  # noqa: BLE001
                continue
            if re.search(r"best\s*regards", raw, flags=re.I):
                sig_idx = i
                break

        if not sig_idx or sig_idx <= 1:
            return 0

        while sig_idx > 1 and removed < 20:
            try:
                prev = word_doc.Paragraphs(sig_idx - 1)
                raw = str(prev.Range.Text or "")
                clean = (
                    raw.replace("\r", "")
                    .replace("\x07", "")
                    .replace("\xa0", " ")
                    .replace("\u200b", "")
                    .strip()
                )
                if clean:
                    break
                prev.Range.Delete()
                removed += 1
                sig_idx -= 1
            except Exception:  # noqa: BLE001
                break
        return removed

    def _normalize_attachments(
        self, attachment: Optional[Union[Path, Sequence[Path]]]
    ) -> list[Path]:
        if attachment is None:
            return []
        if isinstance(attachment, Path):
            return [attachment]
        return [p for p in attachment if p is not None]

    def _windows_apply_body_keep_signature(
        self, mail_item: Any, body_html: str
    ) -> list[InlinePicture]:
        """Chèn body bằng InsertFile — giữ nguyên chữ ký WordEditor."""
        word_doc = None
        try:
            word_doc = mail_item.GetInspector.WordEditor
        except Exception:  # noqa: BLE001
            word_doc = None

        if word_doc is None:
            try:
                mail_item.HTMLBody = body_html + (str(mail_item.HTMLBody or ""))
            except Exception:  # noqa: BLE001
                pass
            return []

        self._windows_delete_leading_empty_paragraphs(word_doc)

        clear_windows_clipboard()
        pictures: list[InlinePicture] = []
        try:
            time.sleep(0.15)
            pictures, _htm = insert_body_html_via_word_file(word_doc, body_html)
            time.sleep(0.2)
            self._windows_trim_gap_before_signature(word_doc)
        except Exception:  # noqa: BLE001
            pass
        clear_windows_clipboard()

        try:
            for i in range(int(mail_item.Attachments.Count), 0, -1):
                att = mail_item.Attachments.Item(i)
                name = str(getattr(att, "FileName", "") or "").lower()
                if name.endswith(".htm") or name.endswith(".html") or name.startswith("vt_body"):
                    if name.startswith("vt_") or "vt_body" in name:
                        att.Delete()
        except Exception:  # noqa: BLE001
            pass

        return pictures

    def send(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Union[Path, Sequence[Path]]] = None,
    ) -> None:
        """New Mail (ẩn UI) → InsertFile body+ảnh → Send — không cướp focus user."""
        import win32com.client  # type: ignore

        from ..sender_smtp import extract_email, is_valid_email

        def _clean(items: list[str]) -> list[str]:
            out: list[str] = []
            seen: set[str] = set()
            for raw in items or []:
                em = extract_email(raw) or (raw or "").strip()
                if not is_valid_email(em):
                    continue
                key = em.lower()
                if key in seen:
                    continue
                seen.add(key)
                out.append(em)
            return out

        to_list = _clean(to_list)
        to_keys = {e.lower() for e in to_list}
        cc_list = [e for e in _clean(list(cc_list or [])) if e.lower() not in to_keys]
        if not to_list:
            raise ValueError("Thiếu / sai địa chỉ To — không gửi.")

        prepared = prepare_body_html_for_outlook(body_html)
        if not (prepared or "").strip():
            raise ValueError("Nội dung mail trống — không gửi.")

        files = self._normalize_attachments(attachment)
        outlook = win32com.client.Dispatch("Outlook.Application")
        account = self._windows_pick_account(outlook)
        mail_item = outlook.CreateItem(0)
        pictures: list[InlinePicture] = []
        isolate = IsolateState()

        try:
            if account is not None:
                try:
                    mail_item.SendUsingAccount = account
                except Exception:  # noqa: BLE001
                    pass

            # Nhớ cửa sổ đang làm (chat / mail khác) TRƯỚC Display
            prev_hwnd = get_foreground_hwnd()

            try:
                insp = mail_item.GetInspector
                try:
                    _ = insp.WordEditor
                except Exception:  # noqa: BLE001
                    pass
            except Exception:  # noqa: BLE001
                pass

            try:
                mail_item.Display(False)
                # Ẩn nhanh — sleep ngắn hơn trước đây để giảm khe nháy
                time.sleep(0.12)
                isolate = isolate_after_display(
                    mail_item, prev_foreground=prev_hwnd, settle_ms=0.05
                )
            except Exception:  # noqa: BLE001
                isolate = IsolateState(prev_foreground=prev_hwnd)

            mail_item.To = "; ".join(to_list)
            mail_item.CC = "; ".join(cc_list) if cc_list else ""
            mail_item.Subject = subject
            # Xác minh To/Cc trước khi chèn body
            got_to = [
                extract_email(p) or p.strip()
                for p in re.split(r"[;,]+", str(mail_item.To or ""))
                if (extract_email(p) or p.strip())
            ]
            if len(got_to) < len(to_list):
                raise RuntimeError(
                    f"To chưa gắn đủ trong Outlook: {mail_item.To!r} (cần {to_list})"
                )
            reassert_user_focus(isolate)

            # WordEditor vẫn chạy trên Inspector ẩn (chưa Disable)
            pictures = self._windows_apply_body_keep_signature(mail_item, prepared)
            reassert_user_focus(isolate)

            for path in files:
                mail_item.Attachments.Add(str(path.resolve()))

            # Nội dung xong → khóa input cửa sổ autosend (lớp C)
            lock_input_after_content(isolate)

            try:
                mail_item.Save()
                time.sleep(0.25)
            except Exception:  # noqa: BLE001
                pass

            unlock_before_send(isolate)
            mail_item.Send()
            # Trả focus lần cuối sau Send (Inspector thường tự đóng)
            reassert_user_focus(isolate)
        finally:
            try:
                time.sleep(0.8)
            except Exception:  # noqa: BLE001
                pass
            cleanup_temp_files(pictures)
