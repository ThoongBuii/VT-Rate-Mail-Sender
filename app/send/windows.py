"""Windows Outlook send via COM — độc lập hoàn toàn với macOS.

Luồng (cấu trúc cũ, ổn định):
1) CreateItem + Display = New Mail (Outlook tự gắn chữ ký + ảnh)
2) Điền To / Cc / Subject chính xác qua COM
3) Merge body HTML phía trên chữ ký (giữ CID ảnh chữ ký)
4) Ảnh trong body (data:image) → attachment CID ẩn (Outlook mới hiện được)
5) Đính file (nhiều file) → Send
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any, Optional, Sequence, Union

from ..models import AppConfig
from ..outlook_html import prepare_body_html_for_outlook
from .html_merge import merge_body_with_outlook_signature
from .inline_images import (
    InlineCidImage,
    attach_cid_images,
    cleanup_temp_files,
    replace_data_images_with_cid,
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
        """Xóa đoạn trống đầu New Mail (thường ~2 dòng trước chữ ký)."""
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
        """Xóa đoạn trống ngay trước 'Best Regards' / chữ ký."""
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
    ) -> list[InlineCidImage]:
        """
        Chèn nội dung HTML vào thân mail, giữ chữ ký Outlook (ảnh CID còn).
        Ảnh data:image trong body → CID ẩn (Outlook gửi mới hiện được).
        """
        word_doc = None
        try:
            insp = mail_item.GetInspector
            word_doc = insp.WordEditor
        except Exception:  # noqa: BLE001
            word_doc = None

        if word_doc is not None:
            self._windows_delete_leading_empty_paragraphs(word_doc)

        # 1) data:image → cid: trong HTML (không đụng chữ ký)
        body_ready, inline_items = replace_data_images_with_cid(body_html)
        existing = str(getattr(mail_item, "HTMLBody", None) or "")
        # 2) set HTMLBody đã merge
        mail_item.HTMLBody = merge_body_with_outlook_signature(body_ready, existing)
        # 3) đính ảnh CID sau HTMLBody (Outlook mới map đúng)
        attach_cid_images(mail_item, inline_items)

        try:
            for i in range(int(mail_item.Attachments.Count), 0, -1):
                att = mail_item.Attachments.Item(i)
                name = str(getattr(att, "FileName", "") or "")
                if name.lower() in {"vt_body.htm", "vt_body.html"}:
                    att.Delete()
        except Exception:  # noqa: BLE001
            pass

        try:
            word_doc = mail_item.GetInspector.WordEditor
            if word_doc is not None:
                self._windows_trim_gap_before_signature(word_doc)
        except Exception:  # noqa: BLE001
            pass
        return inline_items

    def send(
        self,
        to_list: list[str],
        cc_list: list[str],
        subject: str,
        body_html: str,
        attachment: Optional[Union[Path, Sequence[Path]]] = None,
    ) -> None:
        """New Mail → điền To/Cc/Subject/Body chính xác → Send (chữ ký Outlook giữ ảnh)."""
        import win32com.client  # type: ignore

        if not to_list:
            raise ValueError("Thiếu địa chỉ To — không gửi.")

        prepared = prepare_body_html_for_outlook(body_html)
        if not (prepared or "").strip():
            raise ValueError("Nội dung mail trống — không gửi.")

        files = self._normalize_attachments(attachment)
        outlook = win32com.client.Dispatch("Outlook.Application")
        account = self._windows_pick_account(outlook)
        mail_item = outlook.CreateItem(0)  # olMailItem
        inline_items: list[InlineCidImage] = []

        try:
            if account is not None:
                try:
                    mail_item.SendUsingAccount = account
                except Exception:  # noqa: BLE001
                    pass

            # Mở New Mail để Outlook gắn chữ ký mặc định (giống click New Email)
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
                time.sleep(0.45)
            except Exception:  # noqa: BLE001
                pass

            # Điền các ô header chính xác trước (To / Cc / Subject)
            mail_item.To = "; ".join(to_list)
            mail_item.CC = "; ".join(cc_list) if cc_list else ""
            mail_item.Subject = subject

            # Body phía trên chữ ký + ảnh inline → CID
            inline_items = self._windows_apply_body_keep_signature(mail_item, prepared)

            for path in files:
                mail_item.Attachments.Add(str(path.resolve()))

            mail_item.Send()
        finally:
            # Cho Outlook kịp đọc file CID trước khi xóa temp
            try:
                time.sleep(0.35)
            except Exception:  # noqa: BLE001
                pass
            cleanup_temp_files(inline_items)
