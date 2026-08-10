"""Windows Outlook send via COM — độc lập macOS."""
from __future__ import annotations

import time
from pathlib import Path
from typing import Any, Optional

from ..models import AppConfig
from ..outlook_html import prepare_body_html_for_outlook
from .html_merge import merge_body_with_outlook_signature

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

        """Xóa đoạn trống ngay trước 'Best Regards' / chữ ký (gap ~2 dòng)."""

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

    def _windows_apply_body_keep_signature(self, mail_item: Any, body_html: str) -> None:

        """

        Chèn nội dung HTML vào thân mail (không đính file .htm),

        giữ chữ ký Outlook + cắt khoảng trống đầu / trước Best Regards.

        """

        word_doc = None

        try:

            insp = mail_item.GetInspector

            word_doc = insp.WordEditor

        except Exception:  # noqa: BLE001

            word_doc = None

        if word_doc is not None:

            self._windows_delete_leading_empty_paragraphs(word_doc)

        existing = str(getattr(mail_item, "HTMLBody", None) or "")

        mail_item.HTMLBody = merge_body_with_outlook_signature(body_html, existing)

        # Cắt gap sau khi merge; đồng thời bỏ mọi attachment tạm vt_body.htm (nếu còn)

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

    def send(

        self,

        to_list: list[str],

        cc_list: list[str],

        subject: str,

        body_html: str,

        attachment: Optional[Path],

    ) -> None:

        import win32com.client  # type: ignore

        from .outlook_html import prepare_body_html_for_outlook

        outlook = win32com.client.Dispatch("Outlook.Application")

        account = self._windows_pick_account(outlook)

        prepared = prepare_body_html_for_outlook(body_html)

        mail_item = outlook.CreateItem(0)

        if account is not None:

            try:

                mail_item.SendUsingAccount = account

            except Exception:  # noqa: BLE001

                pass

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

        self._windows_apply_body_keep_signature(mail_item, prepared)

        mail_item.To = "; ".join(to_list)

        if cc_list:

            mail_item.CC = "; ".join(cc_list)

        mail_item.Subject = subject

        if attachment:

            mail_item.Attachments.Add(str(attachment.resolve()))

        mail_item.Send()
