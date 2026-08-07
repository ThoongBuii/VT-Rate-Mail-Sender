from __future__ import annotations

import json
import threading
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Optional

import customtkinter as ctk

from .importer import import_agencies
from .models import AgencyMail, AppConfig, MailStatus, SendProgress
from .queue_worker import SemiAutoQueue
from .sender import OutlookDesktopSender
from .template_engine import (
    SUGGESTED_SUBJECT,
    SUGGESTED_TEMPLATE_HTML,
    render_body_html,
    wrap_preview_document,
)

ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / "config.json"
LOG_DIR = ROOT / "logs"

STATUS_COLORS = {
    MailStatus.PENDING: "#8a8a8a",
    MailStatus.READY: "#1f6aa5",
    MailStatus.SENDING: "#c47a00",
    MailStatus.SENT: "#1a7f37",
    MailStatus.FAILED: "#c62828",
    MailStatus.SKIPPED: "#6d4c41",
}


class VTRateMailApp(ctk.CTk):
    """UI đồng nhất kiểu Outlook: danh bạ | soạn/preview HTML | Semi-Auto."""

    def __init__(self) -> None:
        super().__init__()
        self.title("VT Rate Mail Sender")
        self.geometry("1320x800")
        self.minsize(1120, 700)

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.config_data = self._load_config()
        self.sender = OutlookDesktopSender(self.config_data)
        self.queue = SemiAutoQueue(
            sender=self.sender,
            delay_min=self.config_data.delay_min_seconds,
            delay_max=self.config_data.delay_max_seconds,
            on_progress=self._on_progress,
        )
        self.mails: list[AgencyMail] = []
        self._selected_index: Optional[int] = None
        self._login_busy = False
        self._preview_mode = ctk.StringVar(value="origin")  # origin | preview
        self._html_frame = None

        self.subject_var = ctk.StringVar()
        self.attachment_var = ctk.StringVar()
        self.template_html = SUGGESTED_TEMPLATE_HTML

        self._build_ui()
        self._load_compose_defaults()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(200, self._refresh_auth_label)
        self.after(300, self._refresh_html_view)

    def _load_config(self) -> AppConfig:
        if CONFIG_PATH.exists():
            return AppConfig.from_dict(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
        example = ROOT / "config.example.json"
        if example.exists():
            return AppConfig.from_dict(json.loads(example.read_text(encoding="utf-8")))
        return AppConfig.from_dict({})

    def _save_config(self) -> None:
        self.config_data.delay_min_seconds = int(self.delay_min_var.get() or 10)
        self.config_data.delay_max_seconds = int(self.delay_max_var.get() or 20)
        self.config_data.default_subject = self.subject_var.get().strip()
        self.config_data.default_attachment = self.attachment_var.get().strip()
        self.config_data.default_template = self.template_html
        self.sender.config = self.config_data
        CONFIG_PATH.write_text(
            json.dumps(self.config_data.to_dict(), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    def _load_compose_defaults(self) -> None:
        self.subject_var.set(self.config_data.default_subject or SUGGESTED_SUBJECT)
        self.attachment_var.set(self.config_data.default_attachment)
        tpl = self.config_data.default_template or SUGGESTED_TEMPLATE_HTML
        self.template_html = tpl

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_header()

        body = ctk.CTkFrame(self, fg_color="#dce3ec", corner_radius=0)
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_columnconfigure(0, weight=2)
        body.grid_columnconfigure(1, weight=5)
        body.grid_rowconfigure(1, weight=1)

        self._build_toolbar(body)
        self._build_left(body)
        self._build_right(body)

    def _build_header(self) -> None:
        header = ctk.CTkFrame(self, fg_color="#0f2c4c", corner_radius=0, height=64)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            header, text="VT Rate Mail Sender", font=ctk.CTkFont(size=20, weight="bold"), text_color="#fff"
        ).grid(row=0, column=0, padx=18, pady=(10, 0), sticky="w")
        ctk.CTkLabel(
            header,
            text="Soạn nội dung như New Mail Outlook (dán bảng + chữ ký) · Semi-Auto",
            font=ctk.CTkFont(size=12),
            text_color="#b8c9dc",
        ).grid(row=1, column=0, padx=18, pady=(0, 8), sticky="w")
        self.auth_label = ctk.CTkLabel(header, text="Outlook: chưa kết nối", text_color="#d7e4f2")
        self.auth_label.grid(row=0, column=1, rowspan=2, sticky="e", padx=10)
        self.btn_login = ctk.CTkButton(header, text="Mở Outlook", width=120, command=self._connect_outlook)
        self.btn_login.grid(row=0, column=2, rowspan=2, padx=(0, 16))

    def _build_toolbar(self, parent: ctk.CTkFrame) -> None:
        bar = ctk.CTkFrame(parent, fg_color="#ffffff", corner_radius=10)
        bar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=12, pady=(10, 6))
        bar.grid_columnconfigure(6, weight=1)
        ctk.CTkButton(bar, text="Import Excel", width=120, command=self._import).grid(
            row=0, column=0, padx=(12, 6), pady=10
        )
        ctk.CTkButton(bar, text="Kiểm tra", width=90, fg_color="#2e7d32", command=self._validate_all).grid(
            row=0, column=1, padx=6, pady=10
        )
        ctk.CTkLabel(bar, text="Delay").grid(row=0, column=2, padx=(14, 4))
        self.delay_min_var = ctk.StringVar(value=str(self.config_data.delay_min_seconds))
        self.delay_max_var = ctk.StringVar(value=str(self.config_data.delay_max_seconds))
        ctk.CTkEntry(bar, textvariable=self.delay_min_var, width=42).grid(row=0, column=3)
        ctk.CTkLabel(bar, text="–").grid(row=0, column=4)
        ctk.CTkEntry(bar, textvariable=self.delay_max_var, width=42).grid(row=0, column=5)
        self.stats_label = ctk.CTkLabel(bar, text="0 agency", text_color="#555")
        self.stats_label.grid(row=0, column=6, sticky="e", padx=14)

    def _build_left(self, parent: ctk.CTkFrame) -> None:
        left = ctk.CTkFrame(parent, fg_color="#ffffff", corner_radius=12)
        left.grid(row=1, column=0, sticky="nsew", padx=(12, 6), pady=(0, 12))
        left.grid_rowconfigure(1, weight=1)
        left.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(left, text="Danh sách agency", font=ctk.CTkFont(size=15, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(12, 0)
        )
        ctk.CTkLabel(
            left,
            text="Excel: Company · Account Name · Account Mail · Mail cc",
            text_color="#777",
            font=ctk.CTkFont(size=11),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(36, 4))
        self.list_frame = ctk.CTkScrollableFrame(left, fg_color="#f4f7fb")
        self.list_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.list_frame.grid_columnconfigure(0, weight=1)
        self._row_widgets: list[ctk.CTkFrame] = []

    def _build_right(self, parent: ctk.CTkFrame) -> None:
        right = ctk.CTkFrame(parent, fg_color="#ffffff", corner_radius=12)
        right.grid(row=1, column=1, sticky="nsew", padx=(6, 12), pady=(0, 12))
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(3, weight=3)
        right.grid_rowconfigure(6, weight=1)

        # --- Outlook-like header fields ---
        header_box = ctk.CTkFrame(right, fg_color="#eef2f7", corner_radius=8)
        header_box.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        header_box.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(header_box, text="Subject", width=90, anchor="w", text_color="#555").grid(
            row=0, column=0, sticky="w", padx=10, pady=(10, 4)
        )
        ctk.CTkEntry(header_box, textvariable=self.subject_var).grid(
            row=0, column=1, columnspan=2, sticky="ew", padx=(0, 10), pady=(10, 4)
        )

        ctk.CTkLabel(header_box, text="Attachment", width=90, anchor="w", text_color="#555").grid(
            row=1, column=0, sticky="w", padx=10, pady=4
        )
        ctk.CTkEntry(header_box, textvariable=self.attachment_var).grid(
            row=1, column=1, sticky="ew", padx=(0, 6), pady=4
        )
        ctk.CTkButton(header_box, text="Chọn file", width=90, fg_color="#4a5568", command=self._pick_attachment).grid(
            row=1, column=2, padx=(0, 10), pady=4
        )

        # --- Compose actions (Outlook New Mail) ---
        actions = ctk.CTkFrame(right, fg_color="transparent")
        actions.grid(row=1, column=0, sticky="ew", padx=10, pady=(2, 4))
        ctk.CTkButton(
            actions,
            text="✉  Soạn trong Outlook (New Mail)",
            width=240,
            fg_color="#0f2c4c",
            command=self._compose_in_outlook,
        ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(
            actions, text="Đồng bộ từ Outlook", width=150, fg_color="#1f6aa5", command=self._sync_from_outlook
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            actions, text="Nạp HTML", width=90, fg_color="#4a5568", command=self._import_html
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            actions, text="Gợi ý lại", width=80, fg_color="#6b7280", command=self._reset_suggested
        ).pack(side="left", padx=4)

        # --- View mode ---
        mode = ctk.CTkFrame(right, fg_color="transparent")
        mode.grid(row=2, column=0, sticky="ew", padx=10, pady=(2, 2))
        ctk.CTkLabel(mode, text="Hiển thị:", text_color="#555").pack(side="left")
        ctk.CTkRadioButton(
            mode, text="Nội dung gốc (template)", variable=self._preview_mode, value="origin", command=self._refresh_html_view
        ).pack(side="left", padx=8)
        ctk.CTkRadioButton(
            mode,
            text="Preview agency đang chọn",
            variable=self._preview_mode,
            value="preview",
            command=self._refresh_html_view,
        ).pack(side="left", padx=8)
        ctk.CTkLabel(
            mode,
            text="{{agency_company}}  {{account_name}}",
            text_color="#888",
            font=ctk.CTkFont(size=11),
        ).pack(side="right")

        # --- HTML preview frame ---
        preview_wrap = ctk.CTkFrame(right, fg_color="#f7f9fc", corner_radius=8, border_width=1, border_color="#cfd8e3")
        preview_wrap.grid(row=3, column=0, sticky="nsew", padx=10, pady=4)
        preview_wrap.grid_rowconfigure(0, weight=1)
        preview_wrap.grid_columnconfigure(0, weight=1)
        self.preview_host = preview_wrap
        self._init_html_frame(preview_wrap)

        # --- Semi-auto controls ---
        controls = ctk.CTkFrame(right, fg_color="transparent")
        controls.grid(row=4, column=0, sticky="ew", padx=10, pady=(8, 2))
        controls.grid_columnconfigure((0, 1, 2), weight=1)
        self.btn_start = ctk.CTkButton(
            controls, text="▶  Semi-Auto", fg_color="#0f2c4c", height=40, command=self._start_send
        )
        self.btn_start.grid(row=0, column=0, padx=3, sticky="ew")
        self.btn_pause = ctk.CTkButton(
            controls, text="⏸  Pause", fg_color="#c47a00", height=40, state="disabled", command=self._toggle_pause
        )
        self.btn_pause.grid(row=0, column=1, padx=3, sticky="ew")
        self.btn_stop = ctk.CTkButton(
            controls, text="⏹  Stop", fg_color="#c62828", height=40, state="disabled", command=self._stop_send
        )
        self.btn_stop.grid(row=0, column=2, padx=3, sticky="ew")

        progress_wrap = ctk.CTkFrame(right, fg_color="transparent")
        progress_wrap.grid(row=5, column=0, sticky="ew", padx=10, pady=(4, 2))
        progress_wrap.grid_columnconfigure(0, weight=1)
        self.progress_bar = ctk.CTkProgressBar(progress_wrap)
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        self.progress_bar.set(0)
        self.progress_label = ctk.CTkLabel(
            progress_wrap,
            text="Soạn trong Outlook → Đồng bộ → Import danh bạ → Semi-Auto",
            anchor="w",
            text_color="#555",
        )
        self.progress_label.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        self.log_box = ctk.CTkTextbox(right, height=100, font=ctk.CTkFont(family="Menlo", size=11))
        self.log_box.grid(row=6, column=0, sticky="nsew", padx=10, pady=(4, 10))

    def _init_html_frame(self, parent: ctk.CTkFrame) -> None:
        try:
            from tkinterweb import HtmlFrame

            frame = HtmlFrame(parent, messages_enabled=False)
            frame.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
            self._html_frame = frame
            self._html_fallback = None
        except Exception:  # noqa: BLE001
            self._html_frame = None
            self._html_fallback = ctk.CTkTextbox(parent, font=ctk.CTkFont(family="Menlo", size=12))
            self._html_fallback.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)

    def _refresh_html_view(self) -> None:
        html_body = self.template_html or ""
        if self._preview_mode.get() == "preview" and self._selected_index is not None and self.mails:
            mail = self.mails[self._selected_index]
            mail.subject = self.subject_var.get().strip()
            mail.template_mail = self.template_html
            mail.attachment = self.attachment_var.get().strip()
            html_body = render_body_html(mail, "")
        doc = wrap_preview_document(html_body)
        if self._html_frame is not None:
            try:
                self._html_frame.load_html(doc)
            except Exception:  # noqa: BLE001
                pass
        elif self._html_fallback is not None:
            self._html_fallback.delete("1.0", "end")
            self._html_fallback.insert("1.0", html_body)

    # ---------------------------------------------------------- Compose
    def _reset_suggested(self) -> None:
        self.subject_var.set(SUGGESTED_SUBJECT)
        self.template_html = SUGGESTED_TEMPLATE_HTML
        self._save_config()
        self._refresh_html_view()

    def _pick_attachment(self) -> None:
        path = filedialog.askopenfilename(
            title="Chọn file đính kèm",
            filetypes=[("Excel / PDF", "*.xlsx *.xls *.pdf"), ("All", "*.*")],
        )
        if path:
            self.attachment_var.set(path)

    def _import_html(self) -> None:
        path = filedialog.askopenfilename(
            title="Nạp file HTML template",
            filetypes=[("HTML", "*.html *.htm"), ("All", "*.*")],
        )
        if not path:
            return
        self.template_html = Path(path).read_text(encoding="utf-8")
        self._save_config()
        self._refresh_html_view()
        messagebox.showinfo("HTML", "Đã nạp template HTML.")

    def _compose_in_outlook(self) -> None:
        if not self.sender.is_ready:
            try:
                self.sender.open_outlook()
                self._refresh_auth_label()
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("Outlook", str(exc))
                return
        self._save_config()

        def worker() -> None:
            try:
                msg = self.sender.open_template_composer(
                    self.template_html, self.subject_var.get().strip()
                )
                self.after(0, lambda: messagebox.showinfo("Soạn trong Outlook", msg))
            except Exception as exc:  # noqa: BLE001
                err = str(exc)
                self.after(0, lambda: messagebox.showerror("Outlook", err))

        threading.Thread(target=worker, daemon=True).start()

    def _sync_from_outlook(self) -> None:
        if not self.sender.is_ready:
            messagebox.showwarning("Outlook", "Hãy Mở Outlook trước.")
            return

        def worker() -> None:
            try:
                html_body = self.sender.sync_template_from_outlook()
                self.after(0, lambda: self._apply_synced_html(html_body))
            except Exception as exc:  # noqa: BLE001
                err = str(exc)
                self.after(0, lambda: messagebox.showerror("Đồng bộ", err))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_synced_html(self, html_body: str) -> None:
        self.template_html = html_body
        self._save_config()
        self._ensure_compose_on_mails()
        self._preview_mode.set("origin")
        self._refresh_html_view()
        messagebox.showinfo(
            "Đã đồng bộ",
            "Đã lấy nội dung HTML từ Outlook (bảng + chữ ký nếu bạn đã dán).\n"
            "Chọn agency và bật “Preview agency” để xem lời chào cá nhân hoá.",
        )

    def _ensure_compose_on_mails(self) -> None:
        subject = self.subject_var.get().strip()
        attachment = self.attachment_var.get().strip()
        for mail in self.mails:
            mail.subject = subject
            mail.template_mail = self.template_html
            mail.attachment = attachment

    # -------------------------------------------------------------- Outlook
    def _refresh_auth_label(self) -> None:
        if self.sender.is_ready:
            self.auth_label.configure(text=f"Outlook sẵn sàng · {self.sender.account_email}")
        else:
            self.auth_label.configure(text="Outlook: bấm “Mở Outlook”")

    def _connect_outlook(self) -> None:
        if self._login_busy:
            return
        self._login_busy = True
        self.btn_login.configure(state="disabled", text="Đang mở…")

        def worker() -> None:
            try:
                msg = self.sender.open_outlook()
                self.after(0, lambda: self._login_done(True, msg))
            except Exception as exc:  # noqa: BLE001
                err = str(exc)
                self.after(0, lambda: self._login_done(False, err))

        threading.Thread(target=worker, daemon=True).start()

    def _login_done(self, ok: bool, detail: str) -> None:
        self._login_busy = False
        self.btn_login.configure(state="normal", text="Mở Outlook")
        if ok:
            self.auth_label.configure(text=f"Outlook sẵn sàng · {self.sender.account_email}")
            messagebox.showinfo("Outlook", detail)
        else:
            self._refresh_auth_label()
            messagebox.showerror("Outlook", detail)

    # -------------------------------------------------------- Import / list
    def _import(self) -> None:
        path = filedialog.askopenfilename(
            title="Import danh bạ (4 cột)",
            filetypes=[("Excel / CSV", "*.xlsx *.csv"), ("Excel", "*.xlsx"), ("CSV", "*.csv")],
        )
        if not path:
            return
        try:
            self.mails = import_agencies(path)
            self._ensure_compose_on_mails()
            self._validate_all(silent=True)
            self._rebuild_list()
            if self.mails:
                self._select_row(0)
            messagebox.showinfo("Import", f"Đã nạp {len(self.mails)} agency.")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Lỗi import", str(exc))

    def _validate_all(self, silent: bool = False) -> None:
        self._ensure_compose_on_mails()
        ok = bad = 0
        for m in self.mails:
            errs = m.validate()
            if not errs and m.attachment.strip():
                try:
                    self.sender.resolve_attachment(m)
                except FileNotFoundError as exc:
                    errs.append(str(exc))
            if errs:
                m.status = MailStatus.PENDING
                m.error = "; ".join(errs)
                bad += 1
            else:
                m.status = MailStatus.READY
                m.error = ""
                ok += 1
        self._rebuild_list()
        if not silent:
            messagebox.showinfo("Kiểm tra", f"Sẵn sàng: {ok}\nCần sửa: {bad}")

    def _rebuild_list(self) -> None:
        for w in self._row_widgets:
            w.destroy()
        self._row_widgets.clear()
        head = ctk.CTkFrame(self.list_frame, fg_color="#e8eef5")
        head.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        head.grid_columnconfigure(1, weight=1)
        for col, (text, width) in enumerate([("#", 28), ("Company", 0), ("Email", 150), ("Status", 64)]):
            ctk.CTkLabel(
                head, text=text, width=width or 140, anchor="w", font=ctk.CTkFont(size=11, weight="bold")
            ).grid(row=0, column=col, padx=5, pady=5, sticky="ew" if col == 1 else "w")
        self._row_widgets.append(head)

        for i, mail in enumerate(self.mails):
            row = ctk.CTkFrame(self.list_frame, fg_color="#ffffff", corner_radius=6, cursor="hand2")
            row.grid(row=i + 1, column=0, sticky="ew", pady=2)
            row.grid_columnconfigure(1, weight=1)

            def bind_click(widget, index=i):
                widget.bind("<Button-1>", lambda _e, idx=index: self._select_row(idx))
                for child in widget.winfo_children():
                    child.bind("<Button-1>", lambda _e, idx=index: self._select_row(idx))

            ctk.CTkLabel(row, text=str(i + 1), width=28, text_color="#888").grid(row=0, column=0, padx=5, pady=8)
            ctk.CTkLabel(row, text=mail.agency_company or "(no company)", anchor="w").grid(
                row=0, column=1, sticky="ew"
            )
            ctk.CTkLabel(row, text=mail.account_mail, width=150, anchor="w").grid(row=0, column=2, padx=4)
            color = STATUS_COLORS.get(mail.status, "#666")
            ctk.CTkLabel(row, text=mail.status.value, width=64, text_color=color).grid(row=0, column=3, padx=5)
            bind_click(row)
            self._row_widgets.append(row)

        ready = sum(1 for m in self.mails if m.status == MailStatus.READY)
        sent = sum(1 for m in self.mails if m.status == MailStatus.SENT)
        failed = sum(1 for m in self.mails if m.status == MailStatus.FAILED)
        self.stats_label.configure(
            text=f"{len(self.mails)} agency · ready {ready} · sent {sent} · fail {failed}"
        )

    def _select_row(self, index: int) -> None:
        if index < 0 or index >= len(self.mails):
            return
        self._selected_index = index
        self._ensure_compose_on_mails()
        for i, w in enumerate(self._row_widgets[1:]):
            w.configure(fg_color="#e3f0ff" if i == index else "#ffffff")
        if self._preview_mode.get() == "preview":
            self._refresh_html_view()

    # ------------------------------------------------------------- Semi-auto
    def _start_send(self) -> None:
        if not self.mails:
            messagebox.showwarning("Trống", "Hãy Import Excel trước.")
            return
        self._ensure_compose_on_mails()
        self._save_config()
        if not self.subject_var.get().strip() or not self.template_html.strip():
            messagebox.showwarning("Template", "Thiếu Subject hoặc nội dung mail.")
            return
        attachment = self.attachment_var.get().strip()
        if attachment and not Path(attachment).is_file():
            messagebox.showerror("Attachment", f"Không tìm thấy file:\n{attachment}")
            return
        if not self.sender.is_ready:
            try:
                self.sender.open_outlook()
                self._refresh_auth_label()
            except Exception as exc:  # noqa: BLE001
                messagebox.showwarning("Outlook", str(exc))
                return

        self._validate_all(silent=True)
        to_send = [m for m in self.mails if m.status == MailStatus.READY]
        if not to_send:
            messagebox.showinfo("Chưa sẵn sàng", "Không có agency ready.")
            return
        try:
            dmin = int(self.delay_min_var.get())
            dmax = int(self.delay_max_var.get())
        except ValueError:
            messagebox.showerror("Delay", "Delay phải là số.")
            return
        if not messagebox.askyesno(
            "Xác nhận",
            f"Gửi {len(to_send)} mail qua Outlook?\nDelay {dmin}–{dmax}s\n"
            "Nội dung HTML (bảng/chữ ký) đã đồng bộ sẽ được dùng.",
        ):
            return

        self.queue.configure_delay(dmin, dmax)
        self.btn_start.configure(state="disabled")
        self.btn_pause.configure(state="normal", text="⏸  Pause")
        self.btn_stop.configure(state="normal")
        self.queue.start(to_send)

    def _toggle_pause(self) -> None:
        if not self.queue.is_running:
            return
        if self.queue.progress.is_paused:
            self.queue.resume()
            self.btn_pause.configure(text="⏸  Pause")
        else:
            self.queue.pause()
            self.btn_pause.configure(text="▶  Resume")

    def _stop_send(self) -> None:
        self.queue.stop()
        self.btn_start.configure(state="normal")
        self.btn_pause.configure(state="disabled", text="⏸  Pause")
        self.btn_stop.configure(state="disabled")

    def _on_progress(self, progress: SendProgress) -> None:
        self.after(0, lambda: self._apply_progress(progress))

    def _apply_progress(self, progress: SendProgress) -> None:
        self.progress_bar.set(progress.percent / 100.0)
        if progress.is_paused:
            status = f"Pause · {progress.wait_remaining:.0f}s · {progress.current_agency}"
        elif progress.is_running:
            wait = f" · nghỉ {progress.wait_remaining:.0f}s" if progress.wait_remaining > 0 else ""
            status = f"{progress.done}/{progress.total} · OK {progress.sent} · lỗi {progress.failed}{wait}"
        else:
            status = f"Xong · OK {progress.sent} · lỗi {progress.failed}"
            self.btn_start.configure(state="normal")
            self.btn_pause.configure(state="disabled", text="⏸  Pause")
            self.btn_stop.configure(state="disabled")
            self._persist_logs(progress)
        self.progress_label.configure(text=status)
        self._rebuild_list()
        self.log_box.delete("1.0", "end")
        for entry in progress.logs[-200:]:
            self.log_box.insert(
                "end",
                f"[{entry.display_time()}] {entry.status.upper():7} | "
                f"{entry.agency_company[:24]:<24} | {entry.account_mail} | {entry.message}\n",
            )
        self.log_box.see("end")

    def _persist_logs(self, progress: SendProgress) -> None:
        if not progress.logs:
            return
        LOG_DIR.mkdir(exist_ok=True)
        from datetime import datetime

        path = LOG_DIR / f"send_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        data = [
            {
                "time": e.display_time(),
                "agency": e.agency_company,
                "mail": e.account_mail,
                "subject": e.subject,
                "status": e.status,
                "message": e.message,
            }
            for e in progress.logs
        ]
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    def _on_close(self) -> None:
        if self.queue.is_running:
            if not messagebox.askyesno("Đang gửi", "Stop và thoát?"):
                return
            self.queue.stop()
        try:
            self._save_config()
        except Exception:
            pass
        self.destroy()


def run() -> None:
    app = VTRateMailApp()
    app.mainloop()
