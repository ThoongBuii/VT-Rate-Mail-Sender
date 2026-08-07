from __future__ import annotations

import json
import threading
import uuid
from pathlib import Path
from typing import Any, Optional

from flask import Flask, jsonify, render_template, request

from .importer import import_agencies
from .models import AgencyMail, AppConfig, MailStatus, SendProgress
from .paths import bundle_dir, user_data_dir, web_dir
from .queue_worker import SemiAutoQueue
from .sender import OutlookDesktopSender
from .template_engine import SUGGESTED_SUBJECT, SUGGESTED_TEMPLATE_HTML, render_body_html, render_subject

ROOT = user_data_dir()
BUNDLE = bundle_dir()
CONFIG_PATH = ROOT / "config.json"
TEMPLATE_STORE = ROOT / "last_template.html"
LOG_DIR = ROOT / "logs"
UPLOAD_DIR = ROOT / "attachments"
WEB_DIR = web_dir()

flask_app = Flask(
    __name__,
    template_folder=str(WEB_DIR / "templates"),
    static_folder=str(WEB_DIR / "static"),
)


class AppState:
    def __init__(self) -> None:
        self.config = self._load_config()
        self.sender = OutlookDesktopSender(self.config)
        self.queue = SemiAutoQueue(
            sender=self.sender,
            delay_min=self.config.delay_min_seconds,
            delay_max=self.config.delay_max_seconds,
            on_progress=self._on_progress,
        )
        self.mails: list[AgencyMail] = []
        self.subject: str = self.config.default_subject or SUGGESTED_SUBJECT
        self.attachment: str = self.config.default_attachment or ""
        self.template_html: str = self._load_template_html()
        self.selected_index: int = 0
        self.progress: dict[str, Any] = {
            "total": 0,
            "done": 0,
            "sent": 0,
            "failed": 0,
            "skipped": 0,
            "percent": 0,
            "is_running": False,
            "is_paused": False,
            "wait_remaining": 0,
            "current_agency": "",
            "message": "Sẵn sàng",
            "logs": [],
        }
        self._lock = threading.Lock()

    def _load_config(self) -> AppConfig:
        cfg_path = CONFIG_PATH if CONFIG_PATH.exists() else (BUNDLE / "config.example.json")
        if cfg_path.exists():
            return AppConfig.from_dict(json.loads(cfg_path.read_text(encoding="utf-8")))
        return AppConfig.from_dict({})

    def _load_template_html(self) -> str:
        if TEMPLATE_STORE.exists():
            text = TEMPLATE_STORE.read_text(encoding="utf-8").strip()
            if text:
                return text
        # migrate from old huge config.json default_template
        if CONFIG_PATH.exists():
            try:
                data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                old = str(data.get("default_template") or "").strip()
                if old:
                    TEMPLATE_STORE.write_text(old, encoding="utf-8")
                    return old
            except Exception:  # noqa: BLE001
                pass
        return SUGGESTED_TEMPLATE_HTML

    def save_config(self) -> None:
        self.config.default_subject = self.subject
        self.config.default_attachment = self.attachment
        # Template HTML lưu file riêng (tránh config.json phình vì chữ ký/ảnh)
        TEMPLATE_STORE.write_text(self.template_html or "", encoding="utf-8")
        self.config.default_template = ""
        self.config.delay_min_seconds = self.queue.delay_min
        self.config.delay_max_seconds = self.queue.delay_max
        CONFIG_PATH.write_text(
            json.dumps(self.config.to_dict(), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    def apply_compose_to_mails(self) -> None:
        for mail in self.mails:
            mail.subject = self.subject
            mail.template_mail = self.template_html
            mail.attachment = self.attachment

    def _on_progress(self, progress: SendProgress) -> None:
        with self._lock:
            if progress.is_paused:
                message = f"Pause · {progress.wait_remaining:.0f}s · {progress.current_agency}"
            elif progress.is_running:
                wait = f" · nghỉ {progress.wait_remaining:.0f}s" if progress.wait_remaining > 0 else ""
                message = (
                    f"{progress.done}/{progress.total} · OK {progress.sent} · "
                    f"lỗi {progress.failed}{wait}"
                )
            else:
                message = f"Xong · OK {progress.sent} · lỗi {progress.failed}"
                if progress.logs:
                    self._persist_logs(progress)

            self.progress = {
                "total": progress.total,
                "done": progress.done,
                "sent": progress.sent,
                "failed": progress.failed,
                "skipped": progress.skipped,
                "percent": progress.percent,
                "is_running": progress.is_running,
                "is_paused": progress.is_paused,
                "wait_remaining": progress.wait_remaining,
                "current_agency": progress.current_agency,
                "message": message,
                "logs": [
                    {
                        "time": e.display_time(),
                        "status": e.status,
                        "agency": e.agency_company,
                        "mail": e.account_mail,
                        "message": e.message,
                    }
                    for e in progress.logs[-200:]
                ],
            }

    def _persist_logs(self, progress: SendProgress) -> None:
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

    def snapshot(self) -> dict[str, Any]:
        ready = sum(1 for m in self.mails if m.status == MailStatus.READY)
        sent = sum(1 for m in self.mails if m.status == MailStatus.SENT)
        failed = sum(1 for m in self.mails if m.status == MailStatus.FAILED)
        return {
            "outlook_ready": self.sender.is_ready,
            "outlook_account": self.sender.account_email,
            "subject": self.subject,
            "attachment": self.attachment,
            "template_html": self.template_html,
            "selected_index": self.selected_index,
            "stats": {
                "total": len(self.mails),
                "ready": ready,
                "sent": sent,
                "failed": failed,
            },
            "mails": [
                {
                    "index": i,
                    "agency_company": m.agency_company,
                    "account_name": m.account_name,
                    "account_mail": m.account_mail,
                    "mail_cc": m.mail_cc,
                    "status": m.status.value,
                    "error": m.error,
                }
                for i, m in enumerate(self.mails)
            ],
            "delay_min": self.queue.delay_min,
            "delay_max": self.queue.delay_max,
            "progress": self.progress,
            "suggested_subject": SUGGESTED_SUBJECT,
            "suggested_html": SUGGESTED_TEMPLATE_HTML,
        }


STATE = AppState()


@flask_app.get("/")
def index():
    return render_template("index.html")


@flask_app.get("/api/state")
def api_state():
    return jsonify(STATE.snapshot())


@flask_app.get("/api/progress")
def api_progress():
    return jsonify(STATE.progress)


@flask_app.get("/api/preview")
def api_preview():
    """Preview mail đã render cho 1 agency (kiểm tra lần cuối trước khi gửi)."""
    if not STATE.mails:
        return jsonify({"ok": False, "error": "Chưa có danh sách agency"}), 400
    try:
        idx = int(request.args.get("index", STATE.selected_index or 0))
    except ValueError:
        idx = 0
    idx = max(0, min(idx, len(STATE.mails) - 1))
    STATE.selected_index = idx
    STATE.apply_compose_to_mails()
    mail = STATE.mails[idx]
    preview = STATE.sender.preview(mail)
    return jsonify(
        {
            "ok": True,
            "index": idx,
            "total": len(STATE.mails),
            "agency_company": mail.agency_company,
            "account_name": mail.account_name,
            "status": mail.status.value,
            "error": mail.error,
            "from": preview.get("from") or STATE.sender.account_email,
            "to": preview.get("to") or mail.account_mail,
            "cc": preview.get("cc") or mail.mail_cc or "",
            "subject": preview.get("subject") or render_subject(mail),
            "attachment": preview.get("attachment") or "(không có)",
            "attachment_ok": preview.get("attachment_ok", True),
            "attachment_error": preview.get("attachment_error") or "",
            "body_html": preview.get("body_html") or render_body_html(mail, ""),
        }
    )


@flask_app.post("/api/outlook/open")
def api_outlook_open():
    try:
        msg = STATE.sender.open_outlook()
        return jsonify({"ok": True, "message": msg, "account": STATE.sender.account_email})
    except Exception as exc:  # noqa: BLE001
        return jsonify({"ok": False, "error": str(exc)}), 400


@flask_app.post("/api/compose")
def api_compose():
    data = request.get_json(force=True, silent=True) or {}
    if "subject" in data:
        STATE.subject = str(data.get("subject") or "").strip()
    if "template_html" in data:
        STATE.template_html = str(data.get("template_html") or "")
    if "attachment" in data:
        STATE.attachment = str(data.get("attachment") or "").strip()
    if "delay_min" in data:
        STATE.queue.delay_min = int(data.get("delay_min") or 10)
    if "delay_max" in data:
        STATE.queue.delay_max = max(STATE.queue.delay_min, int(data.get("delay_max") or 20))
    if "selected_index" in data:
        STATE.selected_index = int(data.get("selected_index") or 0)
    STATE.apply_compose_to_mails()
    STATE.save_config()
    return jsonify({"ok": True})


@flask_app.post("/api/import")
def api_import():
    f = request.files.get("file")
    if not f:
        return jsonify({"ok": False, "error": "Thiếu file Excel"}), 400
    UPLOAD_DIR.mkdir(exist_ok=True)
    suffix = Path(f.filename or "list.xlsx").suffix.lower() or ".xlsx"
    path = UPLOAD_DIR / f"import_{uuid.uuid4().hex}{suffix}"
    f.save(path)
    try:
        STATE.mails = import_agencies(path)
        STATE.apply_compose_to_mails()
        _validate_mails()
        STATE.selected_index = 0 if STATE.mails else 0
        return jsonify({"ok": True, "count": len(STATE.mails), "state": STATE.snapshot()})
    except Exception as exc:  # noqa: BLE001
        return jsonify({"ok": False, "error": str(exc)}), 400


@flask_app.post("/api/attachment")
def api_attachment():
    f = request.files.get("file")
    if not f:
        return jsonify({"ok": False, "error": "Thiếu file"}), 400
    UPLOAD_DIR.mkdir(exist_ok=True)
    name = Path(f.filename or "attachment.bin").name
    dest = UPLOAD_DIR / name
    f.save(dest)
    STATE.attachment = str(dest.resolve())
    STATE.apply_compose_to_mails()
    STATE.save_config()
    return jsonify({"ok": True, "path": STATE.attachment, "name": name})


@flask_app.post("/api/validate")
def api_validate():
    STATE.apply_compose_to_mails()
    ok, bad = _validate_mails()
    return jsonify({"ok": True, "ready": ok, "bad": bad, "state": STATE.snapshot()})


def _validate_mails() -> tuple[int, int]:
    ok = bad = 0
    for m in STATE.mails:
        errs = m.validate()
        if not errs and m.attachment.strip():
            try:
                STATE.sender.resolve_attachment(m)
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
    return ok, bad


@flask_app.post("/api/send/start")
def api_send_start():
    data = request.get_json(force=True, silent=True) or {}
    STATE.subject = str(data.get("subject") or STATE.subject).strip()
    STATE.template_html = str(data.get("template_html") or STATE.template_html)
    if "attachment" in data:
        STATE.attachment = str(data.get("attachment") or "").strip()
    dmin = int(data.get("delay_min") or STATE.queue.delay_min)
    dmax = int(data.get("delay_max") or STATE.queue.delay_max)
    STATE.queue.configure_delay(dmin, dmax)
    STATE.apply_compose_to_mails()
    STATE.save_config()

    if not STATE.mails:
        return jsonify({"ok": False, "error": "Chưa có danh sách agency"}), 400
    if not STATE.subject or not STATE.template_html.strip():
        return jsonify({"ok": False, "error": "Thiếu Subject hoặc nội dung mail"}), 400
    if STATE.attachment and not Path(STATE.attachment).is_file():
        return jsonify({"ok": False, "error": f"Không tìm thấy file: {STATE.attachment}"}), 400
    if not STATE.sender.is_ready:
        try:
            STATE.sender.open_outlook()
        except Exception as exc:  # noqa: BLE001
            return jsonify({"ok": False, "error": str(exc)}), 400

    _validate_mails()
    to_send = [m for m in STATE.mails if m.status == MailStatus.READY]
    if not to_send:
        return jsonify({"ok": False, "error": "Không có agency ready"}), 400
    if STATE.queue.is_running:
        return jsonify({"ok": False, "error": "Đang chạy Semi-Auto"}), 400

    try:
        STATE.queue.start(to_send)
    except Exception as exc:  # noqa: BLE001
        return jsonify({"ok": False, "error": str(exc)}), 400
    return jsonify({"ok": True, "count": len(to_send)})


@flask_app.post("/api/send/pause")
def api_send_pause():
    STATE.queue.pause()
    return jsonify({"ok": True})


@flask_app.post("/api/send/resume")
def api_send_resume():
    STATE.queue.resume()
    return jsonify({"ok": True})


@flask_app.post("/api/send/stop")
def api_send_stop():
    STATE.queue.stop()
    return jsonify({"ok": True})


def create_flask_app() -> Flask:
    return flask_app
