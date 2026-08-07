from __future__ import annotations

import random
import threading
import time
from typing import Callable, Optional

from .models import AgencyMail, MailStatus, SendLogEntry, SendProgress
from .sender import OutlookDesktopSender


ProgressCallback = Callable[[SendProgress], None]


class SemiAutoQueue:
    """
    Gửi từng mail một, nghỉ giữa các lần gửi.
    Không gửi đồng loạt / không BCC hàng loạt.
    Hỗ trợ Pause / Resume / Stop.
    """

    def __init__(
        self,
        sender: OutlookDesktopSender,
        delay_min: int = 10,
        delay_max: int = 20,
        on_progress: Optional[ProgressCallback] = None,
    ):
        self.sender = sender
        self.delay_min = delay_min
        self.delay_max = delay_max
        self.on_progress = on_progress
        self._thread: Optional[threading.Thread] = None
        self._stop = threading.Event()
        self._pause = threading.Event()
        self._pause.clear()  # not paused
        self.progress = SendProgress()
        self._mails: list[AgencyMail] = []

    def configure_delay(self, delay_min: int, delay_max: int) -> None:
        self.delay_min = max(0, delay_min)
        self.delay_max = max(self.delay_min, delay_max)

    def _emit(self) -> None:
        if self.on_progress:
            self.on_progress(self.progress)

    def _log(self, mail: AgencyMail, status: str, message: str = "") -> None:
        entry = SendLogEntry(
            timestamp=time.time(),
            row_index=mail.row_index,
            agency_company=mail.agency_company,
            account_mail=mail.account_mail,
            subject=mail.subject,
            status=status,
            message=message,
        )
        self.progress.logs.append(entry)

    def start(self, mails: list[AgencyMail], start_from: int = 0) -> None:
        if self.is_running:
            raise RuntimeError("Đang gửi rồi — hãy Pause/Stop trước.")
        self._mails = mails
        self._stop.clear()
        self._pause.clear()
        pending = [m for m in mails[start_from:] if m.status in (MailStatus.PENDING, MailStatus.READY, MailStatus.FAILED)]
        self.progress = SendProgress(
            total=len(pending),
            is_running=True,
            is_paused=False,
            logs=list(self.progress.logs),
        )
        self._thread = threading.Thread(target=self._run, args=(pending,), daemon=True)
        self._thread.start()
        self._emit()

    @property
    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def pause(self) -> None:
        self._pause.set()
        self.progress.is_paused = True
        self._emit()

    def resume(self) -> None:
        self._pause.clear()
        self.progress.is_paused = False
        self._emit()

    def stop(self) -> None:
        self._stop.set()
        self._pause.clear()  # unblock if waiting on pause
        self.progress.is_running = False
        self.progress.is_paused = False
        self._emit()

    def _wait_interruptible(self, seconds: float) -> bool:
        """Wait with pause/stop support. Returns False if stopped."""
        end = time.time() + seconds
        while time.time() < end:
            if self._stop.is_set():
                return False
            while self._pause.is_set():
                self.progress.is_paused = True
                self.progress.wait_remaining = max(0.0, end - time.time())
                self._emit()
                if self._stop.is_set():
                    return False
                time.sleep(0.2)
            self.progress.is_paused = False
            self.progress.wait_remaining = max(0.0, end - time.time())
            self._emit()
            time.sleep(0.2)
        self.progress.wait_remaining = 0.0
        return not self._stop.is_set()

    def _run(self, mails: list[AgencyMail]) -> None:
        try:
            for idx, mail in enumerate(mails):
                if self._stop.is_set():
                    break

                while self._pause.is_set():
                    self.progress.is_paused = True
                    self._emit()
                    if self._stop.is_set():
                        break
                    time.sleep(0.2)
                if self._stop.is_set():
                    break

                self.progress.current_index = idx
                self.progress.current_agency = mail.agency_company or mail.account_mail
                mail.status = MailStatus.SENDING
                self._emit()

                try:
                    errors = mail.validate()
                    if errors:
                        mail.status = MailStatus.SKIPPED
                        mail.error = "; ".join(errors)
                        self.progress.skipped += 1
                        self._log(mail, "skipped", mail.error)
                    else:
                        self.sender.send_one(mail)
                        mail.status = MailStatus.SENT
                        mail.sent_at = time.time()
                        mail.error = ""
                        self.progress.sent += 1
                        self._log(mail, "sent", "OK")
                except Exception as exc:  # noqa: BLE001 — log and continue queue
                    mail.status = MailStatus.FAILED
                    mail.error = str(exc)
                    self.progress.failed += 1
                    self._log(mail, "failed", str(exc))

                self._emit()

                # Nghỉ giữa các mail (không nghỉ sau mail cuối)
                if idx < len(mails) - 1 and not self._stop.is_set():
                    delay = random.uniform(self.delay_min, self.delay_max)
                    if not self._wait_interruptible(delay):
                        break
        finally:
            self.progress.is_running = False
            self.progress.is_paused = False
            self.progress.wait_remaining = 0.0
            self.progress.current_agency = ""
            self._emit()
