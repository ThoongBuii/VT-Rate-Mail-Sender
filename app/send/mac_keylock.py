"""Khoá phím macOS tạm thời (CGEventTap) — chặn gõ user, cho phép phím synthetic của app.

Cần Accessibility (cùng quyền với AX Subject).
"""

from __future__ import annotations

import threading
import time
from typing import Optional


class KeyboardLock:
    """
    Nuốt KeyDown/KeyUp/FlagsChanged từ bàn phím thật.
    Bọc `_cg_key`: bật allow_synth để Tab/Cmd+V/Cmd+Return vẫn chạy.
    """

    def __init__(self) -> None:
        self._tap = None
        self._src = None
        self._thread: Optional[threading.Thread] = None
        self._loop = None
        self._allow_synth = False
        self._ready = threading.Event()
        self._failed: Optional[str] = None

    @property
    def active(self) -> bool:
        return self._tap is not None and self._failed is None

    def allow_synth(self, on: bool) -> None:
        self._allow_synth = bool(on)

    def start(self) -> bool:
        if self._tap is not None:
            return True
        self._ready.clear()
        self._failed = None
        self._thread = threading.Thread(
            target=self._run_loop, name="vt-keylock", daemon=True
        )
        self._thread.start()
        if not self._ready.wait(timeout=2.0):
            self._failed = "timeout"
            return False
        return self._failed is None

    def stop(self) -> None:
        try:
            from Quartz import (  # type: ignore
                CFRunLoopStop,
                CGEventTapEnable,
            )

            if self._tap is not None:
                try:
                    CGEventTapEnable(self._tap, False)
                except Exception:  # noqa: BLE001
                    pass
            loop = self._loop
            if loop is not None:
                try:
                    CFRunLoopStop(loop)
                except Exception:  # noqa: BLE001
                    pass
        except Exception:  # noqa: BLE001
            pass
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=1.0)
        self._tap = None
        self._src = None
        self._loop = None
        self._thread = None
        self._allow_synth = False

    def _run_loop(self) -> None:
        try:
            from Quartz import (  # type: ignore
                CFMachPortCreateRunLoopSource,
                CFRunLoopAddSource,
                CFRunLoopGetCurrent,
                CFRunLoopRun,
                CGEventMaskBit,
                CGEventTapCreate,
                CGEventTapEnable,
                kCGEventFlagsChanged,
                kCGEventKeyDown,
                kCGEventKeyUp,
                kCGEventTapOptionDefault,
                kCGHeadInsertEventTap,
                kCGSessionEventTap,
                kCFRunLoopCommonModes,
            )
        except Exception as exc:  # noqa: BLE001
            self._failed = str(exc)
            self._ready.set()
            return

        mask = (
            CGEventMaskBit(kCGEventKeyDown)
            | CGEventMaskBit(kCGEventKeyUp)
            | CGEventMaskBit(kCGEventFlagsChanged)
        )
        lock = self

        def _callback(_proxy, etype, event, _refcon):
            if lock._allow_synth:
                return event
            if etype in (kCGEventKeyDown, kCGEventKeyUp, kCGEventFlagsChanged):
                return None
            return event

        try:
            tap = CGEventTapCreate(
                kCGSessionEventTap,
                kCGHeadInsertEventTap,
                kCGEventTapOptionDefault,
                mask,
                _callback,
                None,
            )
        except Exception as exc:  # noqa: BLE001
            self._failed = str(exc)
            self._ready.set()
            return

        if tap is None:
            self._failed = "CGEventTapCreate returned None (cần Accessibility)"
            self._ready.set()
            return

        self._tap = tap
        self._loop = CFRunLoopGetCurrent()
        self._src = CFMachPortCreateRunLoopSource(None, tap, 0)
        CFRunLoopAddSource(self._loop, self._src, kCFRunLoopCommonModes)
        CGEventTapEnable(tap, True)
        self._ready.set()
        CFRunLoopRun()


_LOCK: Optional[KeyboardLock] = None


def get_keyboard_lock() -> KeyboardLock:
    global _LOCK
    if _LOCK is None:
        _LOCK = KeyboardLock()
    return _LOCK


def keyboard_guard(*, pre_focus_s: float = 0.5):
    """
    Context: bật khoá phím → chờ pre_focus_s → (caller focus Outlook + gửi) → tắt khoá.
    """

    class _Guard:
        def __enter__(self_inner):
            lock = get_keyboard_lock()
            ok = lock.start()
            if ok and pre_focus_s > 0:
                time.sleep(pre_focus_s)
            self_inner.lock = lock
            self_inner.ok = ok
            return lock

        def __exit__(self_inner, *exc):
            try:
                self_inner.lock.stop()
            except Exception:  # noqa: BLE001
                pass
            return False

    return _Guard()
