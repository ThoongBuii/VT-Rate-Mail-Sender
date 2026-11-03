"""Khoá phím macOS tạm thời (CGEventTap) — chặn gõ user, cho phép phím synthetic của app.

Cần Accessibility (cùng quyền với AX Subject).
Phím do app bắn mang SYNTH_USERDATA nên tap cho qua kể cả khi đang nuốt phím thật.
allow_synth đếm lồng nhau, chỉ bật trong lúc post từng phím — không mở khóa suốt burst.
"""

from __future__ import annotations

import threading
import time
from typing import Optional

# CGEventSource userData — tap luôn cho qua, kể cả khi allow_synth đã hạ.
SYNTH_USERDATA = 0x56544D41

# CGEventType khi tap bị tắt (CGEventTypes.h). pyobjc có thể đưa số có dấu hoặc không dấu.
# Timeout: hệ thống tắt vì callback chậm — chỉ bật lại khi đang trong một mail.
# User input: chính app gọi CGEventTapEnable(false) lúc gửi xong — không được bật lại.
_TAP_TIMEOUT = frozenset({-2, 0xFFFFFFFE})
_TAP_USER_OFF = frozenset({-1, 0xFFFFFFFF})
_TAP_DISABLED = _TAP_TIMEOUT | _TAP_USER_OFF


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
        self._synth_depth = 0
        self._want_enabled = False
        self._state_lock = threading.Lock()
        self._ready = threading.Event()
        self._failed: Optional[str] = None

    @property
    def active(self) -> bool:
        return self._tap is not None and self._failed is None

    def allow_synth(self, on: bool) -> None:
        """Bật/tắt theo cặp. Lồng nhau: cờ chỉ tắt khi mọi lớp đã thả."""
        if on:
            self._synth_depth += 1
        else:
            self._synth_depth = max(0, self._synth_depth - 1)
        self._allow_synth = self._synth_depth > 0

    def set_enabled(self, on: bool) -> None:
        """Bật hoặc tắt nuốt phím. Cờ và lệnh enable nằm cùng một khóa với callback."""
        from Quartz import CGEventTapEnable  # type: ignore

        with self._state_lock:
            self._want_enabled = bool(on)
            tap = self._tap
            if tap is None:
                return
            try:
                CGEventTapEnable(tap, bool(on))
            except Exception:  # noqa: BLE001
                pass

    def start(self) -> bool:
        # Tap cũ còn sống thì dùng lại. Tạo mới mỗi mail sẽ chồng tap và nuốt Tab từ mail ~10.
        if self._thread is not None and self._thread.is_alive():
            self.set_enabled(True)
            return self._failed is None and self._tap is not None
        self._want_enabled = True
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

            with self._state_lock:
                # Hạ cờ trước. Callback timeout không được bật tap trở lại sau mail.
                self._want_enabled = False
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
            self._thread.join(timeout=2.0)
        self._synth_depth = 0
        self._allow_synth = False
        if self._thread and self._thread.is_alive():
            # Run loop chưa thoát — tắt tap, giữ nguyên để mail sau không tạo tap mới.
            self.set_enabled(False)
            return
        self._tap = None
        self._src = None
        self._loop = None
        self._thread = None

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
        tap_holder: dict = {}
        get_mark = None
        mark_field = None
        try:
            from Quartz import (  # type: ignore
                CGEventGetIntegerValueField,
                kCGEventSourceUserData,
            )

            get_mark = CGEventGetIntegerValueField
            mark_field = kCGEventSourceUserData
        except Exception:  # noqa: BLE001
            pass

        def _callback(_proxy, etype, event, _refcon):
            if etype in _TAP_DISABLED:
                # Chỉ cứu tap khi hệ thống tự tắt giữa một mail. Tắt có chủ đích thì để tắt.
                if etype in _TAP_TIMEOUT:
                    with lock._state_lock:
                        if lock._want_enabled:
                            tap = tap_holder.get("tap")
                            if tap is not None:
                                try:
                                    CGEventTapEnable(tap, True)
                                except Exception:  # noqa: BLE001
                                    pass
                return event
            if get_mark is not None and event is not None:
                try:
                    if int(get_mark(event, mark_field)) == SYNTH_USERDATA:
                        return event
                except Exception:  # noqa: BLE001
                    pass
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
            tap_holder["tap"] = tap
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
        with self._state_lock:
            if self._want_enabled:
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
