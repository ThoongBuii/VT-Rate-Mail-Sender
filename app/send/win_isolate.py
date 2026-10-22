"""Tách UI New Mail autosend khỏi cửa sổ Outlook đang làm việc (Windows).

A–D: nhớ focus → Display → ẩn/off-screen Inspector → trả focus.
Không đụng nội dung mail / WordEditor / Send — chỉ xử lý HWND.
"""

from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Any, Optional

# Win32
_SW_HIDE = 0
_SW_SHOW = 5
_SWP_NOSIZE = 0x0001
_SWP_NOZORDER = 0x0004
_SWP_NOACTIVATE = 0x0010
_SWP_SHOWWINDOW = 0x0040
_SWP_HIDEWINDOW = 0x0080
_HWND_BOTTOM = 1


@dataclass
class IsolateState:
    prev_foreground: int = 0
    inspector_hwnd: int = 0
    hidden: bool = False
    disabled: bool = False


def _user32():
    import ctypes

    return ctypes.windll.user32  # type: ignore[attr-defined]


def get_foreground_hwnd() -> int:
    try:
        return int(_user32().GetForegroundWindow() or 0)
    except Exception:  # noqa: BLE001
        return 0


def inspector_hwnd(inspector: Any) -> int:
    """Outlook Inspector.HWND (Outlook 2013+)."""
    try:
        hwnd = int(getattr(inspector, "HWND", 0) or 0)
        if hwnd:
            return hwnd
    except Exception:  # noqa: BLE001
        pass
    return 0


def _is_window(hwnd: int) -> bool:
    if not hwnd:
        return False
    try:
        return bool(_user32().IsWindow(hwnd))
    except Exception:  # noqa: BLE001
        return False


def hide_inspector_window(hwnd: int) -> bool:
    """Ẩn + đẩy ra ngoài màn hình + không activate (giảm nháy / cướp focus)."""
    if not _is_window(hwnd):
        return False
    user32 = _user32()
    try:
        # Đẩy ra ngoài vùng nhìn thấy trước khi hide
        user32.SetWindowPos(
            hwnd,
            _HWND_BOTTOM,
            -32000,
            -32000,
            0,
            0,
            _SWP_NOSIZE | _SWP_NOACTIVATE,
        )
    except Exception:  # noqa: BLE001
        pass
    try:
        user32.ShowWindow(hwnd, _SW_HIDE)
    except Exception:  # noqa: BLE001
        return False
    try:
        # Nhấn mạnh: không đưa lên foreground
        user32.SetWindowPos(
            hwnd,
            _HWND_BOTTOM,
            -32000,
            -32000,
            0,
            0,
            _SWP_NOSIZE | _SWP_NOACTIVATE | _SWP_HIDEWINDOW,
        )
    except Exception:  # noqa: BLE001
        pass
    return True


def disable_inspector_window(hwnd: int) -> bool:
    """Chặn nhận chuột/phím — chỉ gọi khi WordEditor đã xong (tránh rủi ro nội dung)."""
    if not _is_window(hwnd):
        return False
    try:
        _user32().EnableWindow(hwnd, False)
        return True
    except Exception:  # noqa: BLE001
        return False


def enable_inspector_window(hwnd: int) -> bool:
    if not _is_window(hwnd):
        return False
    try:
        _user32().EnableWindow(hwnd, True)
        return True
    except Exception:  # noqa: BLE001
        return False


def restore_foreground(hwnd: int) -> bool:
    """Trả focus về cửa sổ user đang làm (chat / mail khác)."""
    if not hwnd or not _is_window(hwnd):
        return False
    user32 = _user32()
    try:
        # AttachThreadInput giúp SetForegroundWindow đáng tin hơn trên Win10/11
        import ctypes

        pid = ctypes.c_ulong()
        target_tid = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        cur_tid = user32.GetCurrentThreadId()
        attached = False
        if target_tid and cur_tid and target_tid != cur_tid:
            attached = bool(user32.AttachThreadInput(cur_tid, target_tid, True))
        try:
            user32.ShowWindow(hwnd, _SW_SHOW)
            user32.SetForegroundWindow(hwnd)
            user32.BringWindowToTop(hwnd)
        finally:
            if attached:
                user32.AttachThreadInput(cur_tid, target_tid, False)
        return True
    except Exception:  # noqa: BLE001
        try:
            user32.SetForegroundWindow(hwnd)
            return True
        except Exception:  # noqa: BLE001
            return False


def isolate_after_display(
    mail_item: Any,
    *,
    prev_foreground: int,
    settle_ms: float = 0.08,
) -> IsolateState:
    """
    Gọi ngay sau Display(False).
    Ẩn Inspector autosend + trả focus — chưa Disable (để WordEditor/InsertFile an toàn).
    """
    state = IsolateState(prev_foreground=int(prev_foreground or 0))
    insp = None
    try:
        insp = mail_item.GetInspector
    except Exception:  # noqa: BLE001
        insp = None

    hwnd = inspector_hwnd(insp) if insp is not None else 0
    # Đôi khi HWND chưa sẵn sàng ngay sau Display
    if not hwnd:
        time.sleep(settle_ms)
        try:
            insp = mail_item.GetInspector
            hwnd = inspector_hwnd(insp)
        except Exception:  # noqa: BLE001
            hwnd = 0

    state.inspector_hwnd = hwnd
    if hwnd:
        state.hidden = hide_inspector_window(hwnd)

    if state.prev_foreground and state.prev_foreground != hwnd:
        restore_foreground(state.prev_foreground)
        # Outlook đôi khi giật focus lại sau Display
        time.sleep(0.05)
        restore_foreground(state.prev_foreground)

    return state


def reassert_user_focus(state: IsolateState) -> None:
    """Gọi lại sau các bước dài (InsertFile) nếu Outlook kéo focus."""
    if not state:
        return
    if state.inspector_hwnd:
        hide_inspector_window(state.inspector_hwnd)
    if state.prev_foreground and state.prev_foreground != state.inspector_hwnd:
        restore_foreground(state.prev_foreground)


def lock_input_after_content(state: IsolateState) -> None:
    """Sau khi nội dung mail đã xong: Disable Inspector (lớp C) + giữ ẩn + trả focus."""
    if not state or not state.inspector_hwnd:
        return
    hide_inspector_window(state.inspector_hwnd)
    state.disabled = disable_inspector_window(state.inspector_hwnd)
    if state.prev_foreground and state.prev_foreground != state.inspector_hwnd:
        restore_foreground(state.prev_foreground)


def unlock_before_send(state: IsolateState) -> None:
    """Trước Send: bật lại enable (Send không cần UI, nhưng an toàn hơn trên một số build)."""
    if not state or not state.inspector_hwnd:
        return
    if state.disabled:
        enable_inspector_window(state.inspector_hwnd)
        state.disabled = False
    # Vẫn giữ ẩn — không Show
    hide_inspector_window(state.inspector_hwnd)
    if state.prev_foreground and state.prev_foreground != state.inspector_hwnd:
        restore_foreground(state.prev_foreground)
