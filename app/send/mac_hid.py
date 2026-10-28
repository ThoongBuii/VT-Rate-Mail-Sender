"""Mac HID tối thiểu — quy trình chuẩn (đã probe Legacy):

AS (mac.py): New Mail + To/Cc/Subject  →  caret mặc định ở To
HID (trong KeyboardLock — chặn gõ user / Zalo):
  1) open(msg id) + buộc Outlook frontmost
  2) AX focus ô To, rồi select ô Subject (field editor thật)
  3) Đúng 1 phím Tab thuần (không Cmd) gửi thẳng Outlook → body
  4) Cmd+V → Cmd+Return

Không đọc lại focus trước khi dán. Không chuột. Không thử AX Body.
"""

from __future__ import annotations

import subprocess
import time

from .mac_keylock import SYNTH_USERDATA, get_keyboard_lock


def _run_nsapplescript(source: str) -> str:
    """AppleScript in-process (Accessibility của app) — không spawn osascript."""
    try:
        from Foundation import NSAppleScript  # type: ignore
    except ImportError:
        r = subprocess.run(
            ["osascript", "-e", source],
            capture_output=True,
            text=True,
            timeout=30,
        )
        if r.returncode != 0:
            raise RuntimeError((r.stderr or r.stdout or "osascript failed").strip())
        return (r.stdout or "").strip()

    script = NSAppleScript.alloc().initWithSource_(source)
    if script is None:
        raise RuntimeError("Không tạo được NSAppleScript.")
    result, error = script.executeAndReturnError_(None)
    if error:
        try:
            desc = str(error.objectForKey_("NSAppleScriptErrorMessage") or error)
        except Exception:  # noqa: BLE001
            desc = str(error)
        raise RuntimeError(desc or "NSAppleScript failed")
    if result is None:
        return ""
    try:
        return str(result.stringValue() or "")
    except Exception:  # noqa: BLE001
        return str(result)


def _esc_as(s: str) -> str:
    return (s or "").replace("\\", "\\\\").replace('"', '\\"')


def _compose_picker_as(subject_hint: str) -> str:
    """Chọn cửa sổ New Mail (title Legacy ≈ subject • account; ưu tiên đuôi agency)."""
    full = (subject_hint or "").strip()
    needle = _esc_as(full[:40])
    needle2 = _esc_as(full[:24]) if len(full) >= 24 else needle
    # Title hay dạng «… - Agency - Account» — đuôi agency ổn định hơn đầu subject
    tail = full
    if " - " in full:
        tail = full.rsplit(" - ", 1)[-1].strip()
    elif "||" in full:
        tail = full.rsplit("||", 1)[-1].strip()
    needle_tail = _esc_as(tail[:36]) if tail else ""
    return f'''
    set needle to "{needle}"
    set needle2 to "{needle2}"
    set needleTail to "{needle_tail}"
    set targetWin to missing value

    try
      set fw to front window
      set fn to name of fw as text
      if needle is not "" and fn contains needle then
        set targetWin to fw
      else if needle2 is not "" and fn contains needle2 then
        set targetWin to fw
      else if needleTail is not "" and fn contains needleTail then
        set targetWin to fw
      else if (fn contains "Untitled") or (fn contains "@") or (fn contains "New Message") or (fn contains "Outgoing") then
        set targetWin to fw
      else if (length of fn > 24) and (fn does not contain "Inbox") and (fn does not contain "Sent items") and (fn does not contain "Sent Items") and (fn does not contain "Drafts") and (fn does not contain "Deleted") then
        set targetWin to fw
      end if
    end try

    if targetWin is missing value then
      repeat with w in windows
        try
          set n to name of w as text
          if (needle is not "" and n contains needle) or (needle2 is not "" and n contains needle2) or (needleTail is not "" and n contains needleTail) then
            set targetWin to w
            exit repeat
          end if
        end try
      end repeat
    end if

    if targetWin is missing value then
      repeat with w in windows
        try
          set n to name of w as text
          if (n contains "Untitled") or (n contains "@") or (n contains "New Message") or (n contains "Outgoing") then
            set targetWin to w
            exit repeat
          end if
        end try
      end repeat
    end if

    if targetWin is missing value then
      repeat with w in windows
        try
          set n to name of w as text
          if (length of n > 28) and (n does not contain "Inbox") and (n does not contain "Sent items") and (n does not contain "Sent Items") and (n does not contain "Drafts") and (n does not contain "Deleted") and (n does not contain "Junk") then
            set targetWin to w
            exit repeat
          end if
        end try
      end repeat
    end if
'''


def activate_outlook(*, settle: float = 0.02) -> None:
    try:
        _run_nsapplescript(
            '''
tell application "Microsoft Outlook" to activate
tell application "System Events"
  tell process "Microsoft Outlook"
    set frontmost to true
  end tell
end tell
'''
        )
    except Exception:  # noqa: BLE001
        subprocess.run(
            ["osascript", "-e", 'tell application "Microsoft Outlook" to activate'],
            capture_output=True,
            timeout=12,
        )
    if settle > 0:
        time.sleep(settle)


def open_compose_by_id(msg_id: str) -> bool:
    """Mở lại New Mail theo id (một AppleScript: activate + open)."""
    mid = (msg_id or "").strip()
    if not mid or not mid.isdigit():
        return False
    script = f'''
tell application "Microsoft Outlook"
  activate
  try
    set msg to outgoing message id {mid}
    open msg
    return "yes"
  on error
    try
      set msg to message id {mid}
      open msg
      return "yes"
    on error
      return "no"
    end try
  end try
end tell
'''
    try:
        return _run_nsapplescript(script).strip().lower().startswith("yes")
    except Exception:  # noqa: BLE001
        return False


def close_compose_by_id(msg_id: str) -> bool:
    """Đóng đúng New Mail theo id (không Cmd+W đụng cửa sổ compose khác)."""
    mid = (msg_id or "").strip()
    if not mid or not mid.isdigit():
        return False
    script = f'''
tell application "Microsoft Outlook"
  try
    set msg to outgoing message id {mid}
    close msg saving no
    return "yes"
  on error
    try
      set msg to message id {mid}
      close msg saving no
      return "yes"
    on error
      return "no"
    end try
  end try
end tell
'''
    try:
        return _run_nsapplescript(script).strip().lower().startswith("yes")
    except Exception:  # noqa: BLE001
        return False


def claim_outlook_compose(*, subject_hint: str = "", msg_id: str = "") -> bool:
    """Một vòng nhanh: open(id) + raise compose (không lặp activate thừa)."""
    if msg_id:
        open_compose_by_id(msg_id)
    return raise_compose_window(subject_hint=subject_hint)


def raise_compose_window(*, subject_hint: str = "") -> bool:
    script = f'''
tell application "Microsoft Outlook" to activate
tell application "System Events"
  tell process "Microsoft Outlook"
    set frontmost to true
{_compose_picker_as(subject_hint)}
    if targetWin is missing value then return "no"
    try
      perform action "AXRaise" of targetWin
    end try
    try
      set value of attribute "AXMain" of targetWin to true
    end try
    return "yes"
  end tell
end tell
'''
    try:
        return _run_nsapplescript(script).strip().lower().startswith("yes")
    except Exception:  # noqa: BLE001
        return False


def wait_compose_window(
    *, timeout: float = 1.4, subject_hint: str = "", msg_id: str = ""
) -> bool:
    """Chờ New Mail — claim gọn, timeout ngắn (tránh đợi chồng nhiều vòng)."""
    # Lần 1 ngay
    if claim_outlook_compose(subject_hint=subject_hint, msg_id=msg_id):
        return True
    deadline = time.time() + max(0.2, timeout)
    while time.time() < deadline:
        time.sleep(0.05)
        if claim_outlook_compose(subject_hint=subject_hint, msg_id=msg_id):
            return True
    return False


_SYNTH_SOURCE = None
_SYNTH_SOURCE_FAILED = False


def _synth_source():
    """Event source có userData — tap nhận ra phím app dù allow_synth đã tắt."""
    global _SYNTH_SOURCE, _SYNTH_SOURCE_FAILED
    if _SYNTH_SOURCE_FAILED:
        return None
    if _SYNTH_SOURCE is not None:
        return _SYNTH_SOURCE
    try:
        from Quartz import (  # type: ignore
            CGEventSourceCreate,
            CGEventSourceSetUserData,
            kCGEventSourceStatePrivate,
        )

        # Private: không ăn modifier bàn phím thật, không để Cmd dính sang Tab mail sau.
        src = CGEventSourceCreate(kCGEventSourceStatePrivate)
        CGEventSourceSetUserData(src, SYNTH_USERDATA)
        _SYNTH_SOURCE = src
    except Exception:  # noqa: BLE001
        _SYNTH_SOURCE_FAILED = True
        return None
    return _SYNTH_SOURCE


def _mark_synth(event) -> None:
    try:
        from Quartz import CGEventSetIntegerValueField, kCGEventSourceUserData  # type: ignore

        CGEventSetIntegerValueField(event, kCGEventSourceUserData, SYNTH_USERDATA)
    except Exception:  # noqa: BLE001
        pass


_OUTLOOK_PID = 0


def _outlook_pid() -> int:
    global _OUTLOOK_PID
    if _OUTLOOK_PID > 0:
        return _OUTLOOK_PID
    try:
        from AppKit import NSWorkspace  # type: ignore

        for app in NSWorkspace.sharedWorkspace().runningApplications():
            bid = str(app.bundleIdentifier() or "")
            name = str(app.localizedName() or "")
            if bid == "com.microsoft.Outlook" or name == "Microsoft Outlook":
                _OUTLOOK_PID = int(app.processIdentifier())
                return _OUTLOOK_PID
    except Exception:  # noqa: BLE001
        return 0
    return 0


def _post_to_outlook(event) -> None:
    """Đưa phím thẳng vào Outlook — tap phiên (khoá phím) không đứng trên đường này."""
    from Quartz import CGEventPost, CGEventPostToPid, kCGHIDEventTap  # type: ignore

    pid = _outlook_pid()
    if pid > 0:
        CGEventPostToPid(pid, event)
        return
    CGEventPost(kCGHIDEventTap, event)


def _clear_outlook_modifiers() -> None:
    """Thả Cmd/Shift/Opt/Ctrl. Cmd dính từ mail trước biến Tab thành Cmd+Tab."""
    from Quartz import (  # type: ignore
        CGEventCreate,
        CGEventCreateKeyboardEvent,
        CGEventPost,
        CGEventSetFlags,
        CGEventSetType,
        kCGEventFlagsChanged,
        kCGHIDEventTap,
    )

    src = _synth_source()
    flags_ev = CGEventCreate(src)
    if flags_ev is not None:
        CGEventSetType(flags_ev, kCGEventFlagsChanged)
        CGEventSetFlags(flags_ev, 0)
        _post_to_outlook(flags_ev)
    for code in (55, 56, 58, 59):
        up = CGEventCreateKeyboardEvent(None, code, False)
        CGEventSetFlags(up, 0)
        CGEventPost(kCGHIDEventTap, up)


def _cg_key(keycode: int, *, flags: int = 0) -> None:
    from Quartz import (  # type: ignore
        CGEventCreateKeyboardEvent,
        CGEventSetFlags,
    )

    lock = get_keyboard_lock()
    lock.allow_synth(True)
    try:
        src = _synth_source()
        down = CGEventCreateKeyboardEvent(src, keycode, True)
        up = CGEventCreateKeyboardEvent(src, keycode, False)
        _mark_synth(down)
        _mark_synth(up)
        CGEventSetFlags(down, flags)
        CGEventSetFlags(up, flags)
        _post_to_outlook(down)
        time.sleep(0.012)
        _post_to_outlook(up)
        time.sleep(0.04)
        if flags:
            _clear_outlook_modifiers()
    finally:
        lock.allow_synth(False)


def _tab_once() -> None:
    """Đúng một Tab thuần tới Outlook. Tap tắt nên không nuốt; Cmd đã thả nên không thành Cmd+Tab."""
    lock = get_keyboard_lock()
    lock.set_enabled(False)
    try:
        _clear_outlook_modifiers()
        _run_nsapplescript(
            '''
tell application "System Events"
  tell process "Microsoft Outlook"
    set frontmost to true
  end tell
  delay 0.03
  key code 48
end tell
'''
        )
        # Tap vẫn tắt cho tới khi Outlook nhận Tab. Bật lại sớm thì tap nuốt đúng phím này.
        time.sleep(0.06)
    finally:
        lock.set_enabled(True)


def _cg_cmd_v() -> None:
    from Quartz import kCGEventFlagMaskCommand  # type: ignore

    _cg_key(9, flags=kCGEventFlagMaskCommand)


def _cg_cmd_return() -> None:
    from Quartz import kCGEventFlagMaskCommand  # type: ignore

    _cg_key(36, flags=kCGEventFlagMaskCommand)


_TAIL_FOCUS = """
    if addrField is not missing value then
      try
        set focused of addrField to true
      end try
      delay 0.04
    end if
    set focused of found to true
    try
      perform action "AXPress" of found
    end try
    try
      set value of attribute "AXSelectedTextRange" of found to {1, 500}
    end try
    delay 0.05
    return "ok"
"""


def _subject_field_script(subject_hint: str, tail: str, *, activate: bool) -> str:
    """Tìm ô Subject. activate=False khi chỉ đọc — tránh activate kéo focus ngược lại."""
    want = _esc_as((subject_hint or "").strip())
    want_short = _esc_as((subject_hint or "").strip()[:56])
    head = 'tell application "Microsoft Outlook" to activate\n' if activate else ""
    front = "    set frontmost to true\n" if activate else ""
    raise_win = """
    try
      perform action "AXRaise" of targetWin
    end try
""" if activate else ""
    return f"""{head}tell application "System Events"
  tell process "Microsoft Outlook"
{front}{_compose_picker_as(subject_hint)}
    if targetWin is missing value then return "no-win"
{raise_win}
    set wantSubj to "{want}"
    set wantShort to "{want_short}"
    set found to missing value
    set addrField to missing value
    set splitG to missing value
    try
      set splitG to item 1 of (UI elements of targetWin)
    end try

    if splitG is not missing value then
      try
        repeat with tf in text fields of splitG
          try
            set dsc to description of tf as text
            if addrField is missing value and (dsc is "To:" or dsc begins with "To:" or dsc is "To") then
              set addrField to tf
            end if
            if found is missing value and (dsc is "Subject:" or dsc begins with "Subject" or dsc contains "Chủ đề") then
              set found to tf
            end if
          end try
        end repeat
      end try

      if found is missing value then
        try
          repeat with tf in text fields of splitG
            try
              set v to value of tf as text
              if v is wantSubj or (wantShort is not "" and v contains wantShort) then
                set found to tf
                exit repeat
              end if
            end try
          end repeat
        end try
      end if

      if found is missing value then
        try
          repeat with gkid in UI elements of splitG
            try
              repeat with tf in text fields of gkid
                try
                  set dsc to description of tf as text
                  if addrField is missing value and (dsc is "To:" or dsc begins with "To:" or dsc is "To") then
                    set addrField to tf
                  end if
                  if found is missing value and (dsc is "Subject:" or dsc begins with "Subject" or dsc contains "Chủ đề") then
                    set found to tf
                  end if
                end try
                try
                  set v to value of tf as text
                  if v is wantSubj or (wantShort is not "" and v contains wantShort) then
                    set found to tf
                    exit repeat
                  end if
                end try
              end repeat
            end try
            if found is not missing value then exit repeat
          end repeat
        end try
      end if
    end if

    if found is missing value then return "no-field"
{tail}
  end tell
end tell
"""


def _focus_subject_via_ax(*, subject_hint: str = "") -> str:
    """
    AX set focused Subject trong AXSplitGroup.
    Trả: ok | no-win | no-field | err:...
    """
    try:
        raw = _run_nsapplescript(
            _subject_field_script(subject_hint, _TAIL_FOCUS, activate=True)
        )
        return (raw or "").strip().lower() or "no-field"
    except Exception as exc:  # noqa: BLE001
        return f"err:{exc}"


def focus_subject_carefully(
    *, subject_hint: str = "", msg_id: str = "", already_raised: bool = False
) -> str:
    """AX Subject. already_raised=True → bỏ wait/claim lại (tránh đợi chồng)."""
    if not already_raised:
        if not wait_compose_window(
            timeout=1.4, subject_hint=subject_hint, msg_id=msg_id
        ):
            raise RuntimeError(
                "Không focus được cửa sổ New Mail (Zalo đang giữ front?).\n"
                "Legacy ON · thử lại."
            )
    ax = _focus_subject_via_ax(subject_hint=subject_hint)
    if ax != "ok":
        claim_outlook_compose(subject_hint=subject_hint, msg_id=msg_id)
        ax = _focus_subject_via_ax(subject_hint=subject_hint)
    if ax != "ok":
        raise RuntimeError(
            f"AX không focus được ô Subject ({ax}).\n"
            "Legacy ON · Accessibility cho VT Rate Mail Sender."
        )
    return "ax"


def paste_and_send_keyboard(
    *,
    subject_hint: str = "",
    msg_id: str = "",
    already_raised: bool = False,
) -> None:
    """To → select Subject → đúng 1 Tab thuần → paste → Cmd+Return."""
    focus_subject_carefully(
        subject_hint=subject_hint, msg_id=msg_id, already_raised=already_raised
    )
    _tab_once()
    _cg_cmd_v()
    time.sleep(0.18)
    _cg_cmd_return()
    time.sleep(0.04)
