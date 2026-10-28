"""Mac HID tối thiểu — quy trình chuẩn (đã probe Legacy):

AS (mac.py): New Mail + To/Cc/Subject  →  caret mặc định ở To
HID (trong KeyboardLock — chặn gõ user):
  1) Raise New Mail (Outlook đã activate từ AS)
  2) AX set focused ô Subject (AXSplitGroup / description «Subject:»)
  3) Tab đúng 1 lần → body
  4) Cmd+V → chờ ngắn → Cmd+Return

Không chuột. Không force Space. Không thử AX Body (Legacy không cho).
"""

from __future__ import annotations

import subprocess
import time

from .mac_keylock import get_keyboard_lock


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
    """Chọn cửa sổ New Mail (title Legacy ≈ subject • account)."""
    full = (subject_hint or "").strip()
    needle = _esc_as(full[:40])
    needle2 = _esc_as(full[:24]) if len(full) >= 24 else needle
    return f'''
    set needle to "{needle}"
    set needle2 to "{needle2}"
    set targetWin to missing value

    try
      set fw to front window
      set fn to name of fw as text
      if needle is not "" and fn contains needle then
        set targetWin to fw
      else if needle2 is not "" and fn contains needle2 then
        set targetWin to fw
      else if (fn contains "Untitled") or (fn contains "@") or (fn contains "New Message") or (fn contains "Outgoing") then
        set targetWin to fw
      else if (length of fn > 24) and (fn does not contain "Inbox") and (fn does not contain "Sent items") and (fn does not contain "Sent Items") and (fn does not contain "Drafts") and (fn does not contain "Deleted") then
        set targetWin to fw
      end if
    end try

    if targetWin is missing value and needle is not "" then
      repeat with w in windows
        try
          set n to name of w as text
          if n contains needle or (needle2 is not "" and n contains needle2) then
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


def activate_outlook(*, settle: float = 0.03) -> None:
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


def wait_compose_window(*, timeout: float = 1.8, subject_hint: str = "") -> bool:
    deadline = time.time() + max(0.2, timeout)
    while time.time() < deadline:
        activate_outlook(settle=0.02)
        if raise_compose_window(subject_hint=subject_hint):
            return True
        time.sleep(0.05)
    return False


def has_compose_window(*, subject_hint: str = "") -> bool:
    return raise_compose_window(subject_hint=subject_hint)


def _cg_key(keycode: int, *, flags: int = 0) -> None:
    from Quartz import (  # type: ignore
        CGEventCreateKeyboardEvent,
        CGEventPost,
        CGEventSetFlags,
        kCGHIDEventTap,
    )

    lock = get_keyboard_lock()
    lock.allow_synth(True)
    try:
        down = CGEventCreateKeyboardEvent(None, keycode, True)
        up = CGEventCreateKeyboardEvent(None, keycode, False)
        if flags:
            CGEventSetFlags(down, flags)
            CGEventSetFlags(up, flags)
        CGEventPost(kCGHIDEventTap, down)
        time.sleep(0.012)
        CGEventPost(kCGHIDEventTap, up)
    finally:
        lock.allow_synth(False)


def _cg_tab() -> None:
    _cg_key(48)


def _cg_escape() -> None:
    _cg_key(53)


def _cg_cmd_v() -> None:
    from Quartz import kCGEventFlagMaskCommand  # type: ignore

    _cg_key(9, flags=kCGEventFlagMaskCommand)


def _cg_cmd_return() -> None:
    from Quartz import kCGEventFlagMaskCommand  # type: ignore

    _cg_key(36, flags=kCGEventFlagMaskCommand)


def _cg_cmd_w() -> None:
    from Quartz import kCGEventFlagMaskCommand  # type: ignore

    _cg_key(13, flags=kCGEventFlagMaskCommand)


def _focus_subject_via_ax(*, subject_hint: str = "") -> str:
    """
    AX set focused Subject trong AXSplitGroup.
    Trả: ok | no-win | no-field | err:...
    """
    want = _esc_as((subject_hint or "").strip())
    want_short = _esc_as((subject_hint or "").strip()[:56])
    script = f'''
tell application "Microsoft Outlook" to activate
tell application "System Events"
  tell process "Microsoft Outlook"
    set frontmost to true
{_compose_picker_as(subject_hint)}
    if targetWin is missing value then return "no-win"
    try
      perform action "AXRaise" of targetWin
    end try

    set wantSubj to "{want}"
    set wantShort to "{want_short}"
    set found to missing value
    set splitG to missing value
    try
      set splitG to item 1 of (UI elements of targetWin)
    end try

    if splitG is not missing value then
      try
        repeat with tf in text fields of splitG
          try
            set dsc to description of tf as text
            if dsc is "Subject:" or dsc begins with "Subject" or dsc contains "Chủ đề" then
              set found to tf
              exit repeat
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
                  if dsc is "Subject:" or dsc begins with "Subject" then
                    set found to tf
                    exit repeat
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

    set focused of found to true
    delay 0.03
    try
      if (focused of found) is true then return "ok"
    end try
    set focused of found to true
    delay 0.03
    return "ok"
  end tell
end tell
'''
    try:
        return _run_nsapplescript(script).strip().lower() or "no-field"
    except Exception as exc:  # noqa: BLE001
        return f"err:{exc}"


def probe_subject_ax(*, subject_hint: str = "") -> str:
    return _focus_subject_via_ax(subject_hint=subject_hint)


def focus_subject_carefully(*, subject_hint: str = "") -> str:
    """Raise New Mail + AX focused Subject (retry 1 lần)."""
    if not wait_compose_window(timeout=1.6, subject_hint=subject_hint):
        raise RuntimeError(
            "Không focus được cửa sổ New Mail.\n"
            "Legacy ON · thử lại."
        )
    ax = _focus_subject_via_ax(subject_hint=subject_hint)
    if ax != "ok":
        activate_outlook(settle=0.03)
        raise_compose_window(subject_hint=subject_hint)
        ax = _focus_subject_via_ax(subject_hint=subject_hint)
    if ax != "ok":
        raise RuntimeError(
            f"AX không focus được ô Subject ({ax}).\n"
            "Legacy ON · Accessibility cho VT Rate Mail Sender."
        )
    time.sleep(0.04)
    return "ax"


def _subject_tab_paste(*, subject_hint: str = "") -> None:
    """Chuẩn: AX Subject → Tab 1 → Cmd+V → chờ ngắn."""
    focus_subject_carefully(subject_hint=subject_hint)
    _cg_tab()
    time.sleep(0.08)
    _cg_cmd_v()
    time.sleep(0.36)


def paste_body_keyboard(*, has_cc: bool = True, subject_hint: str = "") -> None:
    _ = has_cc
    _subject_tab_paste(subject_hint=subject_hint)


def paste_and_send_keyboard(*, has_cc: bool = True, subject_hint: str = "") -> None:
    """AX Subject → Tab 1 → paste → Cmd+Return."""
    _ = has_cc
    _subject_tab_paste(subject_hint=subject_hint)
    _cg_cmd_return()
    time.sleep(0.08)


def send_compose_cmd_return(*, already_in_body: bool = True, subject_hint: str = "") -> None:
    activate_outlook(settle=0.03)
    raise_compose_window(subject_hint=subject_hint)
    if not already_in_body:
        focus_subject_carefully(subject_hint=subject_hint)
        _cg_tab()
        time.sleep(0.08)
    _cg_cmd_return()
    time.sleep(0.08)


def close_front_compose_keyboard() -> None:
    if not raise_compose_window():
        activate_outlook(settle=0.03)
        if not raise_compose_window():
            return
    _cg_cmd_w()
    time.sleep(0.06)
    _cg_escape()
    time.sleep(0.02)


# --- API tương thích ---

def paste_body_via_subject_tab(*, has_cc: bool = True) -> None:
    paste_body_keyboard(has_cc=has_cc)


def paste_and_send_via_tabs(*, has_cc: bool) -> None:
    paste_and_send_keyboard(has_cc=has_cc)


def paste_html_into_compose_body(*, y_ratio: float = 0.36) -> None:
    _ = y_ratio
    paste_body_keyboard(has_cc=True)


def focus_compose_body(*, y_ratio: float = 0.38) -> None:
    _ = y_ratio
    focus_subject_carefully()
    _cg_tab()
    time.sleep(0.1)


def focus_subject_field() -> None:
    focus_subject_carefully()


def click_send_button() -> None:
    send_compose_cmd_return(already_in_body=True)
