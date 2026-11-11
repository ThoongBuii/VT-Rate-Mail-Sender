"""Mac HID tối thiểu — quy trình chuẩn (đã probe Legacy):

AS (mac.py): New Mail + To/Cc/Subject  →  caret mặc định ở To
HID (trong KeyboardLock — chặn gõ user / Zalo):
  1) open(msg id) + buộc Outlook frontmost
  2) AX set focused ô Subject
  3) Copy nội dung thư bằng WebKit (như Cmd+C) → Tab 1 lần → body → Cmd+V thật
     → gửi khi nội dung đã vào thư

Ghi HTML tay lên clipboard thiếu RTF / WebKit pasteboard, ô soạn bỏ qua.
Cmd+V bắn bằng CGEventPostToPid không kích hoạt Paste. Tắt tap trong lúc bấm.

Tab/Cmd gửi vào đúng PID Outlook đang sống. Outlook tắt rồi mở lại thì lấy PID mới.
Chưa rời Subject thì không dán và không gửi. Không chuột. Không thử AX Body.
"""

from __future__ import annotations

import subprocess
import threading
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


def _synth_source():
    """Source mới mỗi lần bắn phím — Cmd của mail trước không dính thành Cmd+Tab."""
    try:
        from Quartz import (  # type: ignore
            CGEventSourceCreate,
            CGEventSourceSetUserData,
            kCGEventSourceStatePrivate,
        )

        src = CGEventSourceCreate(kCGEventSourceStatePrivate)
        if src is not None:
            CGEventSourceSetUserData(src, SYNTH_USERDATA)
        return src
    except Exception:  # noqa: BLE001
        return None


def _mark_synth(event) -> None:
    try:
        from Quartz import CGEventSetIntegerValueField, kCGEventSourceUserData  # type: ignore

        CGEventSetIntegerValueField(event, kCGEventSourceUserData, SYNTH_USERDATA)
    except Exception:  # noqa: BLE001
        pass


_OUTLOOK_PID = 0


def _is_outlook_app(app) -> bool:
    bid = str(app.bundleIdentifier() or "")
    name = str(app.localizedName() or "")
    return bid == "com.microsoft.Outlook" or name == "Microsoft Outlook"


def _outlook_pid() -> int:
    """PID Outlook đang sống. Tắt Outlook rồi mở lại thì bỏ PID cũ."""
    global _OUTLOOK_PID
    try:
        from AppKit import NSRunningApplication, NSWorkspace  # type: ignore

        if _OUTLOOK_PID > 0:
            current = NSRunningApplication.runningApplicationWithProcessIdentifier_(
                _OUTLOOK_PID
            )
            if current is not None and _is_outlook_app(current):
                return _OUTLOOK_PID
            _OUTLOOK_PID = 0
        for app in NSWorkspace.sharedWorkspace().runningApplications():
            if _is_outlook_app(app):
                _OUTLOOK_PID = int(app.processIdentifier())
                return _OUTLOOK_PID
    except Exception:  # noqa: BLE001
        _OUTLOOK_PID = 0
    return 0


def _post_to_outlook(event, *, required: bool = True) -> None:
    """Đưa phím thẳng vào Outlook đang chạy. Không bắn ra app khác."""
    from Quartz import CGEventPostToPid  # type: ignore

    pid = _outlook_pid()
    if pid > 0:
        CGEventPostToPid(pid, event)
        return
    if required:
        raise RuntimeError(
            "Outlook không còn là tiến trình app đang giữ.\n"
            "Không gửi — tránh mail không có nội dung."
        )


def _release_modifiers_quiet() -> None:
    """Thả modifier sau khi gửi xong. Không bắn phím vào ô soạn."""
    from Quartz import (  # type: ignore
        CGEventCreate,
        CGEventSetFlags,
        CGEventSetType,
        kCGEventFlagsChanged,
    )

    ev = CGEventCreate(_synth_source())
    if ev is None:
        return
    CGEventSetType(ev, kCGEventFlagsChanged)
    CGEventSetFlags(ev, 0)
    _post_to_outlook(ev, required=False)


def _cg_key(keycode: int, *, flags: int = 0) -> None:
    from Quartz import CGEventCreateKeyboardEvent, CGEventSetFlags  # type: ignore

    lock = get_keyboard_lock()
    lock.set_enabled(True)
    lock.allow_synth(True)
    try:
        src = _synth_source()
        down = CGEventCreateKeyboardEvent(src, keycode, True)
        up = CGEventCreateKeyboardEvent(src, keycode, False)
        _mark_synth(down)
        _mark_synth(up)
        # Tab luôn flags 0. Không kế thừa Cmd còn dính sau nhiều mail.
        CGEventSetFlags(down, flags)
        CGEventSetFlags(up, flags)
        _post_to_outlook(down)
        time.sleep(0.012)
        _post_to_outlook(up)
    finally:
        lock.allow_synth(False)


def _cg_tab() -> None:
    _cg_key(48)


def _outlook_is_frontmost() -> bool:
    try:
        from AppKit import NSWorkspace  # type: ignore

        app = NSWorkspace.sharedWorkspace().frontmostApplication()
        return app is not None and _is_outlook_app(app)
    except Exception:  # noqa: BLE001
        return False


def _hid_cmd_v() -> None:
    """Cmd+V mức HID. Chỉ khi Outlook đang là app trước mặt."""
    from Quartz import (  # type: ignore
        CGEventCreateKeyboardEvent,
        CGEventPost,
        CGEventSetFlags,
        kCGEventFlagMaskCommand,
        kCGHIDEventTap,
    )

    if not _outlook_is_frontmost():
        return
    src = _synth_source()
    down = CGEventCreateKeyboardEvent(src, 9, True)
    up = CGEventCreateKeyboardEvent(src, 9, False)
    if down is None or up is None:
        return
    _mark_synth(down)
    _mark_synth(up)
    CGEventSetFlags(down, kCGEventFlagMaskCommand)
    CGEventSetFlags(up, kCGEventFlagMaskCommand)
    CGEventPost(kCGHIDEventTap, down)
    time.sleep(0.012)
    CGEventPost(kCGHIDEventTap, up)


def _compact(text: str | None) -> str:
    return "".join((text or "").split()).lower()


def _copy_rendered_message(html: str) -> None:
    """Copy đúng thao tác: mở nội dung thư trong WebKit, chọn hết, Cmd+C.

    Phải chạy trên main thread. Cửa sổ trong suốt, ngoài màn hình, không chiếm Outlook.
    """
    from AppKit import (  # type: ignore
        NSBackingStoreBuffered,
        NSMakeRect,
        NSPasteboard,
        NSWindow,
    )
    from Foundation import NSDate, NSRunLoop  # type: ignore
    from WebKit import WKWebView, WKWebViewConfiguration  # type: ignore

    from .html_merge import html_visible_text

    fragment = (html or "").strip()
    if not fragment:
        raise RuntimeError(
            "Không copy được nội dung thư trên app.\n"
            "Không gửi — cửa sổ New Mail vẫn mở để sửa tay."
        )
    if "<html" not in fragment.lower():
        page = "<!DOCTYPE html><html><body>" + fragment + "</body></html>"
    else:
        page = fragment

    web = WKWebView.alloc().initWithFrame_configuration_(
        NSMakeRect(0, 0, 960, 1200),
        WKWebViewConfiguration.alloc().init(),
    )
    win = NSWindow.alloc().initWithContentRect_styleMask_backing_defer_(
        NSMakeRect(-3200, -3200, 960, 1200),
        0,
        NSBackingStoreBuffered,
        False,
    )
    win.setReleasedWhenClosed_(False)
    win.setAlphaValue_(0)
    win.setIgnoresMouseEvents_(True)
    try:
        win.contentView().addSubview_(web)
        web.loadHTMLString_baseURL_(page, None)
        deadline = time.time() + 3
        while web.isLoading() and time.time() < deadline:
            NSRunLoop.currentRunLoop().runUntilDate_(
                NSDate.dateWithTimeIntervalSinceNow_(0.05)
            )
        win.orderFrontRegardless()
        win.makeFirstResponder_(web)
        NSRunLoop.currentRunLoop().runUntilDate_(NSDate.dateWithTimeIntervalSinceNow_(0.08))
        web.selectAll_(None)
        web.copy_(None)
        NSRunLoop.currentRunLoop().runUntilDate_(NSDate.dateWithTimeIntervalSinceNow_(0.08))
    finally:
        try:
            win.orderOut_(None)
        except Exception:  # noqa: BLE001
            pass
        try:
            web.removeFromSuperview()
        except Exception:  # noqa: BLE001
            pass

    plain = NSPasteboard.generalPasteboard().stringForType_("public.utf8-plain-text") or ""
    mark = _compact(html_visible_text(html))[:32]
    if len(mark) < 8 or mark not in _compact(plain):
        raise RuntimeError(
            "Không copy được nội dung thư trên app.\n"
            "Không gửi — cửa sổ New Mail vẫn mở để sửa tay."
        )


def _write_clipboard(html: str) -> None:
    """Copy nội dung thư trên main thread. Outlook chỉ nhận pasteboard do WebKit tạo."""
    if threading.current_thread() is threading.main_thread():
        _copy_rendered_message(html)
        return
    from Foundation import NSOperationQueue  # type: ignore

    box: dict = {"err": None}
    done = threading.Event()

    def _block() -> None:
        try:
            _copy_rendered_message(html)
        except Exception as exc:  # noqa: BLE001
            box["err"] = exc
        finally:
            done.set()

    NSOperationQueue.mainQueue().addOperationWithBlock_(_block)
    if not done.wait(timeout=6):
        raise RuntimeError(
            "Không copy được nội dung thư trên app.\n"
            "Không gửi — cửa sổ New Mail vẫn mở để sửa tay."
        )
    if box["err"] is not None:
        raise box["err"]


def _paste_keystroke() -> None:
    """Cmd+V thật vào ô đang focus.

    Tắt tap trong lúc bấm. WebView không coi Cmd+V từ CGEventPostToPid là Paste,
    còn keystroke System Events chạy qua session tap và bị nuốt.
    """
    lock = get_keyboard_lock()
    lock.set_enabled(False)
    try:
        if not _outlook_is_frontmost():
            return
        _run_nsapplescript(
            'tell application "System Events" to key code 9 using command down'
        )
        time.sleep(0.1)
    except Exception:  # noqa: BLE001
        _hid_cmd_v()
    finally:
        lock.set_enabled(True)


def _paste_menu() -> bool:
    """Edit → Paste. Lệnh menu, không đi qua tap phím."""
    script = '''
tell application "System Events"
  tell process "Microsoft Outlook"
    repeat with en in {"Edit", "Sửa", "Chỉnh sửa"}
      try
        set m to menu en of menu bar item en of menu bar 1
        repeat with pn in {"Paste", "Dán"}
          try
            set mi to menu item pn of m
            if enabled of mi is true then
              click mi
              return "yes"
            end if
          end try
        end repeat
      end try
    end repeat
    return "no"
  end tell
end tell
'''
    try:
        return _run_nsapplescript(script).strip().lower().startswith("yes")
    except Exception:  # noqa: BLE001
        return False


def _body_marker(html: str) -> str:
    from .html_merge import html_visible_text

    return " ".join(html_visible_text(html).split())[:48]


def _outgoing_plain(msg_id: str) -> str | None:
    """Plain text của New Mail. None = không đọc được."""
    mid = (msg_id or "").strip()
    if not mid.isdigit():
        return None
    script = f'''
tell application "Microsoft Outlook"
  try
    return content of outgoing message id {mid}
  on error
    return ""
  end try
end tell
'''
    try:
        return _run_nsapplescript(script) or ""
    except Exception:  # noqa: BLE001
        return None


def _norm_plain(text: str | None) -> str:
    return " ".join((text or "").split()).lower()


def _body_landed(msg_id: str, marker: str, before: str | None = None) -> str:
    """yes = nội dung mới đã vào. no = thư không đổi. unknown = không đọc được.

    Chữ ký có sẵn không tính là đã dán. Đoạn mở đầu phải mới xuất hiện, hoặc thân thư dài thêm.
    """
    mark = _norm_plain(marker)
    plain = _outgoing_plain(msg_id)
    if plain is None:
        return "unknown"
    got = _norm_plain(plain)
    prev = _norm_plain(before) if before is not None else None
    grew = prev is not None and len(got) > len(prev) + 12
    if len(mark) < 8:
        return "yes" if grew else "unknown"
    fresh = mark[:24] in got and (prev is None or mark[:24] not in prev)
    if fresh or grew:
        return "yes"
    return "no"


def _wait_landed(msg_id: str, marker: str, before: str | None) -> str:
    """Đọc hai lần. WebView có thể cập nhật content chậm hơn thao tác Paste."""
    last = "no"
    for pause in (0.28, 0.35):
        time.sleep(pause)
        last = _body_landed(msg_id, marker, before)
        if last != "no":
            return last
    return last


def _cg_cmd_return() -> None:
    from Quartz import kCGEventFlagMaskCommand  # type: ignore

    _cg_key(36, flags=kCGEventFlagMaskCommand)


_TAIL_FOCUS = """
    set focused of found to true
    delay 0.03
    try
      if (focused of found) is true then return "ok"
    end try
    set focused of found to true
    delay 0.03
    return "ok"
"""

_TAIL_READ_FOCUS = """
    try
      if (focused of found) is true then return "subject"
    end try
    return "moved"
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


def _read_subject_focus(subject_hint: str) -> str:
    """subject = còn ở Subject. moved = đã rời. unknown = không đọc được."""
    try:
        raw = _run_nsapplescript(
            _subject_field_script(subject_hint, _TAIL_READ_FOCUS, activate=False)
        )
        state = (raw or "").strip().lower()
    except Exception:  # noqa: BLE001
        return "unknown"
    if state in {"subject", "moved"}:
        return state
    return "unknown"


def paste_and_send_keyboard(
    *,
    subject_hint: str = "",
    msg_id: str = "",
    already_raised: bool = False,
    html: str = "",
) -> None:
    """Subject → Tab 1 → ghi clipboard → Paste vào body → gửi khi nội dung đã vào thư."""
    if _outlook_pid() <= 0:
        raise RuntimeError(
            "Outlook không đang chạy cùng app.\n"
            "Không gửi — tránh mail không có nội dung."
        )
    focus_subject_carefully(
        subject_hint=subject_hint, msg_id=msg_id, already_raised=already_raised
    )
    _cg_tab()
    time.sleep(0.04)
    where = _read_subject_focus(subject_hint)
    if where == "subject":
        time.sleep(0.12)
        where = _read_subject_focus(subject_hint)
    if where != "moved":
        raise RuntimeError(
            "Chưa rời ô Subject nên không dán nội dung.\n"
            "Không gửi — cửa sổ New Mail vẫn mở để sửa tay."
        )

    before = _outgoing_plain(msg_id)
    marker = _body_marker(html)
    _write_clipboard(html)
    _paste_keystroke()
    landed = _wait_landed(msg_id, marker, before)
    if landed == "no":
        # Clipboard vẫn là bản WebKit vừa copy. Copy lại sẽ kéo focus khỏi Outlook.
        _paste_menu()
        time.sleep(0.4)
        landed = _body_landed(msg_id, marker, before)
    if landed == "no":
        raise RuntimeError(
            "Đã ở body nhưng nội dung soạn chưa vào thư.\n"
            "Không gửi — cửa sổ New Mail vẫn mở để sửa tay."
        )
    _cg_cmd_return()
    time.sleep(0.05)
    _release_modifiers_quiet()
