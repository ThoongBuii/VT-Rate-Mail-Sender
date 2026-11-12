"""Mac — ghi nội dung, mở New Mail, bấm Send, rồi sang thư sau.

Đã gửi khi Outlook chuyển thư khỏi hộp chờ, không dựa vào cửa sổ còn mở.
Máy đọc Accessibility khác nhau nên cửa sổ dễ báo sai. Chưa xác nhận thì bỏ qua,
không dừng hàng.
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
    # Đuôi subject có tên agency. Đầu subject trùng cả nước, dễ bấm nhầm cửa sổ thư trước.
    needle_end = _esc_as((full[-48:] if len(full) > 24 else full).strip())
    return f'''
    set needle to "{needle}"
    set needle2 to "{needle2}"
    set needleTail to "{needle_tail}"
    set needleEnd to "{needle_end}"
    set targetWin to missing value

    try
      set fw to front window
      set fn to name of fw as text
      if needleEnd is not "" and fn contains needleEnd then
        set targetWin to fw
      else if needleEnd is "" and needle is not "" and fn contains needle then
        set targetWin to fw
      else if needleEnd is "" and needleTail is not "" and fn contains needleTail then
        set targetWin to fw
      end if
    end try

    if targetWin is missing value and needleEnd is not "" then
      repeat with w in windows
        try
          set n to name of w as text
          if n contains needleEnd then
            set targetWin to w
            exit repeat
          end if
        end try
      end repeat
    end if

    if targetWin is missing value and needleEnd is "" then
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

    if targetWin is missing value and needleEnd is "" then
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

    if targetWin is missing value and needleEnd is "" then
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


def _snapshot_pasteboard() -> list[tuple[str, bytes]]:
    from AppKit import NSPasteboard  # type: ignore

    pb = NSPasteboard.generalPasteboard()
    out: list[tuple[str, bytes]] = []
    for ptype in list(pb.types() or []):
        data = pb.dataForType_(ptype)
        if data is None:
            continue
        raw = bytes(data)
        if raw:
            out.append((str(ptype), raw))
    return out


def _place_saved_message(html: str) -> None:
    """Đặt nội dung đã lưu lên clipboard dạng RTF + HTML + chữ.

    WebArchive tự ghép khiến ô soạn bỏ qua Paste. RTF là kiểu body Outlook nhận.
    """
    import html as html_lib
    import re

    from AppKit import (  # type: ignore
        NSAttributedString,
        NSData,
        NSPasteboard,
        NSPasteboardTypeHTML,
        NSPasteboardTypeRTF,
        NSPasteboardTypeString,
    )

    fragment = (html or "").strip()
    if not fragment:
        raise RuntimeError(
            "Chưa có bản copy nội dung thư.\n"
            "Không gửi — cửa sổ New Mail vẫn mở để sửa tay."
        )
    if "<html" not in fragment.lower():
        wrapped = "<!DOCTYPE html><html><body>" + fragment + "</body></html>"
    else:
        wrapped = fragment
    plain = re.sub(r"(?is)<br\s*/?>", "\n", fragment)
    plain = re.sub(r"(?is)<[^>]+>", "", plain)
    plain = html_lib.unescape(plain).strip()

    rtf = None
    try:
        raw = wrapped.encode("utf-8")
        data = NSData.dataWithBytes_length_(raw, len(raw))
        attr, _info, err = NSAttributedString.alloc().initWithData_options_documentAttributes_error_(
            data,
            {"DocumentType": "NSHTML", "CharacterEncoding": 4},
            None,
            None,
        )
        if attr is not None and err is None and attr.length() > 0:
            rtf, _rerr = attr.dataFromRange_documentAttributes_error_(
                (0, attr.length()),
                {"DocumentType": "NSRTF"},
                None,
            )
    except Exception:  # noqa: BLE001
        rtf = None

    types: list[str] = []
    if rtf is not None:
        types.extend([NSPasteboardTypeRTF, "NeXT Rich Text Format v1.0 pasteboard type"])
    types.extend([NSPasteboardTypeHTML, "public.html", "Apple HTML pasteboard type"])
    if plain:
        types.append(NSPasteboardTypeString)
    pb = NSPasteboard.generalPasteboard()
    pb.declareTypes_owner_(types, None)
    if rtf is not None:
        pb.setData_forType_(rtf, NSPasteboardTypeRTF)
        pb.setData_forType_(rtf, "NeXT Rich Text Format v1.0 pasteboard type")
    for ptype in (NSPasteboardTypeHTML, "public.html", "Apple HTML pasteboard type"):
        pb.setString_forType_(wrapped, ptype)
    if plain:
        pb.setString_forType_(plain, NSPasteboardTypeString)


def _js_on_webview(web, source: str, timeout: float = 4.0) -> str:
    from Foundation import NSDate, NSRunLoop  # type: ignore

    box: dict = {"result": "", "done": False}

    def _handler(result, _error) -> None:
        try:
            box["result"] = "" if result is None else str(result)
        except Exception:  # noqa: BLE001
            box["result"] = ""
        box["done"] = True

    web.evaluateJavaScript_completionHandler_(source, _handler)
    deadline = time.time() + timeout
    while not box["done"] and time.time() < deadline:
        NSRunLoop.currentRunLoop().runUntilDate_(NSDate.dateWithTimeIntervalSinceNow_(0.05))
    return str(box["result"] or "")


def _clip_matches(html: str, plain: str) -> bool:
    from .html_merge import html_visible_text

    mark = _compact(html_visible_text(html))[:24]
    if len(mark) < 8:
        return False
    return mark in _compact(plain)


def _copy_from_app_window(html: str) -> list[tuple[str, bytes]] | None:
    """Cmd+C ngay trên cửa sổ app. Cửa sổ WebKit thứ hai không nhận copy khi app đang mở."""
    import json

    try:
        import webview
        from webview.platforms.cocoa import BrowserView  # type: ignore
    except Exception:  # noqa: BLE001
        return None
    if not getattr(webview, "windows", None):
        return None
    browser = BrowserView.instances.get(webview.windows[0].uid)
    if browser is None:
        return None

    from AppKit import NSPasteboard  # type: ignore
    from Foundation import NSDate, NSRunLoop  # type: ignore

    web = browser.webview
    payload = json.dumps(html or "")
    select = (
        "(function(){"
        f"var html={payload};"
        "var el=document.getElementById('vt-copy-once');"
        "if(!el){el=document.createElement('div');el.id='vt-copy-once';"
        "el.contentEditable='true';"
        "el.style.cssText='position:absolute;left:-12000px;top:0;width:960px;';"
        "document.body.appendChild(el);}"
        "el.innerHTML=html;"
        "var sel=window.getSelection();"
        "var range=document.createRange();"
        "range.selectNodeContents(el);"
        "sel.removeAllRanges();sel.addRange(range);"
        "return String((el.innerText||'').length);"
        "})()"
    )
    try:
        browser.window.makeFirstResponder_(web)
        length = _js_on_webview(web, select)
        if not length.isdigit() or int(length) < 8:
            return None
        web.copy_(None)
        NSRunLoop.currentRunLoop().runUntilDate_(NSDate.dateWithTimeIntervalSinceNow_(0.08))
        plain = NSPasteboard.generalPasteboard().stringForType_("public.utf8-plain-text") or ""
        clip = _snapshot_pasteboard() if _clip_matches(html, plain) else None
        _js_on_webview(
            web,
            "var el=document.getElementById('vt-copy-once');"
            "if(el)el.innerHTML='';"
            "var sel=window.getSelection();if(sel)sel.removeAllRanges();'ok'",
        )
    except Exception:  # noqa: BLE001
        return None
    return clip or None


def _copy_rendered_message(html: str) -> list[tuple[str, bytes]]:
    """Một bản copy của thư. Copy trên app không ra thì ghi clipboard HTML, không cắt lượt gửi."""
    fragment = (html or "").strip()
    if not fragment:
        raise RuntimeError("Nội dung thư trống — không gửi.")
    clip = _copy_from_app_window(fragment)
    if clip:
        return clip
    return _html_clip(fragment)


def _html_clip(html: str) -> list[tuple[str, bytes]]:
    """Clipboard HTML + chữ thường. Không gắn WebArchive tự tạo — Outlook hay bỏ qua paste đó."""
    import re

    import html as html_lib

    from AppKit import (  # type: ignore
        NSPasteboard,
        NSPasteboardTypeHTML,
        NSPasteboardTypeString,
    )

    fragment = (html or "").strip()
    if "<html" not in fragment.lower():
        wrapped = "<!DOCTYPE html><html><body>" + fragment + "</body></html>"
    else:
        wrapped = fragment
    plain = re.sub(r"(?is)<br\s*/?>", "\n", fragment)
    plain = re.sub(r"(?is)<[^>]+>", "", plain)
    plain = html_lib.unescape(plain).strip()
    types = [NSPasteboardTypeHTML, "public.html", "Apple HTML pasteboard type"]
    if plain:
        types.append(NSPasteboardTypeString)
    pb = NSPasteboard.generalPasteboard()
    pb.declareTypes_owner_(types, None)
    wrote = False
    for ptype in (NSPasteboardTypeHTML, "public.html", "Apple HTML pasteboard type"):
        try:
            if pb.setString_forType_(wrapped, ptype):
                wrote = True
        except Exception:  # noqa: BLE001
            continue
    if plain:
        try:
            pb.setString_forType_(plain, NSPasteboardTypeString)
        except Exception:  # noqa: BLE001
            pass
    clip = _snapshot_pasteboard()
    if not wrote or not clip:
        raise RuntimeError("Không copy được nội dung thư trên app.")
    return clip


def capture_message_clip(html: str) -> list[tuple[str, bytes]]:
    """Lấy một bản copy của thư và lưu đến khi list xong hoặc Pause rồi sửa nội dung."""
    import hashlib

    from ..message_hold import cached_clip, store_clip

    key = hashlib.sha256((html or "").encode("utf-8")).hexdigest()
    cached = cached_clip(key)
    if cached:
        return cached

    if threading.current_thread() is threading.main_thread():
        clip = _copy_rendered_message(html)
    else:
        from Foundation import NSOperationQueue  # type: ignore

        box: dict = {"err": None, "clip": None}
        done = threading.Event()

        def _block() -> None:
            try:
                box["clip"] = _copy_rendered_message(html)
            except Exception as exc:  # noqa: BLE001
                box["err"] = exc
            finally:
                done.set()

        NSOperationQueue.mainQueue().addOperationWithBlock_(_block)
        if not done.wait(timeout=8) or box["err"] is not None or not box["clip"]:
            clip = _html_clip(html)
        else:
            clip = box["clip"]

    store_clip(key, clip)
    return clip


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


def _paste_menu_once() -> str:
    """Edit → Paste khi lệnh đã bật. Không dùng phím V — Simple Telex nuốt Cmd+V."""
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
            return "disabled"
          end try
        end repeat
      end try
    end repeat
    return "no-menu"
  end tell
end tell
'''
    try:
        return (_run_nsapplescript(script) or "").strip().lower() or "no-menu"
    except Exception:  # noqa: BLE001
        return "no-menu"


def _paste_when_ready() -> str:
    """Đợi Paste bật rồi bấm. Chưa bật thì Cmd+V khi Outlook đang trước mặt."""
    if not _outlook_is_frontmost():
        try:
            _run_nsapplescript(
                'tell application "System Events" to tell process "Microsoft Outlook" '
                "to set frontmost to true"
            )
        except Exception:  # noqa: BLE001
            pass
    last = "no-menu"
    deadline = time.time() + 1.0
    while time.time() < deadline:
        last = _paste_menu_once()
        if last == "yes":
            return last
        time.sleep(0.12)
    _paste_keystroke()
    return last


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


def _write_outgoing_content(msg_id: str, html: str) -> None:
    """Ghi HTML đã soạn vào content. File tạm xoá ngay — không để lại bản copy."""
    import os
    import tempfile

    mid = (msg_id or "").strip()
    fragment = (html or "").strip()
    if not mid.isdigit() or not fragment:
        raise RuntimeError("Không ghi được nội dung thư vào Outlook.")
    fd, path = tempfile.mkstemp(prefix="vtmail_", suffix=".html")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            handle.write(fragment)
        script = f'''
tell application "Microsoft Outlook"
  set htmlText to read (POSIX file "{_esc_as(path)}") as «class utf8»
  set content of outgoing message id {mid} to htmlText
end tell
'''
        _run_nsapplescript(script)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def _delete_outgoing(msg_id: str) -> None:
    mid = (msg_id or "").strip()
    if not mid.isdigit():
        return
    script = f'''
tell application "Microsoft Outlook"
  try
    delete outgoing message id {mid}
  end try
end tell
'''
    try:
        _run_nsapplescript(script)
    except Exception:  # noqa: BLE001
        return


def _content_has_body(msg_id: str, marker: str) -> str:
    """yes = content đã có đoạn thư. no = chưa có. unknown = không đọc được."""
    mark = _norm_plain(marker)
    plain = _outgoing_plain(msg_id)
    if plain is None:
        return "unknown"
    if len(mark) < 8:
        return "unknown"
    if mark[:24] in _norm_plain(plain):
        return "yes"
    return "no"


def _click_send_button(subject_hint: str) -> bool:
    """Bấm nút Send trên đúng cửa sổ New Mail. Không đi qua lệnh AppleScript send."""
    script = f'''
tell application "System Events"
  tell process "Microsoft Outlook"
    set frontmost to true
{_compose_picker_as(subject_hint)}
    if targetWin is missing value then return "no-win"
    try
      perform action "AXRaise" of targetWin
    end try
    set spots to {{targetWin}}
    try
      set spots to spots & toolbars of targetWin
    end try
    try
      set spots to spots & groups of targetWin
    end try
    repeat with spot in spots
      try
        repeat with b in buttons of spot
          set nm to ""
          set dsc to ""
          try
            set nm to name of b as text
          end try
          try
            set dsc to description of b as text
          end try
          if nm is "Send" or nm is "Gửi" or dsc is "Send" or dsc is "Gửi" then
            click b
            return "yes"
          end if
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


def _outgoing_dispatch(msg_id: str) -> str:
    """sent = Outlook đã nhận gửi. gone = thư không còn trong hộp chờ. pending = còn nháp."""
    mid = (msg_id or "").strip()
    if not mid.isdigit():
        return "unknown"
    script = f'''
tell application "Microsoft Outlook"
  try
    if not (exists outgoing message id {mid}) then return "gone"
    set m to outgoing message id {mid}
    try
      if was sent of m is true then return "sent"
    end try
    return "pending"
  on error
    return "unknown"
  end try
end tell
'''
    try:
        state = (_run_nsapplescript(script) or "").strip().lower()
    except Exception:  # noqa: BLE001
        return "unknown"
    if state in {"sent", "gone", "pending", "unknown"}:
        return state
    return "unknown"


def _raise_compose_window(subject_hint: str) -> bool:
    """Đưa đúng New Mail của thư này lên trước. Chưa mở xong thì trả no."""
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


def _press_send_shortcut() -> None:
    """Cmd+Return thật, cùng phím với nút Send. Không bắn phím giả vào PID."""
    script = '''
tell application "System Events"
  tell process "Microsoft Outlook"
    set frontmost to true
    key code 36 using command down
  end tell
end tell
'''
    try:
        _run_nsapplescript(script)
    except Exception:  # noqa: BLE001
        _cg_cmd_return()
        _release_modifiers_quiet()


def _dispatch_done(msg_id: str) -> bool:
    return _outgoing_dispatch(msg_id) in {"sent", "gone"}


def _click_front_send() -> None:
    """Bấm nút Send trên cửa sổ đang trước mặt, nếu nút đó có."""
    script = '''
tell application "System Events"
  tell process "Microsoft Outlook"
    set frontmost to true
    try
      set w to front window
      set spots to {w}
      try
        set spots to spots & toolbars of w
      end try
      try
        set spots to spots & groups of w
      end try
      repeat with spot in spots
        try
          repeat with b in buttons of spot
            set nm to ""
            set dsc to ""
            try
              set nm to name of b as text
            end try
            try
              set dsc to description of b as text
            end try
            if nm is "Send" or nm is "Gửi" or dsc is "Send" or dsc is "Gửi" then
              click b
              return "yes"
            end if
          end repeat
        end try
      end repeat
    end try
    return "no"
  end tell
end tell
'''
    try:
        _run_nsapplescript(script)
    except Exception:  # noqa: BLE001
        return


def send_outgoing(msg_id: str, *, subject_hint: str = "") -> None:
    """Mở đúng thư rồi Cmd+Return. Không cần đọc được tên cửa sổ."""
    del subject_hint
    mid = (msg_id or "").strip()
    if not mid.isdigit():
        raise RuntimeError("Không gửi được thư này.")
    if not open_compose_by_id(mid):
        raise RuntimeError("Không mở được cửa sổ New Mail.")
    time.sleep(0.35)
    for _ in range(2):
        _press_send_shortcut()
        _click_front_send()
        deadline = time.time() + 1.2
        while time.time() < deadline:
            if _dispatch_done(mid):
                return
            time.sleep(0.3)
    raise RuntimeError("Outlook chưa chuyển thư này sang Sent.")
