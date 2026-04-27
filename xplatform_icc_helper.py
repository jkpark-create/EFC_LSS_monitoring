from __future__ import annotations

import argparse
import datetime as dt
import getpass
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


DEFAULT_XPLATFORM_EXE = Path(r"C:\Program Files (x86)\TOBESOFT\XPLATFORM\9.2\XPlatform.exe")
DEFAULT_XPLATFORM_KEY = "KMTC"
DEFAULT_XPLATFORM_XADL = "http://iccv2.kmtc.co.kr/icc/KMTC.xadl"
DEFAULT_LOG_DIR = Path("logs")
DEFAULT_DOCUMENT_NAME = "[\uc601\uc5c5\ud300] LSS & EFC \uc9d5\uc218\uae08\uc561\uc870\ud68c"
DEFAULT_DOWNLOAD_DIR = Path("downloads")
BASE_WINDOW_SIZE = (1280, 728)
LOGIN_WINDOW_SIZE = (648, 368)
CREDENTIAL_TARGET = "EFC_LSS_ICC_XPLATFORM"
FATAL_DIALOG_TITLE_TOKENS = (
    "XPlatform.exe",
    "\uc751\uc6a9 \ud504\ub85c\uadf8\ub7a8 \uc624\ub958",
    "application error",
)
BUSY_MODAL_AREA_RANGE = (10_000, 120_000)
BUSY_MODAL_HEIGHT_RANGE = (60, 150)
BUSY_MODAL_ASPECT_RATIO_MIN = 2.5


@dataclass(frozen=True)
class WindowInfo:
    backend: str
    handle: int
    process_id: int | None
    title: str
    class_name: str
    rectangle: str
    wrapper: object


def load_desktop(backend: str):
    try:
        from pywinauto import Desktop
    except ImportError as exc:
        raise RuntimeError("pywinauto is required. Run: py -m pip install -r requirements.txt") from exc

    return Desktop(backend=backend)


def process_id(wrapper: object) -> int | None:
    getter = getattr(wrapper, "process_id", None)
    if not callable(getter):
        return None
    try:
        return int(getter())
    except Exception:
        return None


def window_title(wrapper: object) -> str:
    getter = getattr(wrapper, "window_text", None)
    if callable(getter):
        try:
            return getter() or ""
        except Exception:
            return ""
    return ""


def window_class(wrapper: object) -> str:
    getter = getattr(wrapper, "class_name", None)
    if callable(getter):
        try:
            return getter() or ""
        except Exception:
            return ""
    return ""


def window_rectangle(wrapper: object) -> str:
    getter = getattr(wrapper, "rectangle", None)
    if callable(getter):
        try:
            return str(getter())
        except Exception:
            return ""
    return ""


def window_handle(wrapper: object) -> int:
    handle = getattr(wrapper, "handle", 0)
    try:
        return int(handle or 0)
    except Exception:
        return 0


def looks_like_xplatform(title: str, class_name: str) -> bool:
    if class_name == "CyWindowClass":
        return True
    if title in {"Login Form"} or title.startswith("KMTC :: ICC"):
        return True
    return any(token.lower() in title.lower() for token in FATAL_DIALOG_TITLE_TOKENS)


def get_window_text_safe(hwnd: int, timeout_ms: int = 50) -> str:
    import ctypes
    from ctypes import wintypes
    import win32con
    user32 = ctypes.windll.user32
    length = wintypes.DWORD()
    res = user32.SendMessageTimeoutW(
        hwnd, win32con.WM_GETTEXTLENGTH, 0, 0, win32con.SMTO_ABORTIFHUNG, timeout_ms, ctypes.byref(length)
    )
    if res == 0 or length.value == 0:
        return ""
    buf = ctypes.create_unicode_buffer(length.value + 1)
    res = user32.SendMessageTimeoutW(
        hwnd, win32con.WM_GETTEXT, length.value + 1, wintypes.LPARAM(ctypes.addressof(buf)), win32con.SMTO_ABORTIFHUNG, timeout_ms, ctypes.byref(length)
    )
    if res == 0:
        return ""
    return buf.value


def collect_windows(backends: Iterable[str] = ("uia", "win32")) -> list[WindowInfo]:
    import win32gui
    import win32process
    windows: list[WindowInfo] = []

    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd):
            class_name = win32gui.GetClassName(hwnd)
            title = get_window_text_safe(hwnd)
            if not looks_like_xplatform(title, class_name):
                return True
                
            rect = win32gui.GetWindowRect(hwnd)
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            
            windows.append(
                WindowInfo(
                    backend="win32",
                    handle=hwnd,
                    process_id=pid,
                    title=title,
                    class_name=class_name,
                    rectangle=str(rect),
                    wrapper=None,
                )
            )
        return True

    win32gui.EnumWindows(callback, None)
    return windows


def visible_area(info: WindowInfo) -> int:
    bounds = window_bounds_from_rectangle(info.rectangle)
    if bounds is None:
        return 0
    left, top, right, bottom = bounds
    return max(0, right - left) * max(0, bottom - top)


def window_bounds_from_rectangle(rect: str) -> tuple[int, int, int, int] | None:
    numbers: list[int] = []
    for token in rect.replace("(", " ").replace(")", " ").replace(",", " ").split():
        if len(token) > 1 and token[0] in {"L", "T", "R", "B"}:
            token = token[1:]
        try:
            numbers.append(int(token))
        except ValueError:
            pass

    if len(numbers) < 4:
        return None
    return tuple(numbers[:4])


def best_capture_window(windows: list[WindowInfo]) -> WindowInfo | None:
    visible = [info for info in windows if visible_area(info) > 0]
    if not visible:
        return None

    login = [info for info in visible if info.title == "Login Form" and info.backend == "uia"]
    if login:
        return login[0]

    app = [info for info in visible if info.title.startswith("KMTC :: ICC")]
    if app:
        return max(app, key=visible_area)

    cy = [info for info in visible if info.class_name == "CyWindowClass"]
    if cy:
        return max(cy, key=visible_area)

    return max(visible, key=visible_area)


def print_windows(windows: list[WindowInfo]) -> None:
    if not windows:
        print("No XPlatform windows were found.")
        return

    for info in windows:
        print(
            "backend={backend} pid={pid} handle={handle} class={class_name!r} "
            "title={title!r} rect={rect}".format(
                backend=info.backend,
                pid=info.process_id,
                handle=info.handle,
                class_name=info.class_name,
                title=info.title,
                rect=info.rectangle,
            )
        )


def fatal_error_windows(windows: list[WindowInfo] | None = None) -> list[WindowInfo]:
    matches: list[WindowInfo] = []
    source = windows if windows is not None else collect_windows()
    for info in source:
        text = " ".join(part for part in (info.title, window_text_snapshot(info)) if part)
        if any(token.lower() in text.lower() for token in FATAL_DIALOG_TITLE_TOKENS):
            matches.append(info)
    return matches


def window_text_snapshot(info: WindowInfo) -> str:
    parts = [info.title]
    descendants = getattr(info.wrapper, "descendants", None)
    if callable(descendants):
        try:
            for child in descendants()[:50]:
                text = window_title(child)
                if text:
                    parts.append(text)
        except Exception:
            pass
    return " ".join(part for part in parts if part)


def visible_blank_modal_windows(windows: list[WindowInfo] | None = None) -> list[WindowInfo]:
    modals: list[WindowInfo] = []
    source = windows if windows is not None else collect_windows()
    main_handles = {info.handle for info in source if info.title.startswith("KMTC :: ICC")}
    for info in source:
        if info.handle in main_handles:
            continue
        if info.class_name != "CyWindowClass" or info.title:
            continue
        bounds = window_bounds_from_rectangle(info.rectangle)
        if bounds is None:
            continue
        left, top, right, bottom = bounds
        width = max(0, right - left)
        height = max(0, bottom - top)
        if height <= 0:
            continue
        area = visible_area(info)
        aspect_ratio = width / height
        if (
            BUSY_MODAL_AREA_RANGE[0] <= area <= BUSY_MODAL_AREA_RANGE[1]
            and BUSY_MODAL_HEIGHT_RANGE[0] <= height <= BUSY_MODAL_HEIGHT_RANGE[1]
            and aspect_ratio >= BUSY_MODAL_ASPECT_RATIO_MIN
        ):
            modals.append(info)
    return modals


def dismiss_fatal_error_dialogs(windows: list[WindowInfo] | None = None) -> bool:
    dialogs = fatal_error_windows(windows)
    if not dialogs:
        return False

    from pywinauto import keyboard

    for dialog in dialogs:
        try:
            bring_to_front(dialog)
            keyboard.send_keys("{ENTER}")
            time.sleep(0.5)
        except Exception as exc:
            print(f"Could not dismiss XPlatform error dialog {dialog.handle}: {exc}", file=sys.stderr)
    return True


def terminate_xplatform_processes(windows: list[WindowInfo] | None = None) -> None:
    pids = {
        info.process_id
        for info in (windows if windows is not None else collect_windows())
        if info.process_id and (info.class_name == "CyWindowClass" or "XPlatform" in info.title)
    }
    for pid in sorted(pids):
        try:
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                check=False,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            print(f"Terminated stale XPlatform process pid={pid}.")
        except Exception as exc:
            print(f"Could not terminate XPlatform process pid={pid}: {exc}", file=sys.stderr)


def recover_from_fatal_xplatform_error(prefix: str, output_dir: Path) -> bool:
    windows = collect_windows()
    if not fatal_error_windows(windows):
        return False

    captures = capture_diagnostic_windows(output_dir, prefix)
    for capture in captures:
        print(f"Diagnostic screenshot: {capture.resolve()}")
    dismiss_fatal_error_dialogs(windows)
    terminate_xplatform_processes(windows)
    time.sleep(3)
    return True


def recover_from_stale_loading_modal(prefix: str, output_dir: Path) -> bool:
    windows = collect_windows()
    if fatal_error_windows(windows):
        return recover_from_fatal_xplatform_error(prefix, output_dir)
    if not visible_blank_modal_windows(windows):
        return False

    captures = capture_diagnostic_windows(output_dir, prefix)
    for capture in captures:
        print(f"Diagnostic screenshot: {capture.resolve()}")
    print("XPlatform is blocked by a loading dialog; restarting the client.")
    terminate_xplatform_processes(windows)
    time.sleep(3)
    return True


def try_dismiss_loading_modal_by_click(main: WindowInfo | None) -> None:
    """Find the persistent loading overlay windows and send WM_CLOSE to forcefully dismiss them."""
    if main is None:
        return
    try:
        import win32gui
        import win32con
        handles = get_safe_blank_modal_handles()
        for hwnd in handles:
            print(f"Force closing stuck loading modal (handle: {hwnd})")
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        time.sleep(1)
    except Exception as exc:
        print(f"Failed to force close modal: {exc}")


def wait_for_blank_modals_to_clear(
    timeout: int,
    output_dir: Path,
    prefix: str,
    *,
    raise_on_timeout: bool = True,
    dismiss_interval: int = 30,
) -> bool:
    """Wait until loading modals clear. Returns True if cleared, False if timed out.
    If raise_on_timeout is False, logs a warning and returns False instead of raising.
    Periodically clicks the main window to try to dismiss a stuck loading overlay.
    """
    deadline = time.monotonic() + timeout
    last_dismiss = time.monotonic()
    while time.monotonic() < deadline:
        windows = collect_windows()
        if fatal_error_windows(windows):
            recover_from_fatal_xplatform_error(prefix, output_dir)
            raise RuntimeError("XPlatform showed an application error and was restarted.")
        if not visible_blank_modal_windows(windows):
            return True
        # Periodically click the main window to try to dismiss a stuck modal
        if time.monotonic() - last_dismiss >= dismiss_interval:
            try_dismiss_loading_modal_by_click(main_window(windows))
            last_dismiss = time.monotonic()
        time.sleep(2)

    captures = capture_diagnostic_windows(output_dir, prefix)
    for capture in captures:
        print(f"Diagnostic screenshot: {capture.resolve()}")
    if raise_on_timeout:
        raise RuntimeError("Timed out waiting for XPlatform loading dialog to disappear.")
    print(f"Warning: loading dialog did not clear within {timeout}s; proceeding anyway.", flush=True)
    return False


def sleep_with_xplatform_checks(seconds: int, output_dir: Path, prefix: str) -> None:
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        windows = collect_windows()
        if fatal_error_windows(windows):
            recover_from_fatal_xplatform_error(prefix, output_dir)
            raise RuntimeError("XPlatform showed an application error and was restarted.")
        time.sleep(min(2, max(0.1, deadline - time.monotonic())))


def capture_window(info: WindowInfo, output_dir: Path, prefix: str = "xplatform") -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    safe_title = "".join(char if char.isalnum() else "_" for char in (info.title or "window")).strip("_")
    output = output_dir / f"{prefix}_{safe_title}_{stamp}.png"

    try:
        from PIL import ImageGrab
        bring_to_front(info)
        time.sleep(0.3)
        left, top, right, bottom = window_rect(info)
        image = ImageGrab.grab(bbox=(left, top, right, bottom))
        image.save(output)
    except Exception as exc:
        print(f"Could not capture window via PIL: {exc}", file=sys.stderr)
    return output


def capture_diagnostic_windows(output_dir: Path, prefix: str) -> list[Path]:
    captures: list[Path] = []
    windows = collect_windows()
    for info in windows:
        if visible_area(info) <= 0:
            continue
        try:
            captures.append(capture_window(info, output_dir, prefix=prefix))
        except Exception as exc:
            print(f"Could not capture diagnostic window {info.handle}: {exc}", file=sys.stderr)
    return captures


def main_window(windows: list[WindowInfo] | None = None) -> WindowInfo | None:
    candidates = [
        info
        for info in (windows if windows is not None else collect_windows())
        if info.title.startswith("KMTC :: ICC") and visible_area(info) > 0
    ]
    if not candidates:
        return None
    return max(candidates, key=visible_area)


def login_window(windows: list[WindowInfo] | None = None) -> WindowInfo | None:
    candidates = [
        info
        for info in (windows if windows is not None else collect_windows())
        if info.title == "Login Form" and visible_area(info) > 0
    ]
    if not candidates:
        return None
    return max(candidates, key=visible_area)


def read_stored_credential(target: str = CREDENTIAL_TARGET) -> tuple[str, str] | None:
    try:
        import win32cred
    except ImportError as exc:
        raise RuntimeError("pywin32 is required for Windows Credential Manager support.") from exc

    try:
        credential = win32cred.CredRead(target, win32cred.CRED_TYPE_GENERIC)
    except Exception:
        return None

    username = str(credential.get("UserName") or "")
    blob = credential.get("CredentialBlob") or b""
    if isinstance(blob, bytes):
        password = blob.decode("utf-16-le", errors="ignore")
        if not password:
            password = blob.decode("utf-8", errors="ignore")
    else:
        password = str(blob)
    return username, password


def save_stored_credential(username: str, password: str, target: str = CREDENTIAL_TARGET) -> None:
    try:
        import win32cred
    except ImportError as exc:
        raise RuntimeError("pywin32 is required for Windows Credential Manager support.") from exc

    win32cred.CredWrite(
        {
            "Type": win32cred.CRED_TYPE_GENERIC,
            "TargetName": target,
            "UserName": username,
            "CredentialBlob": password,
            "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
        },
        0,
    )


def delete_stored_credential(target: str = CREDENTIAL_TARGET) -> bool:
    try:
        import win32cred
    except ImportError as exc:
        raise RuntimeError("pywin32 is required for Windows Credential Manager support.") from exc

    try:
        win32cred.CredDelete(target, win32cred.CRED_TYPE_GENERIC)
        return True
    except Exception:
        return False


def credential_save(args: argparse.Namespace) -> None:
    username = args.username or getpass.getuser()
    password = ""
    if args.password_stdin:
        password = sys.stdin.read().rstrip("\r\n")
    elif args.gui:
        try:
            import tkinter as tk
            from tkinter import simpledialog

            root = tk.Tk()
            root.withdraw()
            password = simpledialog.askstring("ICC Password", "Enter ICC password", show="*") or ""
            root.destroy()
        except Exception as exc:
            raise RuntimeError(f"Could not open credential prompt: {exc}") from exc
    else:
        password = getpass.getpass("ICC password: ")

    if not password:
        raise RuntimeError("Password was empty; credential was not saved.")
    save_stored_credential(username, password, args.target)
    print(f"Credential saved for target {args.target!r} and user {username!r}.")


def credential_status(args: argparse.Namespace) -> None:
    credential = read_stored_credential(args.target)
    if credential is None:
        print(f"No credential found for target {args.target!r}.")
        return
    username, password = credential
    print(
        f"Credential found for target {args.target!r}; "
        f"user={username!r}; password_length={len(password)}"
    )


def credential_delete(args: argparse.Namespace) -> None:
    deleted = delete_stored_credential(args.target)
    if deleted:
        print(f"Credential deleted for target {args.target!r}.")
    else:
        print(f"No credential was deleted for target {args.target!r}.")


def bring_to_front(info: WindowInfo) -> None:
    import win32con
    import win32gui

    win32gui.ShowWindow(info.handle, win32con.SW_RESTORE)
    time.sleep(0.2)
    try:
        win32gui.SetWindowPos(
            info.handle,
            win32con.HWND_TOPMOST,
            0,
            0,
            0,
            0,
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE,
        )
        time.sleep(0.1)
        win32gui.SetWindowPos(
            info.handle,
            win32con.HWND_NOTOPMOST,
            0,
            0,
            0,
            0,
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE,
        )
    except Exception:
        pass
    try:
        win32gui.SetForegroundWindow(info.handle)
    except Exception:
        try:
            from pywinauto import keyboard

            keyboard.send_keys("%")
            time.sleep(0.1)
            win32gui.SetForegroundWindow(info.handle)
        except Exception:
            pass
    try:
        win32gui.BringWindowToTop(info.handle)
    except Exception:
        pass
    time.sleep(0.3)


def window_rect(info: WindowInfo) -> tuple[int, int, int, int]:
    import win32gui

    return win32gui.GetWindowRect(info.handle)


def rel_point(info: WindowInfo, rel_x: int, rel_y: int) -> tuple[int, int]:
    return scaled_point(info, rel_x, rel_y, BASE_WINDOW_SIZE)


def scaled_point(info: WindowInfo, rel_x: int, rel_y: int, base_size: tuple[int, int]) -> tuple[int, int]:
    left, top, right, bottom = window_rect(info)
    width = max(1, right - left)
    height = max(1, bottom - top)
    base_width, base_height = base_size
    return (
        left + round(rel_x * width / base_width),
        top + round(rel_y * height / base_height),
    )


def click_rel(info: WindowInfo, rel_x: int, rel_y: int, *, double: bool = False) -> None:
    from pywinauto import mouse

    coords = rel_point(info, rel_x, rel_y)
    if double:
        mouse.double_click(button="left", coords=coords)
    else:
        mouse.click(button="left", coords=coords)
    time.sleep(0.2)


def click_scaled(
    info: WindowInfo,
    rel_x: int,
    rel_y: int,
    base_size: tuple[int, int],
    *,
    double: bool = False,
) -> None:
    from pywinauto import mouse

    coords = scaled_point(info, rel_x, rel_y, base_size)
    if double:
        mouse.double_click(button="left", coords=coords)
    else:
        mouse.click(button="left", coords=coords)
    time.sleep(0.2)


def paste_text(value: str, *, clear: bool = True) -> None:
    import pyperclip
    from pywinauto import keyboard

    try:
        previous_clipboard = pyperclip.paste()
    except Exception:
        previous_clipboard = None

    if clear:
        keyboard.send_keys("^a")
        time.sleep(0.1)
    pyperclip.copy(value)
    keyboard.send_keys("^v")
    time.sleep(0.2)
    if previous_clipboard is not None:
        try:
            pyperclip.copy(previous_clipboard)
        except Exception:
            pass


def set_focused_text(value: str) -> None:
    from pywinauto import keyboard

    paste_text(value)
    time.sleep(0.1)
    # Some XPlatform password boxes ignore clipboard paste on a restored login form.
    # Backspace and paste a second time makes the operation idempotent when the first
    # Ctrl+V landed in the wrong field or was swallowed while focus was settling.
    keyboard.send_keys("^a")
    time.sleep(0.1)
    keyboard.send_keys("{BACKSPACE}")
    time.sleep(0.1)
    paste_text(value, clear=False)


def edit_controls(info: WindowInfo) -> list[object]:
    descendants = getattr(info.wrapper, "descendants", None)
    if not callable(descendants):
        return []

    for kwargs in ({"control_type": "Edit"}, {"class_name": "Edit"}):
        try:
            controls = list(descendants(**kwargs))
        except Exception:
            continue
        if controls:
            return controls
    return []


def set_control_text(control: object, value: str) -> bool:
    try:
        set_focus = getattr(control, "set_focus", None)
        if callable(set_focus):
            set_focus()
            time.sleep(0.2)
    except Exception:
        pass

    for method_name in ("set_edit_text", "set_text"):
        method = getattr(control, method_name, None)
        if not callable(method):
            continue
        try:
            method(value)
            time.sleep(0.2)
            return True
        except Exception:
            pass

    try:
        set_focused_text(value)
        return True
    except Exception:
        return False


def submit_login(login: WindowInfo) -> None:
    from pywinauto import keyboard

    click_scaled(login, 570, 205, LOGIN_WINDOW_SIZE)
    time.sleep(0.5)
    keyboard.send_keys("{ENTER}")
    time.sleep(0.5)


def try_auto_login(args: argparse.Namespace, login: WindowInfo) -> bool:
    credential = read_stored_credential(args.credential_target)
    if credential is None:
        return False

    username, password = credential
    if not password:
        return False

    print("Attempting ICC auto-login with Windows Credential Manager.")
    for attempt in range(1, args.login_auto_attempts + 1):
        fresh_login = login_window() or login
        bring_to_front(fresh_login)

        controls = edit_controls(fresh_login)
        if len(controls) >= 2:
            if username:
                set_control_text(controls[0], username)
            set_control_text(controls[1], password)
        else:
            if username:
                click_scaled(fresh_login, 510, 112, LOGIN_WINDOW_SIZE)
                set_focused_text(username)
            click_scaled(fresh_login, 510, 160, LOGIN_WINDOW_SIZE)
            set_focused_text(password)

        submit_login(fresh_login)

        deadline = time.monotonic() + args.login_after_wait
        while time.monotonic() < deadline:
            windows = collect_windows()
            if login_window(windows) is None and main_window(windows) is not None:
                print("ICC auto-login completed with Windows Credential Manager.")
                return True
            time.sleep(1)

        if attempt < args.login_auto_attempts:
            print(f"ICC auto-login did not finish; retrying ({attempt + 1}/{args.login_auto_attempts}).")

    captures = capture_diagnostic_windows(Path(args.output_dir), "xplatform_login_fail")
    for capture in captures:
        print(f"Login diagnostic screenshot: {capture.resolve()}")
    return login_window() is None


def ensure_main_window(args: argparse.Namespace) -> WindowInfo:
    windows = collect_windows()
    main = main_window(windows)

    if main is None and not args.no_launch:
        launch_xplatform(args)
        deadline = time.monotonic() + args.launch_timeout
        while time.monotonic() < deadline:
            windows = collect_windows()
            main = main_window(windows)
            if main is not None:
                break
            if login_window(windows) is not None:
                break
            time.sleep(1)

    windows = collect_windows()
    login = login_window(windows)
    if login is not None:
        if try_auto_login(args, login):
            windows = collect_windows()
        else:
            windows = collect_windows()

    login = login_window(windows)
    if login is not None:
        if args.login_timeout <= 0:
            raise RuntimeError("ICC login is required. Log in manually, then run the task again.")
        wait_args = argparse.Namespace(
            timeout=args.login_timeout,
            interval=10,
            screenshot=args.screenshot,
            output_dir=args.output_dir,
        )
        result = wait_login(wait_args)
        if result:
            raise RuntimeError("Timed out waiting for manual ICC login.")

    main = main_window()
    if main is None:
        raise RuntimeError("ICC main window was not found.")

    bring_to_front(main)
    return main


def open_on_demand_data(info: WindowInfo) -> None:
    from pywinauto import keyboard
    bring_to_front(info)
    time.sleep(2.0)
    
    # 1. 상단 Document 메뉴 클릭 (1280x728 기준 x=420, y=55)
    print("Clicking Document menu...")
    click_rel(info, 420, 55)
    time.sleep(2.5)
    
    # 2. On-Demand Data 메뉴 검색 (검색창 클릭 x=915, y=55)
    # 검색창을 확실히 활성화하기 위해 더블 클릭 시도
    print("Searching for On-Demand Data...")
    click_rel(info, 915, 55, double=True)
    time.sleep(1.0)
    keyboard.send_keys("^a{BACKSPACE}")
    time.sleep(0.5)
    keyboard.send_keys("On-Demand Data", with_spaces=True)
    time.sleep(1.5)
    keyboard.send_keys("{ENTER}")
    print("Waiting for On-Demand Data tab to open...")
    time.sleep(7.0)


def select_document(info: WindowInfo, document_name: str) -> None:
    from pywinauto import keyboard
    bring_to_front(info)
    time.sleep(1.5)

    # 텍스트 박스 먼저 더블 클릭하여 포커스 및 전체 선택
    click_rel(info, 280, 187, double=True)
    time.sleep(1.0)
    keyboard.send_keys("^a{BACKSPACE}")
    time.sleep(1.0)
    
    # 드롭다운 화살표 클릭하여 리스트 활성화 시도
    click_rel(info, 420, 187)
    time.sleep(1.0)
    
    # 키워드 입력
    search_keyword = "징수금액조회" if "징수금액" in document_name else document_name
    keyboard.send_keys(search_keyword, with_spaces=True)
    time.sleep(1.5)
    
    # 필터링 트리거 및 선택
    keyboard.send_keys("{SPACE}{BACKSPACE}")
    time.sleep(3.0)
    keyboard.send_keys("{DOWN}{ENTER}")
    time.sleep(2.0)
    keyboard.send_keys("{TAB}")
    time.sleep(3.0)


def set_condition_value(info: WindowInfo, rel_x: int, rel_y: int, value: str) -> None:
    from pywinauto import keyboard

    click_rel(info, rel_x, rel_y)
    click_rel(info, rel_x, rel_y, double=True)
    paste_text(value)
    keyboard.send_keys("{TAB}")
    time.sleep(0.2)


def set_conditions(info: WindowInfo, window: object, org: str, division: str) -> None:
    set_condition_value(info, 205, 235, f"{window.start_year}{window.start_week:02d}")
    set_condition_value(info, 205, 259, f"{window.end_year}{window.end_week:02d}")
    set_condition_value(info, 205, 283, org)
    set_condition_value(info, 205, 307, division)


def close_dynamiclist_workbooks() -> None:
    try:
        import win32com.client

        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return

    for index in range(excel.Workbooks.Count, 0, -1):
        workbook = excel.Workbooks.Item(index)
        name = str(workbook.Name or "").lower()
        full_name = str(getattr(workbook, "FullName", "") or "").lower()
        if name == "dynamiclist.csv" or full_name.endswith("\\dynamiclist.csv"):
            close_workbook_and_empty_excel(workbook)


def close_workbook_and_empty_excel(workbook: object) -> None:
    excel = None
    try:
        excel = workbook.Application
    except Exception:
        pass

    try:
        workbook.Close(SaveChanges=False)
    except Exception as exc:
        print(f"Could not close Excel export workbook: {exc}", file=sys.stderr)
        return

    if excel is None:
        return

    try:
        if int(excel.Workbooks.Count) == 0:
            excel.Quit()
            print("Closed empty Excel application.")
    except Exception as exc:
        print(f"Could not close empty Excel application: {exc}", file=sys.stderr)


def find_excel_export(timeout: int, min_rows: int) -> tuple[object, Path]:
    import win32com.client

    deadline = time.monotonic() + timeout
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            for index in range(1, excel.Workbooks.Count + 1):
                workbook = excel.Workbooks.Item(index)
                name = str(workbook.Name or "")
                if name.lower() != "dynamiclist.csv":
                    continue
                worksheet = workbook.Worksheets.Item(1)
                rows = int(worksheet.UsedRange.Rows.Count)
                full_name = str(workbook.FullName or "")
                if rows >= min_rows and full_name:
                    workbook.Save()
                    return workbook, Path(full_name)
        except Exception as exc:
            last_error = exc
        time.sleep(2)

    if last_error:
        raise RuntimeError(f"Timed out waiting for Excel export: {last_error}") from last_error
    raise RuntimeError("Timed out waiting for Excel export.")


def copy_excel_export(source: Path, output_file: Path) -> Path:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    if source.resolve() != output_file.resolve():
        shutil.copy2(source, output_file)
    return output_file


def run_xplatform_download(args: argparse.Namespace) -> Path:
    from icc_daily_update import parse_date, report_window

    today = parse_date(args.date)
    window = report_window(
        today=today,
        weeks=args.weeks,
        target_year=args.target_year,
        target_week=args.target_week,
        start_year=args.start_year,
        start_week=args.start_week,
    )

    print(
        "XPlatform ICC conditions: "
        f"{window.start_year}{window.start_week:02d} -> {window.end_year}{window.end_week:02d}, "
        f"org {args.org}, division {args.division}"
    )

    recover_from_fatal_xplatform_error("xplatform_before_start_error", Path(args.output_dir))
    recover_from_stale_loading_modal("xplatform_before_start_busy", Path(args.output_dir))

    info = ensure_main_window(args)
    open_on_demand_data(info)
    # Wait for the initial loading overlay to clear, forcefully close it if it persists
    wait_for_blank_modals_to_clear(
        30,  # 30초 대기 후 강제 종료 시도
        Path(args.output_dir),
        "xplatform_open_menu_wait",
        raise_on_timeout=False,
        dismiss_interval=10, # 10초마다 닫기 시도
    )
    info = ensure_main_window(args)
    select_document(info, args.document_name)
    wait_for_blank_modals_to_clear(30, Path(args.output_dir), "xplatform_document_wait", raise_on_timeout=False)
    set_conditions(info, window, args.org, args.division)
    wait_for_blank_modals_to_clear(30, Path(args.output_dir), "xplatform_conditions_wait", raise_on_timeout=False)

    close_dynamiclist_workbooks()
    click_rel(info, 930, 187)
    print(f"Search clicked. Waiting {args.search_wait} seconds for ICC results.")
    sleep_with_xplatform_checks(args.search_wait, Path(args.output_dir), "xplatform_search_wait")
    wait_for_blank_modals_to_clear(30, Path(args.output_dir), "xplatform_search_busy_wait", raise_on_timeout=False)

    workbook = None
    source = None
    last_export_error: Exception | None = None
    for attempt in range(1, args.export_attempts + 1):
        close_dynamiclist_workbooks()
        info = ensure_main_window(args)  # 핸들이 무효화되었을 수 있으므로 갱신
        click_rel(info, 865, 101)
        suffix = "" if args.export_attempts == 1 else f" (attempt {attempt}/{args.export_attempts})"
        print(f"Excel Down clicked. Waiting for DynamicList.CSV in Excel{suffix}.")
        try:
            workbook, source = find_excel_export(args.export_timeout, args.min_export_rows)
            break
        except Exception as exc:
            last_export_error = exc
            captures = capture_diagnostic_windows(Path(args.output_dir), "xplatform_after_export_fail")
            if captures:
                for capture in captures:
                    print(f"Diagnostic screenshot: {capture.resolve()}")
            if attempt >= args.export_attempts:
                break

            print(
                "Excel export was not detected. Retrying Search and Excel Down after "
                f"{args.export_retry_wait} seconds."
            )
            sleep_with_xplatform_checks(
                args.export_retry_wait,
                Path(args.output_dir),
                "xplatform_export_retry_wait",
            )
            info = ensure_main_window(args)
            click_rel(info, 930, 187)
            print(f"Search clicked. Waiting {args.search_wait} seconds for ICC results.")
            sleep_with_xplatform_checks(args.search_wait, Path(args.output_dir), "xplatform_search_retry_wait")
            wait_for_blank_modals_to_clear(30, Path(args.output_dir), "xplatform_search_retry_busy_wait")

    if workbook is None or source is None:
        close_dynamiclist_workbooks()
        raise RuntimeError(
            "Timed out waiting for DynamicList.CSV in Excel after Excel Down. "
            "For scheduled runs, keep the Windows session logged in and unlocked so "
            "XPlatform and Excel can open on the interactive desktop."
        ) from last_export_error

    output = Path(args.output_file)
    copied = copy_excel_export(source, output)
    print(f"Saved XPlatform export: {copied.resolve()}")

    if args.close_excel_export:
        close_workbook_and_empty_excel(workbook)

    if args.screenshot:
        fresh = main_window()
        if fresh:
            capture = capture_window(fresh, Path(args.output_dir))
            print(f"Screenshot: {capture.resolve()}")

    return copied


def launch_xplatform(args: argparse.Namespace) -> None:
    exe = Path(args.exe)
    if not exe.exists():
        raise FileNotFoundError(f"XPlatform executable not found: {exe}")

    existing = collect_windows()
    if existing and not args.force_new:
        print("XPlatform is already running.")
        print_windows(existing)
        return

    command = [str(exe), "-K", args.key, "-X", args.xadl]
    process = subprocess.Popen(command)
    print(f"Started XPlatform pid={process.pid}")
    time.sleep(args.startup_wait)
    print_windows(collect_windows())


def status(args: argparse.Namespace) -> None:
    windows = collect_windows()
    print_windows(windows)

    if args.screenshot:
        target = best_capture_window(windows)
        if not target:
            raise RuntimeError("No visible XPlatform window was available for screenshot.")
        output = capture_window(target, Path(args.output_dir))
        print(f"Screenshot: {output.resolve()}")


def wait_login(args: argparse.Namespace) -> int:
    deadline = time.monotonic() + args.timeout
    last_message = 0.0

    while time.monotonic() < deadline:
        windows = collect_windows()
        has_login = any(info.title == "Login Form" and visible_area(info) > 0 for info in windows)
        app_windows = [
            info
            for info in windows
            if info.title.startswith("KMTC :: ICC") and visible_area(info) > 0
        ]

        if app_windows and not has_login:
            print("Login Form is gone and the ICC main window is visible.")
            if args.screenshot:
                output = capture_window(best_capture_window(windows) or app_windows[0], Path(args.output_dir))
                print(f"Screenshot: {output.resolve()}")
            return 0

        now = time.monotonic()
        if now - last_message >= args.interval:
            if has_login:
                print("Waiting for manual ICC login to finish...")
            elif app_windows:
                print("ICC main window exists, but login completion is not confirmed yet...")
            else:
                print("Waiting for XPlatform windows...")
            last_message = now

        time.sleep(1)

    print("Timed out waiting for manual ICC login.", file=sys.stderr)
    return 2


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Launch and inspect the KMTC ICC XPlatform client.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    launch = subparsers.add_parser("launch", help="Launch the KMTC ICC XPlatform client.")
    launch.add_argument("--exe", default=str(DEFAULT_XPLATFORM_EXE))
    launch.add_argument("--key", default=DEFAULT_XPLATFORM_KEY)
    launch.add_argument("--xadl", default=DEFAULT_XPLATFORM_XADL)
    launch.add_argument("--force-new", action="store_true")
    launch.add_argument("--startup-wait", type=float, default=5.0)
    launch.set_defaults(func=launch_xplatform)

    check = subparsers.add_parser("status", help="List XPlatform windows and optionally capture a screenshot.")
    check.add_argument("--screenshot", action="store_true")
    check.add_argument("--output-dir", default=str(DEFAULT_LOG_DIR))
    check.set_defaults(func=status)

    wait = subparsers.add_parser("wait-login", help="Wait until the manual login screen is dismissed.")
    wait.add_argument("--timeout", type=int, default=300)
    wait.add_argument("--interval", type=int, default=10)
    wait.add_argument("--screenshot", action="store_true")
    wait.add_argument("--output-dir", default=str(DEFAULT_LOG_DIR))
    wait.set_defaults(func=wait_login)

    download = subparsers.add_parser("download", help="Download DynamicList.CSV through the XPlatform UI.")
    download.add_argument("--exe", default=str(DEFAULT_XPLATFORM_EXE))
    download.add_argument("--key", default=DEFAULT_XPLATFORM_KEY)
    download.add_argument("--xadl", default=DEFAULT_XPLATFORM_XADL)
    download.add_argument("--force-new", action="store_true")
    download.add_argument("--startup-wait", type=float, default=5.0)
    download.add_argument("--no-launch", action="store_true")
    download.add_argument("--launch-timeout", type=int, default=60)
    download.add_argument("--login-timeout", type=int, default=0)
    download.add_argument("--login-after-wait", type=int, default=20)
    download.add_argument("--login-auto-attempts", type=int, default=2)
    download.add_argument("--credential-target", default=CREDENTIAL_TARGET)
    download.add_argument("--document-name", default=DEFAULT_DOCUMENT_NAME)
    download.add_argument("--org", default="O")
    download.add_argument("--division", default="D")
    download.add_argument("--weeks", type=int, default=4)
    download.add_argument("--date", help="Override run date as YYYY-MM-DD.")
    download.add_argument("--target-year", type=int)
    download.add_argument("--target-week", type=int)
    download.add_argument("--start-year", type=int)
    download.add_argument("--start-week", type=int)
    download.add_argument(
        "--output-file",
        default=str(DEFAULT_DOWNLOAD_DIR / f"xplatform_DynamicList_{dt.datetime.now():%Y%m%d_%H%M%S}.csv"),
    )
    download.add_argument("--search-wait", type=int, default=45)
    download.add_argument("--export-timeout", type=int, default=120)
    download.add_argument("--export-attempts", type=int, default=2)
    download.add_argument("--export-retry-wait", type=int, default=10)
    download.add_argument("--min-export-rows", type=int, default=100)
    download.add_argument("--close-excel-export", action=argparse.BooleanOptionalAction, default=True)
    download.add_argument("--screenshot", action="store_true")
    download.add_argument("--output-dir", default=str(DEFAULT_LOG_DIR))
    download.set_defaults(func=run_xplatform_download)

    save = subparsers.add_parser("credential-save", help="Save the ICC password to Windows Credential Manager.")
    save.add_argument("--target", default=CREDENTIAL_TARGET)
    save.add_argument("--username", default=getpass.getuser())
    save.add_argument("--password-stdin", action="store_true")
    save.add_argument("--gui", action="store_true")
    save.set_defaults(func=credential_save)

    cred_status = subparsers.add_parser("credential-status", help="Check whether the ICC credential exists.")
    cred_status.add_argument("--target", default=CREDENTIAL_TARGET)
    cred_status.set_defaults(func=credential_status)

    delete = subparsers.add_parser("credential-delete", help="Delete the ICC credential.")
    delete.add_argument("--target", default=CREDENTIAL_TARGET)
    delete.set_defaults(func=credential_delete)

    return parser.parse_args()


def main() -> None:
    args = parse_args()
    result = args.func(args)
    if isinstance(result, int):
        raise SystemExit(result)


if __name__ == "__main__":
    main()
