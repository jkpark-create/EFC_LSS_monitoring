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
    return title in {"Login Form"} or title.startswith("KMTC :: ICC")


def collect_windows(backends: Iterable[str] = ("uia", "win32")) -> list[WindowInfo]:
    windows: list[WindowInfo] = []
    seen: set[tuple[str, int]] = set()

    for backend in backends:
        try:
            desktop = load_desktop(backend)
            wrappers = desktop.windows()
        except Exception as exc:
            print(f"{backend}: could not enumerate windows: {exc}", file=sys.stderr)
            continue

        for wrapper in wrappers:
            title = window_title(wrapper)
            class_name = window_class(wrapper)
            if not looks_like_xplatform(title, class_name):
                continue

            handle = window_handle(wrapper)
            key = (backend, handle)
            if key in seen:
                continue
            seen.add(key)

            windows.append(
                WindowInfo(
                    backend=backend,
                    handle=handle,
                    process_id=process_id(wrapper),
                    title=title,
                    class_name=class_name,
                    rectangle=window_rectangle(wrapper),
                    wrapper=wrapper,
                )
            )

    return windows


def visible_area(info: WindowInfo) -> int:
    rect = info.rectangle
    numbers: list[int] = []
    for token in rect.replace("(", " ").replace(")", " ").replace(",", " ").split():
        if len(token) > 1 and token[0] in {"L", "T", "R", "B"}:
            token = token[1:]
        try:
            numbers.append(int(token))
        except ValueError:
            pass

    if len(numbers) < 4:
        return 0
    left, top, right, bottom = numbers[:4]
    return max(0, right - left) * max(0, bottom - top)


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


def capture_window(info: WindowInfo, output_dir: Path, prefix: str = "xplatform") -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    safe_title = "".join(char if char.isalnum() else "_" for char in (info.title or "window")).strip("_")
    output = output_dir / f"{prefix}_{safe_title}_{stamp}.png"

    wrapper = info.wrapper
    try:
        set_focus = getattr(wrapper, "set_focus", None)
        if callable(set_focus):
            set_focus()
            time.sleep(0.3)
    except Exception:
        pass

    image = wrapper.capture_as_image()
    image.save(output)
    return output


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


def paste_text(value: str) -> None:
    import pyperclip
    from pywinauto import keyboard

    keyboard.send_keys("^a")
    time.sleep(0.1)
    pyperclip.copy(value)
    keyboard.send_keys("^v")
    time.sleep(0.2)


def try_auto_login(args: argparse.Namespace, login: WindowInfo) -> bool:
    credential = read_stored_credential(args.credential_target)
    if credential is None:
        return False

    username, password = credential
    if not password:
        return False

    bring_to_front(login)
    if username:
        click_scaled(login, 510, 112, LOGIN_WINDOW_SIZE)
        paste_text(username)
    click_scaled(login, 510, 160, LOGIN_WINDOW_SIZE)
    paste_text(password)
    click_scaled(login, 570, 205, LOGIN_WINDOW_SIZE)

    deadline = time.monotonic() + args.login_after_wait
    while time.monotonic() < deadline:
        windows = collect_windows()
        if login_window(windows) is None and main_window(windows) is not None:
            print("ICC auto-login completed with Windows Credential Manager.")
            return True
        time.sleep(1)

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
    click_rel(info, 1160, 41)
    paste_text("On-Demand Data")
    click_rel(info, 1240, 41)
    time.sleep(1.0)
    click_rel(info, 955, 96, double=True)
    time.sleep(5.0)


def select_document(info: WindowInfo, document_name: str) -> None:
    from pywinauto import keyboard

    click_rel(info, 260, 138)
    paste_text(document_name)
    keyboard.send_keys("{ENTER}")
    time.sleep(3.0)


def set_condition_value(info: WindowInfo, rel_x: int, rel_y: int, value: str) -> None:
    from pywinauto import keyboard

    click_rel(info, rel_x, rel_y)
    click_rel(info, rel_x, rel_y, double=True)
    paste_text(value)
    keyboard.send_keys("{TAB}")
    time.sleep(0.2)


def set_conditions(info: WindowInfo, window: object, org: str, division: str) -> None:
    set_condition_value(info, 205, 185, f"{window.start_year}{window.start_week:02d}")
    set_condition_value(info, 205, 209, f"{window.end_year}{window.end_week:02d}")
    set_condition_value(info, 205, 233, org)
    set_condition_value(info, 205, 257, division)


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
            workbook.Close(SaveChanges=False)


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

    info = ensure_main_window(args)
    open_on_demand_data(info)
    info = ensure_main_window(args)
    select_document(info, args.document_name)
    set_conditions(info, window, args.org, args.division)

    close_dynamiclist_workbooks()
    click_rel(info, 1190, 138)
    print(f"Search clicked. Waiting {args.search_wait} seconds for ICC results.")
    time.sleep(args.search_wait)

    close_dynamiclist_workbooks()
    click_rel(info, 1110, 101)
    print("Excel Down clicked. Waiting for DynamicList.CSV in Excel.")
    workbook, source = find_excel_export(args.export_timeout, args.min_export_rows)

    output = Path(args.output_file)
    copied = copy_excel_export(source, output)
    print(f"Saved XPlatform export: {copied.resolve()}")

    if args.close_excel_export:
        workbook.Close(SaveChanges=False)

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
