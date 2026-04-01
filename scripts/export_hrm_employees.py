from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
CACHE_ROOT = PROJECT_DIR / ".automation-cache"
APPDATA_ROOT = CACHE_ROOT / "appdata"
LOCALAPPDATA_ROOT = CACHE_ROOT / "localappdata"
TEMP_ROOT = CACHE_ROOT / "temp"

for path in (CACHE_ROOT, APPDATA_ROOT, LOCALAPPDATA_ROOT, TEMP_ROOT):
    path.mkdir(parents=True, exist_ok=True)

os.environ.setdefault("APPDATA", str(APPDATA_ROOT))
os.environ.setdefault("LOCALAPPDATA", str(LOCALAPPDATA_ROOT))
os.environ.setdefault("TEMP", str(TEMP_ROOT))
os.environ.setdefault("TMP", str(TEMP_ROOT))
os.environ.setdefault("PYTHONUTF8", "1")
os.environ.setdefault("PYTHONIOENCODING", "utf-8")

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr.encoding != "utf-8":
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from dotenv import load_dotenv
from pywinauto import Desktop, keyboard, mouse
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.timings import TimeoutError as PywinautoTimeoutError

load_dotenv(PROJECT_DIR / ".env")

HRM_LAUNCHER_PATH = Path(os.getenv("HRM_LAUNCHER_PATH", r"C:\Program Files (x86)\Digiwin\DigiwinHR\Launcher.exe"))
HRM_USERNAME = os.getenv("HRM_USERNAME", "").strip()
HRM_PASSWORD = os.getenv("HRM_PASSWORD", "").strip()
HRM_COMPANY = os.getenv("HRM_COMPANY", "").strip()
HRM_LANGUAGE = os.getenv("HRM_LANGUAGE", "繁體中文").strip()
HRM_EXPORT_DIR = Path(os.getenv("HRM_EXPORT_DIR", str(PROJECT_DIR / "數據資料夾")))
HRM_EXPORT_FILE = HRM_EXPORT_DIR / "員工人數.xlsx"

NO_PAGING_BUTTON = (378, 777)
REFRESH_TOTAL_LINK = (676, 778)
PAGE_SIZE_BOX = (489, 777)
GRID_RIGHT_CLICK_POINT = (594, 592)

TREE_PATH = r"\人事跟蹤管理\員工資料管理\員工資料"
TAB_ALL_EMPLOYEES = "所有員工"
SEARCH_BUTTON_PATTERN = r"查找.*"
CONFIRM_BUTTON_PATTERN = r"確定.*"
SAVE_BUTTON_PATTERN = r"存檔.*|儲存.*|Save.*"
EXPORT_MENU_TITLE = "匯出到Excel"
SAVE_DIALOG_PATTERN = r"另存新檔|Save As"
EXPORT_WINDOW_PATTERNS = ("Exporting", "匯出", "匯出到Excel")`r`nLOGIN_WINDOW_PATTERNS = (r".*HR Solution AiGP.*", r"^Login$")`r`nMAIN_WINDOW_PATTERNS = (r".*HR Solution\[Standard\].*", r".*我的工作桌面.*")


class HRMExportError(RuntimeError):
    pass


def log(message: str) -> None:
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}", flush=True)


def wait_window(title_re: str, timeout: int = 60, backend: str = "win32"):
    desktop = Desktop(backend=backend)
    deadline = time.time() + timeout
    last_error = None
    while time.time() < deadline:
        try:
            window = desktop.window(title_re=title_re)
            if window.exists(timeout=1):
                return window
        except (ElementNotFoundError, PywinautoTimeoutError) as exc:
            last_error = exc
        time.sleep(1)
    raise HRMExportError(f"Could not find window matching {title_re!r}: {last_error}")


def find_first(window, **criteria):
    controls = window.descendants(**criteria)
    if not controls:
        raise HRMExportError(f"Control not found: {criteria}")
    return controls[0]


def click_relative(window, point: tuple[int, int], button: str = "left", double: bool = False) -> None:
    rect = window.rectangle()
    coords = (rect.left + point[0], rect.top + point[1])
    if button == "right":
        mouse.click(button="right", coords=coords)
    elif double:
        mouse.double_click(coords=coords)
    else:
        mouse.click(coords=coords)
    time.sleep(1)


def launch_hrm() -> None:
    if not HRM_LAUNCHER_PATH.exists():
        raise HRMExportError(f"Launcher not found: {HRM_LAUNCHER_PATH}")
    log(f"Launching HRM from {HRM_LAUNCHER_PATH}")
    subprocess.Popen([str(HRM_LAUNCHER_PATH)], cwd=str(HRM_LAUNCHER_PATH.parent))


def fill_login(dialog) -> None:
    edits = dialog.descendants(class_name="Edit")
    if len(edits) < 2:
        raise HRMExportError("Could not locate HRM login edit boxes.")

    edits[0].set_edit_text(HRM_USERNAME)
    edits[1].set_edit_text(HRM_PASSWORD)
    if HRM_COMPANY and len(edits) >= 3:
        edits[2].set_edit_text(HRM_COMPANY)

    try:
        combo = dialog.child_window(class_name="ComboBox")
        if combo.exists(timeout=1):
            combo.select(HRM_LANGUAGE)
    except Exception:
        pass

    confirm = dialog.child_window(title_re=CONFIRM_BUTTON_PATTERN, class_name="Button")
    confirm.wait("enabled", timeout=10)
    confirm.click_input()
    time.sleep(3)


def open_employee_browser(main_window) -> None:
    tree = main_window.child_window(class_name="SysTreeView32")
    try:
        item = tree.get_item(TREE_PATH)
        item.ensure_visible()
        item.click_input(double=True)
    except Exception as exc:
        raise HRMExportError(f"Failed to open 員工資料 browser: {exc}") from exc
    time.sleep(3)


def ensure_all_employees() -> None:
    for backend in ("uia", "win32"):
        try:
            browse = Desktop(backend=backend).window(title_re=r".*員工資料.*HR Solution\[Standard\].*")
            tab = browse.child_window(title=TAB_ALL_EMPLOYEES)
            if tab.exists(timeout=2):
                tab.click_input()
                time.sleep(1)
                return
        except Exception:
            continue
    raise HRMExportError("Could not activate the 所有員工 tab.")


def run_search(main_window) -> None:
    search_button = main_window.child_window(title_re=SEARCH_BUTTON_PATTERN, class_name="Button")
    search_button.wait("enabled", timeout=10)
    search_button.click_input()
    time.sleep(4)


def configure_no_paging(main_window) -> None:
    click_relative(main_window, NO_PAGING_BUTTON)
    click_relative(main_window, REFRESH_TOTAL_LINK)
    click_relative(main_window, PAGE_SIZE_BOX)
    keyboard.send_keys("^a{BACKSPACE}100000")
    time.sleep(1)


def export_grid_to_excel(main_window) -> None:
    click_relative(main_window, GRID_RIGHT_CLICK_POINT, button="right")
    time.sleep(1)
    for menu in Desktop(backend="win32").windows(class_name="#32768"):
        try:
            item = menu.child_window(title=EXPORT_MENU_TITLE)
            if item.exists(timeout=1):
                item.click_input()
                time.sleep(2)
                return
        except Exception:
            continue
    keyboard.send_keys("{DOWN 8}{ENTER}")
    time.sleep(2)


def save_export() -> None:
    HRM_EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    dialog = wait_window(SAVE_DIALOG_PATTERN, timeout=30, backend="uia")
    try:
        filename_box = dialog.child_window(auto_id="1001", control_type="Edit")
        if not filename_box.exists(timeout=1):
            raise RuntimeError
    except Exception:
        filename_box = find_first(dialog, control_type="Edit")
    filename_box.set_edit_text(str(HRM_EXPORT_FILE))
    save_button = dialog.child_window(title_re=SAVE_BUTTON_PATTERN, control_type="Button")
    save_button.click_input()
    time.sleep(2)

    for title_re in (r"確認另存新檔", r"Confirm Save As", r".*確認.*", r".*覆蓋.*"):
        try:
            confirm = Desktop(backend="uia").window(title_re=title_re)
            if confirm.exists(timeout=2):
                try:
                    yes_button = confirm.child_window(title_re=r"是.*|Yes.*|確定.*", control_type="Button")
                    yes_button.click_input()
                    time.sleep(1)
                except Exception:
                    keyboard.send_keys("{ENTER}")
                break
        except Exception:
            continue


def wait_for_export(timeout: int = 180) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        exporting_visible = False
        for title_re in EXPORT_WINDOW_PATTERNS:
            try:
                exporting = Desktop(backend="win32").window(title_re=f".*{title_re}.*")
                if exporting.exists(timeout=1):
                    exporting_visible = True
                    break
            except Exception:
                continue
        if HRM_EXPORT_FILE.exists() and not exporting_visible:
            return
        time.sleep(2)
    raise HRMExportError("Timed out waiting for HRM export to finish.")


def main() -> int:
    if not HRM_USERNAME or not HRM_PASSWORD:
        raise HRMExportError("HRM_USERNAME and HRM_PASSWORD must be set in .env or environment variables.")

    launch_hrm()
    login_dialog = wait_window_any(LOGIN_WINDOW_PATTERNS, timeout=60, backend="win32")
    fill_login(login_dialog)
    main_window = wait_window_any(MAIN_WINDOW_PATTERNS, timeout=90, backend="win32")
    open_employee_browser(main_window)
    ensure_all_employees()
    run_search(main_window)
    configure_no_paging(main_window)
    run_search(main_window)
    export_grid_to_excel(main_window)
    save_export()
    wait_for_export()
    log(f"HRM employee export finished: {HRM_EXPORT_FILE}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except HRMExportError as exc:
        log(f"HRM export failed: {exc}")
        raise SystemExit(1)

