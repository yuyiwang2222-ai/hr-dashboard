from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr.encoding != "utf-8":
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from dotenv import load_dotenv

PROJECT_DIR = Path(__file__).resolve().parents[1]
load_dotenv(PROJECT_DIR / ".env")

ATTENDANCE_SOURCE_PATH = Path(
    os.getenv(
        "ATTENDANCE_SOURCE_PATH",
        r"J:\HR-人資\4.薪酬\2.考勤保險\1.每日出勤\每日出勤總表.xlsx",
    )
)
ATTENDANCE_TARGET_PATH = PROJECT_DIR / "數據資料夾" / "每日出勤總表.xlsx"
CACHE_ROOT = PROJECT_DIR / ".automation-cache"
APPDATA_ROOT = CACHE_ROOT / "appdata"
LOCALAPPDATA_ROOT = CACHE_ROOT / "localappdata"
TEMP_ROOT = CACHE_ROOT / "temp"


class MondayRunError(RuntimeError):
    pass


def log(message: str) -> None:
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}", flush=True)


def build_env() -> dict[str, str]:
    for path in (CACHE_ROOT, APPDATA_ROOT, LOCALAPPDATA_ROOT, TEMP_ROOT):
        path.mkdir(parents=True, exist_ok=True)
    env = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
    env.setdefault("APPDATA", str(APPDATA_ROOT))
    env.setdefault("LOCALAPPDATA", str(LOCALAPPDATA_ROOT))
    env.setdefault("TEMP", str(TEMP_ROOT))
    env.setdefault("TMP", str(TEMP_ROOT))
    return env


def run_command(args: list[str]) -> None:
    command = [sys.executable, *args]
    log(f"Running: {' '.join(command)}")
    result = subprocess.run(command, cwd=PROJECT_DIR, env=build_env())
    if result.returncode != 0:
        raise MondayRunError(f"Command failed with exit code {result.returncode}: {args}")


def copy_attendance_file() -> None:
    if not ATTENDANCE_SOURCE_PATH.exists():
        raise MondayRunError(f"Attendance source file not found: {ATTENDANCE_SOURCE_PATH}")
    ATTENDANCE_TARGET_PATH.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(ATTENDANCE_SOURCE_PATH, ATTENDANCE_TARGET_PATH)
    log(f"Attendance file copied to {ATTENDANCE_TARGET_PATH}")


def run_prepare_phase() -> None:
    copy_attendance_file()
    run_command(["scripts/export_hrm_employees.py"])
    run_command(["sync_data.py"])
    run_command(["scripts/weekly_job.py"])
    run_command(["auto_report.py", "--no-email"])
    log("Prepare phase finished.")


def run_email_phase() -> None:
    run_command(["send_report_email.py"])
    log("Email phase finished.")


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the scheduled Monday HR automation workflow.")
    parser.add_argument("--phase", choices=["prepare", "email", "all"], default="all")
    args = parser.parse_args()

    if args.phase in ("prepare", "all"):
        run_prepare_phase()
    if args.phase in ("email", "all"):
        run_email_phase()
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except MondayRunError as exc:
        log(f"Monday workflow failed: {exc}")
        raise SystemExit(1)
