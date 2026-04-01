from __future__ import annotations

import argparse
import json
import os
import socket
import subprocess
import sys
import time
import webbrowser
from datetime import datetime
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
SOURCE_DIR = PROJECT_DIR / "數據資料夾"
STATE_FILE = PROJECT_DIR / ".monday-workflow-state.json"
STREAMLIT_PORT = 8501
STREAMLIT_URL = f"http://127.0.0.1:{STREAMLIT_PORT}"
WATCHED_FILES = {
    "員工人數.xlsx": PROJECT_DIR / "data" / "employees.xlsx",
    "每日出勤總表.xlsx": PROJECT_DIR / "data" / "attendance.xlsx",
}
DEBOUNCE_SECONDS = 8
POLL_SECONDS = 3


class WorkflowError(RuntimeError):
    pass


def log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}", flush=True)


def file_stamp(path: Path) -> dict[str, int | bool | None]:
    if not path.exists():
        return {"exists": False, "mtime_ns": None, "size": None}

    stat = path.stat()
    return {"exists": True, "mtime_ns": stat.st_mtime_ns, "size": stat.st_size}


def source_path(filename: str) -> Path:
    return SOURCE_DIR / filename


def build_source_snapshot() -> dict[str, dict[str, int | bool | None]]:
    return {name: file_stamp(source_path(name)) for name in WATCHED_FILES}


def files_pending_sync() -> tuple[list[str], list[str]]:
    pending: list[str] = []
    missing: list[str] = []

    for source_name, target_path in WATCHED_FILES.items():
        source_file = source_path(source_name)
        source_stamp = file_stamp(source_file)
        if not source_stamp["exists"]:
            missing.append(source_name)
            continue

        target_stamp = file_stamp(target_path)
        if (
            not target_stamp["exists"]
            or source_stamp["mtime_ns"] != target_stamp["mtime_ns"]
            or source_stamp["size"] != target_stamp["size"]
        ):
            pending.append(source_name)

    return pending, missing


SUBPROCESS_ENV = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}


def run_python_script(relative_path: str) -> None:
    script_path = PROJECT_DIR / relative_path
    command = [sys.executable, str(script_path)]
    log(f"Running: {' '.join(command)}")
    result = subprocess.run(command, cwd=PROJECT_DIR, env=SUBPROCESS_ENV)
    if result.returncode != 0:
        raise WorkflowError(f"{relative_path} exited with code {result.returncode}.")


def is_port_open(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as connection:
        connection.settimeout(1)
        return connection.connect_ex((host, port)) == 0


def wait_for_port(host: str, port: int, timeout_seconds: int = 45) -> bool:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if is_port_open(host, port):
            return True
        time.sleep(1)
    return False


class StreamlitManager:
    def __init__(self) -> None:
        self.process: subprocess.Popen[str] | None = None

    def stop(self) -> None:
        if self.process is None or self.process.poll() is not None:
            self.process = None
            return

        log("Stopping previous local Streamlit process.")
        self.process.terminate()
        try:
            self.process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            self.process.kill()
            self.process.wait(timeout=5)
        self.process = None

    def start(self) -> None:
        if self.process is not None and self.process.poll() is None:
            log(f"Local dashboard already running at {STREAMLIT_URL}.")
            return

        if is_port_open("127.0.0.1", STREAMLIT_PORT):
            log(f"Port {STREAMLIT_PORT} is already in use. Reusing the existing local dashboard.")
            return

        command = [
            sys.executable,
            "-m",
            "streamlit",
            "run",
            "app.py",
            "--server.address",
            "127.0.0.1",
            "--server.port",
            str(STREAMLIT_PORT),
            "--server.headless",
            "true",
        ]
        log("Starting local Streamlit dashboard.")
        self.process = subprocess.Popen(command, cwd=PROJECT_DIR, env=SUBPROCESS_ENV)

        if not wait_for_port("127.0.0.1", STREAMLIT_PORT):
            raise WorkflowError("Streamlit did not start within 45 seconds.")

        log(f"Dashboard is ready at {STREAMLIT_URL}.")


class MondayWorkflow:
    def __init__(self) -> None:
        self.streamlit = StreamlitManager()

    def run(self, reason: str, changed_files: list[str]) -> None:
        details = ", ".join(changed_files) if changed_files else "manual run"
        log(f"Starting workflow because {reason}: {details}")

        run_python_script("sync_data.py")
        run_python_script("scripts/weekly_job.py")
        self.streamlit.stop()
        self.streamlit.start()
        self.write_state(reason=reason, changed_files=changed_files)
        self.open_browser()
        log("Workflow finished.")

    def write_state(self, reason: str, changed_files: list[str]) -> None:
        payload = {
            "last_run_at": datetime.now().isoformat(timespec="seconds"),
            "reason": reason,
            "changed_files": changed_files,
        }
        STATE_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    @staticmethod
    def open_browser() -> None:
        try:
            webbrowser.open(STREAMLIT_URL, new=2)
            log(f"Opened browser: {STREAMLIT_URL}")
        except Exception as exc:
            log(f"Could not open the browser automatically: {exc}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Watch Monday HR source files and run the dashboard workflow.")
    parser.add_argument("--once", action="store_true", help="Run at most once, then exit.")
    parser.add_argument("--force", action="store_true", help="Run immediately even when source and data files already match.")
    args = parser.parse_args()

    workflow = MondayWorkflow()
    log("Starting Monday workflow watcher.")

    pending_files, missing_files = files_pending_sync()
    if args.force:
        workflow.run("manual force", list(WATCHED_FILES))
    elif pending_files and not missing_files:
        workflow.run("source files are newer than synced data", pending_files)
    elif missing_files:
        log(f"Waiting for source files: {', '.join(missing_files)}")
    else:
        log("Source files already match data/. No startup sync was needed.")

    if args.once:
        log("One-shot mode completed.")
        return 0

    log("Workflow ready. Waiting for source file updates.")
    snapshot = build_source_snapshot()
    pending_changes: set[str] = set()
    last_change_at: float | None = None

    while True:
        time.sleep(POLL_SECONDS)
        current_snapshot = build_source_snapshot()

        for name, current_stamp in current_snapshot.items():
            if current_stamp != snapshot[name]:
                pending_changes.add(name)
                last_change_at = time.monotonic()
                log(f"Detected source file change: {name}")

        snapshot = current_snapshot

        if pending_changes and last_change_at is not None:
            quiet_for = time.monotonic() - last_change_at
            if quiet_for >= DEBOUNCE_SECONDS:
                try:
                    workflow.run("source files changed", sorted(pending_changes))
                    pending_changes.clear()
                    snapshot = build_source_snapshot()
                except WorkflowError as exc:
                    log(f"Workflow failed: {exc}")
                    log("Waiting for another quiet window before retrying.")
                    last_change_at = time.monotonic()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        log("Stopped by user.")
        raise SystemExit(130)
