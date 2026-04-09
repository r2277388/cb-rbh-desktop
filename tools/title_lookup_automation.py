from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from paths import process_paths


def interpreter_has_module(python_executable: str, module_name: str) -> bool:
    try:
        result = subprocess.run(
            [
                python_executable,
                "-c",
                (
                    "import importlib.util, sys; "
                    f"sys.exit(0 if importlib.util.find_spec('{module_name}') else 1)"
                ),
            ],
            capture_output=True,
            check=False,
            text=True,
        )
    except OSError:
        return False

    return result.returncode == 0


def get_excel_automation_python() -> str:
    candidates: list[str] = []

    if sys.executable:
        candidates.append(sys.executable)

    path_python = shutil.which("python")
    if path_python:
        candidates.append(path_python)

    candidates.append(r"C:\Python310\python.exe")

    seen: set[str] = set()
    for candidate in candidates:
        normalized = os.path.normcase(os.path.abspath(candidate))
        if normalized in seen:
            continue
        seen.add(normalized)
        if interpreter_has_module(candidate, "pythoncom"):
            return candidate

    raise RuntimeError(
        "Could not find a Python interpreter with pywin32/pythoncom for Excel automation."
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Manage the scheduled weekly Title Lookup workbook refresh."
    )
    parser.add_argument(
        "command",
        choices=["status", "register", "disable"],
        help="Show task status, create/update the task, or disable the weekly scheduled task.",
    )
    return parser.parse_args()


def parse_schtasks_list_output(output: str) -> dict[str, str]:
    values: dict[str, str] = {}
    for raw_line in output.splitlines():
        line = raw_line.strip()
        if not line or ":" not in line:
            continue
        key, value = line.split(":", 1)
        values[key.strip()] = value.strip()
    return values


def query_task_status() -> tuple[bool, dict[str, str]]:
    result = subprocess.run(
        [
            "schtasks",
            "/Query",
            "/TN",
            process_paths.TITLE_LOOKUP_TASK_NAME,
            "/FO",
            "LIST",
            "/V",
        ],
        capture_output=True,
        check=False,
        text=True,
    )

    if result.returncode != 0:
        return False, {}

    return True, parse_schtasks_list_output(result.stdout)


def build_refresh_command() -> str:
    python_executable = Path(get_excel_automation_python()).resolve()
    script_path = process_paths.EXCEL_REFRESH_SCRIPT
    workbook_path = process_paths.TITLE_LOOKUP_WORKBOOK
    connection_name = process_paths.TITLE_LOOKUP_CONNECTION_NAME
    table_name = process_paths.TITLE_LOOKUP_TABLE_NAME
    return (
        f'"{python_executable}" "{script_path}" "{workbook_path}" '
        f'--connection "{connection_name}" --table "{table_name}"'
    )


def print_status() -> int:
    exists, details = query_task_status()

    print("Weekly Title Lookup Automation")
    print(f"  Task name:      {process_paths.TITLE_LOOKUP_TASK_NAME}")
    print(f"  Location:       {process_paths.TITLE_LOOKUP_TASK_LOCATION}")
    print(f"  Schedule:       {process_paths.TITLE_LOOKUP_SCHEDULE_DESCRIPTION}")
    print(f"  Workbook:       {process_paths.TITLE_LOOKUP_WORKBOOK}")
    print(f"  Connection:     {process_paths.TITLE_LOOKUP_CONNECTION_NAME}")
    print(f"  Output table:   {process_paths.TITLE_LOOKUP_TABLE_NAME}")
    print(f"  Refresh command:{build_refresh_command()}")

    if not exists:
        print("  Status:         Not currently scheduled")
        return 0

    print("  Status:         Scheduled")
    for key in ("Next Run Time", "Last Run Time", "Last Result", "Status", "Schedule"):
        if key in details:
            label = key if key != "Schedule" else "Task Scheduler"
            print(f"  {label:<15}{details[key]}")
    return 0


def register_task() -> int:
    command = build_refresh_command()
    result = subprocess.run(
        [
            "schtasks",
            "/Create",
            "/TN",
            process_paths.TITLE_LOOKUP_TASK_NAME,
            "/SC",
            "WEEKLY",
            "/D",
            process_paths.TITLE_LOOKUP_SCHEDULE_DAY,
            "/ST",
            process_paths.TITLE_LOOKUP_SCHEDULE_TIME,
            "/TR",
            command,
            "/RL",
            "LIMITED",
            "/IT",
            "/F",
        ],
        capture_output=True,
        check=False,
        text=True,
    )

    if result.returncode != 0:
        print(result.stdout.strip())
        print(result.stderr.strip(), file=sys.stderr)
        return result.returncode

    print(result.stdout.strip())
    print()
    return print_status()


def disable_task() -> int:
    result = subprocess.run(
        [
            "schtasks",
            "/Change",
            "/TN",
            process_paths.TITLE_LOOKUP_TASK_NAME,
            "/Disable",
        ],
        capture_output=True,
        check=False,
        text=True,
    )

    if result.returncode != 0:
        print(result.stdout.strip())
        print(result.stderr.strip(), file=sys.stderr)
        return result.returncode

    print(result.stdout.strip())
    print()
    return print_status()


def main() -> int:
    args = parse_args()
    if args.command == "status":
        return print_status()
    if args.command == "register":
        return register_task()
    return disable_task()


if __name__ == "__main__":
    raise SystemExit(main())
