from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from paths import process_paths


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Manage the scheduled General Editorial data variation archive."
    )
    parser.add_argument(
        "command",
        choices=["status", "register", "disable"],
        help="Show task status, create/update the task, or disable the scheduled task.",
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
            process_paths.GEN_EDITORIAL_TASK_NAME,
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


def build_run_command() -> str:
    python_executable = ROOT_DIR / "venv" / "Scripts" / "python.exe"
    return f'"{python_executable}" "{process_paths.GEN_EDITORIAL_VARIATIONS_SCRIPT}" run'


def print_status() -> int:
    exists, details = query_task_status()

    print("General Editorial Data Variations Automation")
    print(f"  Task name:      {process_paths.GEN_EDITORIAL_TASK_NAME}")
    print(f"  Location:       {process_paths.GEN_EDITORIAL_TASK_LOCATION}")
    print(f"  Schedule:       {process_paths.GEN_EDITORIAL_SCHEDULE_DESCRIPTION}")
    print(f"  Source:         {process_paths.GEN_EDITORIAL_SOURCE_WORKBOOK}")
    print(f"  Cache:          {process_paths.GEN_EDITORIAL_CACHE_FILE}")
    print(f"  Report:         {process_paths.GEN_EDITORIAL_REPORT_FILE}")
    print(f"  Run command:    {build_run_command()}")

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
    result = subprocess.run(
        [
            "schtasks",
            "/Create",
            "/TN",
            process_paths.GEN_EDITORIAL_TASK_NAME,
            "/SC",
            "WEEKLY",
            "/D",
            process_paths.GEN_EDITORIAL_SCHEDULE_DAYS,
            "/ST",
            process_paths.GEN_EDITORIAL_SCHEDULE_TIME,
            "/TR",
            build_run_command(),
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
            process_paths.GEN_EDITORIAL_TASK_NAME,
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
