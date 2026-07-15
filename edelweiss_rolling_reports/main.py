from __future__ import annotations

import argparse
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.append(str(REPO_ROOT))

from shared.bookscan_calendar import bookscan_week

try:
    from .rolling_customer_sales import (
        REFRESH_LOOKBACK_WEEKS,
        build_customer_sales_report,
        check_source_weeks,
        get_delta_week_status,
        get_latest_cache_week,
        get_latest_sql_week,
        print_cache_refresh_summary,
        print_result_summary,
        print_week_check,
        refresh_caches_only,
        validate_source_workbook,
    )
except ImportError:
    from rolling_customer_sales import (
        REFRESH_LOOKBACK_WEEKS,
        build_customer_sales_report,
        check_source_weeks,
        get_delta_week_status,
        get_latest_cache_week,
        get_latest_sql_week,
        print_cache_refresh_summary,
        print_result_summary,
        print_week_check,
        refresh_caches_only,
        validate_source_workbook,
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Edelweiss rolling reports helper.")
    parser.add_argument("--run-menu", action="store_true", help="Launch the interactive Edelweiss menu.")
    parser.add_argument("--refresh-lookback-weeks", type=int, default=REFRESH_LOOKBACK_WEEKS, help="Optional overlap weeks to re-pull from SQL; default 0 means only weeks after the cache max.")
    return parser


def _format_week_status(week) -> str:
    if week is None:
        return "None found"
    saturday_note = "" if week.weekday() == 5 else " (not a Saturday)"
    bookscan = bookscan_week(week)
    return f"Week {bookscan.week}, {week.strftime('%A, %B')} {week.day}, {bookscan.year}{saturday_note}"


def _build_week_status_message() -> str:
    latest_sql_week = get_latest_sql_week()
    latest_cache_week = get_latest_cache_week()
    delta_status = get_delta_week_status(latest_cache_week=latest_cache_week, latest_sql_week=latest_sql_week)
    lines = [
        f"SQL is updated through: {_format_week_status(latest_sql_week)}",
        f"Current Edelweiss cache through: {_format_week_status(latest_cache_week)}",
        f"Expected next Edelweiss week: {_format_week_status(delta_status.expected_next_week)}",
    ]
    if delta_status.missing_weeks:
        recent_missing = ", ".join(week.strftime("%m/%d/%Y") for week in delta_status.missing_weeks)
        lines.append(f"Missing SQL weeks detected since cache max: {len(delta_status.missing_weeks)}")
        lines.append(f"Missing delta week(s): {recent_missing}")
    else:
        lines.append("Missing SQL weeks detected since cache max: 0")
    lines.append("")
    lines.append("Would you like to proceed with the Edelweiss rolling report for this SQL week?")
    return "\n".join(lines)


def confirm_build_from_week_status() -> bool:
    print()
    print(_build_week_status_message())
    answer = input("Proceed? [Y/n]: ").strip().lower()
    return answer not in {"n", "no"}


def validate_source_for_action() -> None:
    source = validate_source_workbook()
    print(
        f"Validated Edelweiss source: {source.workbook.name} "
        f"through {source.expected_week:%m/%d/%Y}"
    )


def run_menu(refresh_lookback_weeks: int = REFRESH_LOOKBACK_WEEKS) -> None:
    while True:
        print("\nEdelweiss Rolling Reports")
        print()
        print("    1. Check Edelweiss SQL weekly coverage")
        print("    2. Refresh Edelweiss caches")
        print("    3. Build Edelweiss rolling report")
        print("    4. Build local review Edelweiss rolling report")
        print("    5. Exit")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice == "1":
            print("\nChecking Edelweiss SQL weekly coverage...")
            print_week_check(check_source_weeks())
            continue

        if choice == "2":
            validate_source_for_action()
            print("\nRefreshing Edelweiss caches...")
            result = refresh_caches_only(refresh_lookback_weeks=refresh_lookback_weeks)
            print_cache_refresh_summary(result)
            continue

        if choice == "3":
            validate_source_for_action()
            if not confirm_build_from_week_status():
                print("Edelweiss rolling report cancelled.")
                continue
            print("\nBuilding Edelweiss rolling report...")
            result = build_customer_sales_report(
                refresh_sales=True,
                refresh_lookback_weeks=refresh_lookback_weeks,
            )
            print_result_summary(result)
            continue

        if choice == "4":
            validate_source_for_action()
            if not confirm_build_from_week_status():
                print("Edelweiss rolling report cancelled.")
                continue
            print("\nBuilding local review Edelweiss rolling report...")
            result = build_customer_sales_report(
                refresh_sales=True,
                refresh_lookback_weeks=refresh_lookback_weeks,
                local_only=True,
            )
            print_result_summary(result)
            continue

        if choice in {"5", "q", "quit", "exit"}:
            print("Exiting Edelweiss Rolling Reports.")
            return

        print("Invalid choice. Please select a valid option.")


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    run_menu(refresh_lookback_weeks=args.refresh_lookback_weeks)


if __name__ == "__main__":
    main()


