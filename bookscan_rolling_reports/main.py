from __future__ import annotations

import argparse
from tkinter import Tk, messagebox

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
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Bookscan rolling reports helper.")
    parser.add_argument("--run-menu", action="store_true", help="Launch the interactive Bookscan menu.")
    parser.add_argument(
        "--refresh-lookback-weeks",
        type=int,
        default=REFRESH_LOOKBACK_WEEKS,
        help="How many recent weekly Bookscan weeks to re-pull from SQL when refreshing the sales cache.",
    )
    return parser


def _format_week_status(week) -> str:
    if week is None:
        return "None found"
    saturday_note = "" if week.weekday() == 5 else " (not a Saturday)"
    return (
        f"Week {week.isocalendar().week}, "
        f"{week.strftime('%A, %B')} {week.day}, {week.year}"
        f"{saturday_note}"
    )


def _build_week_status_message() -> str:
    latest_sql_week = get_latest_sql_week()
    latest_cache_week = get_latest_cache_week()
    delta_status = get_delta_week_status(
        latest_cache_week=latest_cache_week,
        latest_sql_week=latest_sql_week,
    )
    lines = [
        f"SQL is updated through: {_format_week_status(latest_sql_week)}",
        f"Current Bookscan cache through: {_format_week_status(latest_cache_week)}",
        f"Expected next Bookscan week: {_format_week_status(delta_status.expected_next_week)}",
    ]
    if delta_status.missing_weeks:
        recent_missing = ", ".join(
            week.strftime("%m/%d/%Y") for week in delta_status.missing_weeks
        )
        lines.append(f"Missing SQL weeks detected since cache max: {len(delta_status.missing_weeks)}")
        lines.append(f"Missing delta week(s): {recent_missing}")
    else:
        lines.append("Missing SQL weeks detected since cache max: 0")
    lines.append("")
    lines.append("Would you like to proceed with the Bookscan rolling report for this SQL week?")
    return "\n".join(lines)


def confirm_build_from_week_status() -> bool:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        return messagebox.askyesno(
            "Run Bookscan Rolling Report",
            _build_week_status_message(),
            parent=root,
        )
    finally:
        root.destroy()


def run_menu(refresh_lookback_weeks: int = REFRESH_LOOKBACK_WEEKS) -> None:
    while True:
        print("\nBookscan Rolling Reports")
        print()
        print("    1. Check Bookscan SQL weekly coverage")
        print("    2. Refresh Bookscan caches")
        print("    3. Build Bookscan rolling report (main only)")
        print("    4. Build Bookscan rolling report (main + DP versions)")
        print("    5. Build local review Bookscan rolling report")
        print("    6. Exit")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice == "1":
            print("\nChecking Bookscan SQL weekly coverage...")
            print_week_check(check_source_weeks())
            continue

        if choice == "2":
            print("\nRefreshing Bookscan caches...")
            result = refresh_caches_only(refresh_lookback_weeks=refresh_lookback_weeks)
            print_cache_refresh_summary(result)
            continue

        if choice == "3":
            if not confirm_build_from_week_status():
                print("Bookscan rolling report cancelled.")
                continue
            print("\nBuilding Bookscan rolling report...")
            result = build_customer_sales_report(
                refresh_lookback_weeks=refresh_lookback_weeks
            )
            print_result_summary(result)
            continue

        if choice == "4":
            if not confirm_build_from_week_status():
                print("Bookscan rolling report cancelled.")
                continue
            print("\nBuilding Bookscan rolling report + DP versions...")
            result = build_customer_sales_report(
                refresh_lookback_weeks=refresh_lookback_weeks,
                save_dp=True,
            )
            print_result_summary(result)
            continue

        if choice == "5":
            if not confirm_build_from_week_status():
                print("Bookscan rolling report cancelled.")
                continue
            print("\nBuilding local review Bookscan rolling report...")
            result = build_customer_sales_report(
                refresh_lookback_weeks=refresh_lookback_weeks,
                local_only=True,
            )
            print_result_summary(result)
            continue

        if choice in {"6", "q", "quit", "exit"}:
            print("Exiting Bookscan Rolling Reports.")
            return

        print("Invalid choice. Please select a valid option.")


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    run_menu(refresh_lookback_weeks=args.refresh_lookback_weeks)


if __name__ == "__main__":
    main()
