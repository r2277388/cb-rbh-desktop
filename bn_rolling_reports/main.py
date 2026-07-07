from __future__ import annotations

import argparse
import sys
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.append(str(REPO_ROOT))

from config import BASE_FOLDER
from shared.bookscan_calendar import bookscan_week
from pos_combiner import (
    build_combined_pos,
    collect_pos_source_files,
    get_candidate_raw_folders,
    resolve_raw_folder,
)
from inventory_working import build_inventory_working_file, find_inventory_source_file
from sales_working import build_sales_working_file, find_sales_source_file
from rolling_customer_sales import (
    build_customer_sales_report,
    get_latest_cache_week,
    get_latest_sql_week,
    print_cache_refresh_summary,
    print_result_summary as print_rolling_result_summary,
    refresh_caches_only,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Barnes & Noble rolling reports helper.")
    parser.add_argument(
        "--raw-folder",
        help="Full path to a yyyy_mm_dd_raw_files folder. Defaults to the latest matching folder.",
    )
    parser.add_argument(
        "--output-file",
        help="Optional full output path for the combined POS Excel file.",
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="Print the first 25 rows after building the combined file.",
    )
    parser.add_argument(
        "--allow-missing-pos-files",
        action="store_true",
        help="Build the combined POS file even when one or more expected POS category files are missing.",
    )
    parser.add_argument(
        "--default-raw-folder",
        help="Preferred raw folder to use as the default in the interactive menu.",
    )
    return parser


def prompt_for_raw_folder(default_raw_folder: str | Path | None = None) -> Path:
    def choose_raw_folder_in_window() -> Path | None:
        initial_dir = str(BASE_FOLDER) if BASE_FOLDER.exists() else str(Path.home())
        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            selected = filedialog.askdirectory(
                initialdir=initial_dir,
                title="Choose the Barnes & Noble raw files folder",
            )
        finally:
            root.destroy()

        if not selected:
            return None
        return resolve_raw_folder(selected)

    candidates = get_candidate_raw_folders()
    preferred = None
    if default_raw_folder:
        preferred = resolve_raw_folder(default_raw_folder)

    if candidates:
        latest = preferred or candidates[-1]
        while True:
            print()
            print("Choose the Barnes & Noble raw folder:")
            print(f"    1. Use latest folder: {latest}")
            print("    2. Choose a folder in a window")
            print("    3. Paste a full folder path")
            print()
            choice = input("Choose an option: ").strip().lower()

            if choice in {"", "1", "latest", "default"}:
                return resolve_raw_folder(latest)

            if choice in {"2", "window", "browse", "folder"}:
                selected = choose_raw_folder_in_window()
                if selected is not None:
                    return selected
                print("No folder was selected.")
                continue

            if choice in {"3", "paste", "path"}:
                user_value = input("Paste the full raw folder path: ").strip()
                if user_value:
                    return resolve_raw_folder(user_value)
                print("No folder path was entered.")
                continue

            print("Invalid choice. Please select a valid option.")

    while True:
        print()
        print("The default Barnes & Noble base folder was not available.")
        print("    1. Choose a folder in a window")
        print("    2. Paste a full raw folder path")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "window", "browse", "folder"}:
            selected = choose_raw_folder_in_window()
            if selected is not None:
                return selected
            print("No folder was selected.")
            continue

        if choice in {"2", "paste", "path"}:
            user_value = input("Paste the full raw folder path: ").strip()
            if user_value:
                return resolve_raw_folder(user_value)
            print("No folder path was entered.")
            continue

        print("Invalid choice. Please select a valid option.")


def print_result_summary(result) -> None:
    print()
    print(f"Raw folder: {result.raw_folder}")
    print("Source files:")
    for file_path in result.source_files:
        print(f"  - {file_path.name}")
    if result.missing_keywords:
        print("Missing expected POS categories:")
        for keyword in result.missing_keywords:
            print(f"  - {keyword}")
    print(f"Rows before de-duplication: {result.rows_before_dedup:,}")
    print(f"Duplicate EAN rows removed: {result.duplicate_rows_removed:,}")
    print(f"Rows dropped for invalid/missing EAN: {result.rows_with_missing_ean:,}")
    print(f"Rows after de-duplication: {result.rows_after_dedup:,}")
    print(f"Saved file: {result.output_file}")


def print_sales_result_summary(result) -> None:
    print()
    print(f"Sales source file: {result.source_file.name}")
    print(f"Matched ISBN overrides: {result.matched_updates:,}")
    print(f"Appended new ISBN rows: {result.appended_rows:,}")
    print(f"Rows removed by ISBN whitelist: {result.excluded_rows:,}")
    print(f"Final data shape: {result.final_shape}")
    print(f"Saved file: {result.output_file}")
    print(f"Removed ISBNs file: {result.removed_isbns_file}")


def print_inventory_result_summary(result) -> None:
    print()
    print(f"Inventory source file: {result.source_file.name}")
    print(f"Matched ISBN overrides: {result.matched_updates:,}")
    print(f"Appended new ISBN rows: {result.appended_rows:,}")
    print(f"Rows removed by ISBN whitelist: {result.excluded_rows:,}")
    print(f"Final data shape: {result.final_shape}")
    print(f"Saved file: {result.output_file}")
    print(f"Removed ISBNs file: {result.removed_isbns_file}")


def _format_week_status(week) -> str:
    if week is None:
        return "None found"
    saturday_note = "" if week.weekday() == 5 else " (not a Saturday)"
    bookscan = bookscan_week(week)
    return f"Week {bookscan.week}, {week.strftime('%A, %B')} {week.day}, {bookscan.year}{saturday_note}"


def confirm_refresh_cache(latest_sql_week, latest_cache_week) -> bool:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        return messagebox.askyesno(
            "Barnes & Noble Cache Refresh",
            (
                "Have you uploaded the Sales and Inventory files to CBQ yet?\n\n"
                f"Latest SQL week: {_format_week_status(latest_sql_week)}\n"
                f"Current cache week: {_format_week_status(latest_cache_week)}\n\n"
                "Choose Yes to refresh the cache now."
            ),
            parent=root,
        )
    finally:
        root.destroy()


def confirm_build_report(latest_sql_week, latest_cache_week) -> bool:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        return messagebox.askyesno(
            "Run Barnes & Noble Rolling Report",
            (
                f"SQL is updated through: {_format_week_status(latest_sql_week)}\n"
                f"Current B&N cache through: {_format_week_status(latest_cache_week)}\n\n"
                "Would you like to proceed with the Barnes & Noble rolling report for this week?"
            ),
            parent=root,
        )
    finally:
        root.destroy()

def confirm_partial_pos_build(raw_folder: Path) -> bool:
    selection = collect_pos_source_files(raw_folder)
    print()
    print("Expected POS categories: toy, gift, cal, gc")
    print("POS files found:")
    if selection.source_files:
        for file_path in selection.source_files:
            print(f"  - {file_path.name}")
    else:
        print("  - None")

    if not selection.missing_keywords:
        return True

    print("Missing expected POS categories:")
    for keyword in selection.missing_keywords:
        print(f"  - {keyword}")

    if not selection.source_files:
        print("No usable POS files were found, so the combined POS file cannot be built.")
        return False

    while True:
        choice = input("Proceed with the available POS files only? [y/N]: ").strip().lower()
        if choice in {"y", "yes"}:
            return True
        if choice in {"", "n", "no"}:
            return False
        print("Please enter y or n.")


def run_menu(default_raw_folder: str | Path | None = None) -> None:
    raw_folder = resolve_raw_folder(default_raw_folder) if default_raw_folder else prompt_for_raw_folder()

    while True:
        print("\nBarnes & Noble Rolling Reports")
        print()
        print("    1. Build All Three (Steps 2, 3, & 4 together) (Main Step 1)")
        print("    2. Build Combined POS File")
        print("    3. Build working Sales file from Sales*.xlsx and pos_combined")
        print("    4. Build working Inventory file from Inventory*.xlsx and pos_combined")
        print("    5. Refresh Cache (After Uploading Sales & Inventory to CBQ) (Main Step 2)")
        print("    6. Build B&N rolling report (CB + DP Version)")
        print("    7. Build B&N rolling report (CB + DP Version) + DP versions (Main Step 3)")
        print("    8. Build local review B&N rolling report")
        print("    9. Exit")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice == "1":
            print("\nRunning Step 1 of 3: Build Combined POS File...")
            if not confirm_partial_pos_build(raw_folder):
                print("Build All Three cancelled before creating the combined POS file.")
                continue
            pos_result = build_combined_pos(raw_folder=raw_folder, allow_missing_files=True)
            print_result_summary(pos_result)

            print("\nRunning Step 2 of 3: Build working Sales file...")
            sales_result = build_sales_working_file(raw_folder=raw_folder)
            print_sales_result_summary(sales_result)

            print("\nRunning Step 3 of 3: Build working Inventory file...")
            inventory_result = build_inventory_working_file(raw_folder=raw_folder)
            print_inventory_result_summary(inventory_result)
            continue

        if choice == "2":
            print("\nRunning Build Combined POS File...")
            if not confirm_partial_pos_build(raw_folder):
                print("Combined POS build cancelled.")
                continue
            result = build_combined_pos(raw_folder=raw_folder, allow_missing_files=True)
            print_result_summary(result)
            continue

        if choice == "3":
            print("\nRunning Build working Sales file...")
            result = build_sales_working_file(raw_folder=raw_folder)
            print_sales_result_summary(result)
            continue

        if choice == "4":
            print("\nRunning Build working Inventory file...")
            result = build_inventory_working_file(raw_folder=raw_folder)
            print_inventory_result_summary(result)
            continue

        if choice == "5":
            latest_sql_week = get_latest_sql_week()
            latest_cache_week = get_latest_cache_week()
            print()
            print(
                "Latest week currently in CBQ: "
                f"{latest_sql_week.strftime('%Y-%m-%d') if latest_sql_week is not None else 'None found'}"
            )
            print(
                "Latest week currently in cache: "
                f"{latest_cache_week.strftime('%Y-%m-%d') if latest_cache_week is not None else 'None found'}"
            )
            if not confirm_refresh_cache(latest_sql_week, latest_cache_week):
                print("Cache refresh cancelled.")
                continue
            print("\nRefreshing caches...")
            result = refresh_caches_only(raw_folder=raw_folder)
            print_cache_refresh_summary(result)
            continue

        if choice == "6":
            latest_sql_week = get_latest_sql_week()
            latest_cache_week = get_latest_cache_week()
            if not confirm_build_report(latest_sql_week, latest_cache_week):
                print("B&N rolling report cancelled.")
                continue
            print("\nBuilding B&N rolling report (CB + DP Version)...")
            result = build_customer_sales_report(raw_folder=raw_folder)
            print_rolling_result_summary(result)
            continue

        if choice == "7":
            latest_sql_week = get_latest_sql_week()
            latest_cache_week = get_latest_cache_week()
            if not confirm_build_report(latest_sql_week, latest_cache_week):
                print("B&N rolling report cancelled.")
                continue
            print("\nBuilding B&N rolling report (CB + DP Version) + DP versions...")
            result = build_customer_sales_report(raw_folder=raw_folder, save_dp=True)
            print_rolling_result_summary(result)
            continue

        if choice == "8":
            latest_sql_week = get_latest_sql_week()
            latest_cache_week = get_latest_cache_week()
            if not confirm_build_report(latest_sql_week, latest_cache_week):
                print("B&N rolling report cancelled.")
                continue
            print("\nBuilding local review B&N rolling report...")
            result = build_customer_sales_report(raw_folder=raw_folder, local_only=True)
            print_rolling_result_summary(result)
            continue

        if choice in {"9", "q", "quit", "exit"}:
            return

        print("Invalid choice. Please select a valid option.")


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.raw_folder or args.output_file or args.preview:
        print("\nRunning Build Combined POS File...")
        result = build_combined_pos(
            raw_folder=args.raw_folder,
            output_file=args.output_file,
            allow_missing_files=args.allow_missing_pos_files,
        )
        print_result_summary(result)
        return

    run_menu(default_raw_folder=args.default_raw_folder)


if __name__ == "__main__":
    main()
