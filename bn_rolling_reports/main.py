from __future__ import annotations

import argparse
from pathlib import Path
from tkinter import Tk, filedialog

from config import BASE_FOLDER
from pos_combiner import (
    build_combined_pos,
    get_candidate_raw_folders,
    resolve_raw_folder,
)
from inventory_working import build_inventory_working_file, find_inventory_source_file
from sales_working import build_sales_working_file, find_sales_source_file


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


def run_menu(default_raw_folder: str | Path | None = None) -> None:
    raw_folder = resolve_raw_folder(default_raw_folder) if default_raw_folder else prompt_for_raw_folder()

    while True:
        print("\nBarnes & Noble Rolling Reports")
        print()
        print("    1. Build All Three")
        print("    2. Build Combined POS File")
        print("    3. Build working Sales file from Sales*.xlsx and pos_combined")
        print("    4. Build working Inventory file from Inventory*.xlsx and pos_combined")
        print("    5. Exit")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice == "1":
            print("\nRunning Step 1 of 3: Build Combined POS File...")
            pos_result = build_combined_pos(raw_folder=raw_folder)
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
            result = build_combined_pos(raw_folder=raw_folder)
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

        if choice in {"5", "q", "quit", "exit"}:
            return

        print("Invalid choice. Please select a valid option.")


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.raw_folder or args.output_file or args.preview:
        print("\nRunning Build Combined POS File...")
        result = build_combined_pos(raw_folder=args.raw_folder, output_file=args.output_file)
        print_result_summary(result)
        return

    run_menu(default_raw_folder=args.default_raw_folder)


if __name__ == "__main__":
    main()
