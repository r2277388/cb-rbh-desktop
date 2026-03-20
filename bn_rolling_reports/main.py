from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from config import BASE_FOLDER
from pos_combiner import (
    build_combined_pos,
    format_output_filename,
    get_candidate_raw_folders,
    parse_week_ending,
    preview_dataframe,
    resolve_raw_folder,
)
from inventory_working import build_inventory_working_file
from sales_working import build_sales_working_file


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
    return parser


def prompt_for_raw_folder() -> Path:
    candidates = get_candidate_raw_folders()
    if candidates:
        latest = candidates[-1]
        default_name = latest.name
        prompt = (
            f"Raw folder [{default_name}] under {BASE_FOLDER} "
            "(press Enter to use default or paste a full folder path): "
        )
        user_value = input(prompt).strip()
        return resolve_raw_folder(user_value or latest)

    user_value = input(
        "Mapped base folder was not available. Paste the full raw folder path: "
    ).strip()
    return resolve_raw_folder(user_value)


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


def view_existing_output() -> None:
    raw_folder = prompt_for_raw_folder()
    week_ending = parse_week_ending(raw_folder.name)
    output_file = raw_folder / format_output_filename(week_ending)

    if not output_file.exists():
        print(f"Combined file not found yet: {output_file}")
        return

    df = pd.read_excel(output_file, dtype={"ISBN": "string", "Imprint": "string"})
    print()
    print(f"Showing first 25 rows from {output_file.name}")
    print(preview_dataframe(df))


def run_menu() -> None:
    while True:
        print("\nBarnes & Noble Rolling Reports")
        print()
        print("    1. Build combined POS file for the raw folder")
        print("    2. View the existing combined POS file on screen")
        print("    3. Build working Sales file from Sales*.xlsx and pos_combined")
        print("    4. Build working Inventory file from Inventory*.xlsx and pos_combined")
        print("    5. Exit")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice == "1":
            raw_folder = prompt_for_raw_folder()
            result = build_combined_pos(raw_folder=raw_folder)
            print_result_summary(result)
            print()
            print("First 25 rows:")
            print(preview_dataframe(result.dataframe))
            continue

        if choice == "2":
            view_existing_output()
            continue

        if choice == "3":
            raw_folder = prompt_for_raw_folder()
            result = build_sales_working_file(raw_folder=raw_folder)
            print()
            print(f"Sales source file: {result.source_file.name}")
            print(f"Matched ISBN overrides: {result.matched_updates:,}")
            print(f"Appended new ISBN rows: {result.appended_rows:,}")
            print(f"Final data shape: {result.final_shape}")
            print(f"Saved file: {result.output_file}")
            continue

        if choice == "4":
            raw_folder = prompt_for_raw_folder()
            result = build_inventory_working_file(raw_folder=raw_folder)
            print()
            print(f"Inventory source file: {result.source_file.name}")
            print(f"Matched ISBN overrides: {result.matched_updates:,}")
            print(f"Appended new ISBN rows: {result.appended_rows:,}")
            print(f"Final data shape: {result.final_shape}")
            print(f"Saved file: {result.output_file}")
            continue

        if choice in {"5", "q", "quit", "exit"}:
            return

        print("Invalid choice. Please select a valid option.")


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.raw_folder or args.output_file or args.preview:
        result = build_combined_pos(raw_folder=args.raw_folder, output_file=args.output_file)
        print_result_summary(result)
        if args.preview:
            print()
            print("First 25 rows:")
            print(preview_dataframe(result.dataframe))
        return

    run_menu()


if __name__ == "__main__":
    main()
