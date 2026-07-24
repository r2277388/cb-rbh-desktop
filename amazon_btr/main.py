from __future__ import annotations

import argparse
import re
import sys
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd

try:
    from amazon_btr.criteria import write_criteria_sheet
    from amazon_btr.workbook import build_active_asins, write_grouped_status_changes
except ModuleNotFoundError:
    from criteria import write_criteria_sheet
    from workbook import build_active_asins, write_grouped_status_changes


BTR_ROOT_FOLDER = Path(r"G:\SALES\Amazon\BORN TO RUN")
BTR_SOURCE_FOLDER = BTR_ROOT_FOLDER / "RBH_BTR_FILES"
REQUIRED_COLUMNS = {"ASIN", "Submission date", "Status", "Status description"}
REJECTED_ACTIVE_OFFER = "rejected: active offer"
ACTIVE_STATUS_ALIASES = {"active", "accepted"}

DOWNLOAD_INSTRUCTIONS = f"""
How to grab the BTR data

Data Source: Amazon Vendor Central

1) Navigate to:
   Orders -> Vendor Initiated Orders -> Go to the Program
2) Open Search Filters by clicking the down arrow.
3) Update filters:
   Status:
       Under Review
       Active
       Completed
   Submission Date:
       From: 6 months ago
       To: Today
4) Click Search.
5) Click Download.
6) Save the file to:
   {BTR_SOURCE_FOLDER}
""".strip()


def _export_timestamp(path: Path) -> datetime:
    match = re.match(
        r"(\d{4}-\d{2}-\d{2})T(\d{2})_(\d{2})_(\d{2})(?:\.\d+)?Z",
        path.stem,
        flags=re.IGNORECASE,
    )
    if match:
        return datetime.strptime(
            f"{match.group(1)} {match.group(2)}:{match.group(3)}:{match.group(4)}",
            "%Y-%m-%d %H:%M:%S",
        )
    return datetime.fromtimestamp(path.stat().st_mtime)


def find_latest_two_files(source_folder: Path = BTR_SOURCE_FOLDER) -> tuple[Path, Path]:
    files = [
        path
        for path in source_folder.glob("*.xlsx")
        if not path.name.startswith("~$")
    ]
    if len(files) < 2:
        raise FileNotFoundError(
            f"At least two BTR .xlsx files are required in {source_folder}."
        )
    previous, current = sorted(files, key=_export_timestamp)[-2:]
    return previous, current


def read_btr_file(path: Path) -> pd.DataFrame:
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message="Workbook contains no default style",
            category=UserWarning,
        )
        data = pd.read_excel(path, sheet_name=0, dtype={"ASIN": "string"})

    missing = REQUIRED_COLUMNS.difference(data.columns)
    if missing:
        raise ValueError(
            f"{path.name} is missing required column(s): {', '.join(sorted(missing))}"
        )
    return data


def clean_btr_data(raw: pd.DataFrame) -> pd.DataFrame:
    """Keep one meaningful, most-recent submission per ASIN."""
    data = raw.copy()
    data["_source_order"] = range(len(data))
    data["_submission_sort"] = pd.to_datetime(
        data["Submission date"], errors="coerce"
    )
    data["_asin_key"] = data["ASIN"].astype("string").str.strip()
    data["_status_key"] = data["Status"].astype("string").str.strip().str.casefold()

    if data["_asin_key"].isna().any() or data["_asin_key"].eq("").any():
        raise ValueError("The BTR source contains a blank ASIN.")
    if data["_submission_sort"].isna().any():
        bad_count = int(data["_submission_sort"].isna().sum())
        raise ValueError(f"The BTR source contains {bad_count} invalid submission date(s).")

    data = data.sort_values(
        ["_asin_key", "_submission_sort", "_source_order"], kind="stable"
    )
    selected_rows: list[pd.Series] = []
    for _, group in data.groupby("_asin_key", sort=False, dropna=False):
        latest = group.iloc[-1]
        if latest["_status_key"] == REJECTED_ACTIVE_OFFER:
            active_history = group[group["_status_key"].isin(ACTIVE_STATUS_ALIASES)]
            if not active_history.empty:
                latest = active_history.iloc[-1]
        selected_rows.append(latest)

    cleaned = pd.DataFrame(selected_rows)
    cleaned = cleaned.sort_values(
        ["_submission_sort", "_asin_key"], ascending=[False, True], kind="stable"
    )
    return cleaned.drop(
        columns=["_source_order", "_submission_sort", "_asin_key", "_status_key"]
    ).reset_index(drop=True)


def build_status_changes(
    previous_cleaned: pd.DataFrame, current_cleaned: pd.DataFrame
) -> pd.DataFrame:
    previous = previous_cleaned.copy()
    current = current_cleaned.copy()
    previous["_asin_key"] = previous["ASIN"].astype("string").str.strip()
    current["_asin_key"] = current["ASIN"].astype("string").str.strip()
    previous_by_asin = previous.set_index("_asin_key", drop=False)

    change_rows: list[dict] = []
    for _, current_row in current.iterrows():
        asin_key = current_row["_asin_key"]
        current_status = str(current_row["Status"]).strip()
        if asin_key not in previous_by_asin.index:
            change_type = "New ASIN"
            previous_status = ""
            previous_description = ""
        else:
            previous_row = previous_by_asin.loc[asin_key]
            previous_status = str(previous_row["Status"]).strip()
            previous_description = previous_row["Status description"]
            if previous_status.casefold() == current_status.casefold():
                continue
            change_type = "Status Changed"

        full_current_row = current_row.drop(labels=["_asin_key"]).to_dict()
        change_rows.append(
            {
                "Change Type": change_type,
                "Previous Status": previous_status,
                "Previous Status description": previous_description,
                **full_current_row,
            }
        )

    output_columns = [
        "Change Type",
        "Previous Status",
        "Previous Status description",
        *current_cleaned.columns,
    ]
    return pd.DataFrame(change_rows, columns=output_columns)


def _format_worksheet(writer: pd.ExcelWriter, sheet_name: str, data: pd.DataFrame) -> None:
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    header_format = workbook.add_format(
        {
            "bold": True,
            "font_color": "white",
            "bg_color": "#1F4E78",
            "border": 1,
            "text_wrap": True,
            "valign": "top",
        }
    )
    date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, max(len(data), 1), max(len(data.columns) - 1, 0))
    worksheet.set_row(0, 32)

    for column_number, column_name in enumerate(data.columns):
        worksheet.write(0, column_number, column_name, header_format)
        values = data[column_name].dropna().astype(str)
        sampled_width = max(
            [len(str(column_name)), *(len(value) for value in values.head(500))]
        )
        width = min(max(sampled_width + 2, 11), 45)
        cell_format = date_format if "date" in column_name.casefold() else None
        worksheet.set_column(column_number, column_number, width, cell_format)


def create_report(
    previous_file: Path,
    current_file: Path,
    output_folder: Path = BTR_ROOT_FOLDER,
) -> tuple[Path, int, int]:
    previous_file = Path(previous_file)
    current_file = Path(current_file)
    raw_current = read_btr_file(current_file)
    previous_cleaned = clean_btr_data(read_btr_file(previous_file))
    current_cleaned = clean_btr_data(raw_current)
    changes = build_status_changes(previous_cleaned, current_cleaned)

    report_date = _export_timestamp(current_file)
    active_asins = build_active_asins(current_cleaned, report_date)
    output_path = output_folder / f"{report_date:%Y_%m_%d}_BTR_Weekly_Review.xlsx"
    output_folder.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(
        output_path,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd",
        date_format="yyyy-mm-dd",
    ) as writer:
        for sheet_name, data in (
            ("Raw_BTR", raw_current),
            ("Cleaned", current_cleaned),
        ):
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            _format_worksheet(writer, sheet_name, data)
        write_grouped_status_changes(writer, changes)
        active_asins.to_excel(writer, sheet_name="Active ASINs", index=False)
        _format_worksheet(writer, "Active ASINs", active_asins)
        write_criteria_sheet(writer)

    return output_path, len(current_cleaned), len(changes)


def run_latest_report() -> Path:
    previous_file, current_file = find_latest_two_files()
    print("\nBTR report comparison:")
    print(f"  Previous: {previous_file}")
    print(f"  Current:  {current_file}")
    output_path, cleaned_count, change_count = create_report(
        previous_file, current_file
    )
    print(f"\nCreated: {output_path}")
    print(f"Cleaned ASINs: {cleaned_count:,}")
    print(f"Status changes/new ASINs: {change_count:,}")
    return output_path


def interactive_menu() -> None:
    while True:
        print("\nAmazon Born to Run (BTR) Weekly Review")
        print()
        print("    1. Create report using the latest two downloads")
        print("    2. View download instructions")
        print()
        print("    99. Back to Amazon Weekly Reporting")
        print()
        try:
            choice = input("Choose an option: ").strip().lower()
        except KeyboardInterrupt:
            print("\nReturning to Amazon Weekly Reporting.")
            return

        if choice in {"1", "run"}:
            try:
                run_latest_report()
            except (FileNotFoundError, OSError, ValueError, ImportError) as exc:
                print(f"\nUnable to create the BTR report: {exc}")
            continue
        if choice in {"2", "instructions", "help"}:
            print(f"\n{DOWNLOAD_INSTRUCTIONS}")
            continue
        if choice in {"99", "back", "b", "return", "menu"}:
            return
        print("Invalid choice. Please select a valid option.")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create the Amazon BTR weekly review.")
    parser.add_argument("--previous", type=Path, help="Older BTR export workbook.")
    parser.add_argument("--current", type=Path, help="Newer BTR export workbook.")
    parser.add_argument("--output-folder", type=Path, default=BTR_ROOT_FOLDER)
    parser.add_argument("--show-instructions", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    if args.show_instructions:
        print(DOWNLOAD_INSTRUCTIONS)
        return
    if bool(args.previous) != bool(args.current):
        raise SystemExit("--previous and --current must be supplied together.")
    if args.previous and args.current:
        output_path, cleaned_count, change_count = create_report(
            args.previous, args.current, args.output_folder
        )
        print(f"Created: {output_path}")
        print(f"Cleaned ASINs: {cleaned_count:,}")
        print(f"Status changes/new ASINs: {change_count:,}")
        return
    interactive_menu()


if __name__ == "__main__":
    main(sys.argv[1:])
