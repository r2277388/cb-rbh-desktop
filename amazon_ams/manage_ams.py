from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from monthly_campaigns import (  # noqa: E402
    AMS_HISTORY_PARQUET,
    CAMPAIGN_FOLDER,
    final_report_file,
    parse_missing_asin_overrides,
    parse_period_from_filename,
    run,
)


def choose_csv_file(default_folder: Path = CAMPAIGN_FOLDER) -> Path:
    files = sorted(default_folder.glob("*.csv"), key=lambda path: path.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"No CSV files found in {default_folder}")

    print()
    print("Available AMS monthly campaign CSV files:")
    for index, path in enumerate(files, start=1):
        print(f"  {index}. {path.name}")
    print()
    choice = input("Choose a file number, or paste a full CSV path: ").strip()
    if choice.isdigit():
        try:
            return files[int(choice) - 1]
        except IndexError as exc:
            raise ValueError("Invalid file selection.") from exc

    path = Path(choice.strip('"'))
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    return path


def collect_missing_asin_overrides() -> dict[str, list[str]]:
    print()
    print("Optional missing-ASIN overrides.")
    print('Format: Campaign name=ASIN1,ASIN2')
    print("Press Enter with no value when done. You can also leave this blank and answer prompts during processing.")
    values: list[str] = []
    while True:
        value = input("Missing-ASIN override: ").strip()
        if not value:
            break
        values.append(value)
    return parse_missing_asin_overrides(values)


def run_new_month() -> None:
    source_file = choose_csv_file()
    overrides = collect_missing_asin_overrides()
    period = parse_period_from_filename(source_file)
    prior_file = input("Prior month CSV override for comparison [blank to use history parquet]: ").strip()
    prior_input = Path(prior_file.strip('"')) if prior_file else None
    output = run(
        source_file,
        prompt=True,
        missing_asin_overrides=overrides,
        prior_input=prior_input,
        save_history=True,
        write_publisher_reports=True,
    )
    print()
    print(f"Finished {period}.")
    print(f"  ALL report: {output}")
    print(f"  PWP report: {final_report_file(period, 'PWP')}")


def replace_month() -> None:
    source_file = choose_csv_file()
    overrides = collect_missing_asin_overrides()
    period = parse_period_from_filename(source_file)
    print()
    print(f"This will replace {period} in the AMS history parquet and regenerate reports.")
    confirm = input("Continue? (y/n): ").strip().lower()
    if confirm not in {"y", "yes"}:
        print("Cancelled.")
        return
    prior_file = input("Prior month CSV override for comparison [blank to use history parquet]: ").strip()
    prior_input = Path(prior_file.strip('"')) if prior_file else None
    run(
        source_file,
        prompt=True,
        missing_asin_overrides=overrides,
        prior_input=prior_input,
        save_history=True,
        write_publisher_reports=True,
    )


def run_report_only() -> None:
    source_file = choose_csv_file()
    overrides = collect_missing_asin_overrides()
    prior_file = input("Prior month CSV override for comparison [blank to use history parquet]: ").strip()
    prior_input = Path(prior_file.strip('"')) if prior_file else None
    run(
        source_file,
        prompt=True,
        missing_asin_overrides=overrides,
        prior_input=prior_input,
        save_history=False,
        write_publisher_reports=True,
    )


def print_history_status() -> None:
    print()
    print(f"AMS history parquet: {AMS_HISTORY_PARQUET}")
    if not AMS_HISTORY_PARQUET.exists():
        print("  Status: missing")
        return
    history = pd.read_parquet(AMS_HISTORY_PARQUET)
    print(f"  Rows:    {len(history):,}")
    if "period" in history.columns:
        periods = sorted(history["period"].dropna().astype(str).unique())
        print(f"  Periods: {', '.join(periods)}")


def main() -> None:
    while True:
        print()
        print("Amazon AMS Manager")
        print()
        print("    1. Run new month, save to parquet, and create reports")
        print("    2. Replace a past month in parquet and recreate reports")
        print("    3. Run report only, without saving parquet")
        print("    4. Show AMS parquet status")
        print("    5. Back to launcher")
        print()
        choice = input("Choose an option: ").strip().lower()
        try:
            if choice == "1":
                run_new_month()
            elif choice == "2":
                replace_month()
            elif choice == "3":
                run_report_only()
            elif choice == "4":
                print_history_status()
            elif choice in {"5", "back", "b", "exit", "quit", "q"}:
                return
            else:
                print("Invalid choice.")
        except Exception as exc:
            print(f"Operation failed: {exc}")


if __name__ == "__main__":
    main()
