import importlib
import re
import subprocess
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

import UPDATE_ams_config as ams_config
from loader_asin_mapping import load_asin_mapping
from loader_item import upload_item
from loader_monthly_reports import load_monthly_data


CONFIG_PATH = Path(__file__).with_name("UPDATE_ams_config.py")
BASELINE_YEAR = "2025"


def reload_config():
    global ams_config
    ams_config = importlib.reload(ams_config)
    return ams_config.tab_dict, ams_config.month_list


def write_config(tab_dict):
    lines = [
        "# Auto-generated AMS config; keep keys in YYYY-MM format.",
        "tab_dict = {",
    ]

    for month in sorted(tab_dict.keys()):
        entry = tab_dict[month]
        tab = entry.get("tab", "USE_main")
        skiprows = int(entry.get("skiprows", 1))
        file_path = str(entry.get("file", ""))
        lines.extend(
            [
                f'    "{month}": {{',
                f'        "tab": "{tab}",',
                f'        "skiprows": {skiprows},',
                f'        "file": r"{file_path}",',
                "    },",
            ]
        )

    lines.extend(["}", "", "month_list = sorted(tab_dict.keys())", ""])
    CONFIG_PATH.write_text("\n".join(lines), encoding="utf-8")


def read_columns_from_entry(entry):
    df = pd.read_excel(
        entry["file"],
        sheet_name=entry["tab"],
        skiprows=entry["skiprows"],
        header=0,
        engine="openpyxl",
    )
    df.columns = df.columns.str.strip().str.lower()
    return set(df.columns)


def baseline_columns(tab_dict):
    baseline_months = sorted([m for m in tab_dict if m.startswith(f"{BASELINE_YEAR}-")])
    if not baseline_months:
        raise RuntimeError(f"No baseline months found for {BASELINE_YEAR}.")

    month_to_cols = {}
    for m in baseline_months:
        month_to_cols[m] = read_columns_from_entry(tab_dict[m])

    # Pick the most common baseline schema to avoid single-month anomalies.
    signatures = Counter(tuple(sorted(cols)) for cols in month_to_cols.values())
    expected_tuple, _ = signatures.most_common(1)[0]
    expected = set(expected_tuple)

    if len(signatures) > 1:
        print("WARNING: baseline year has more than one column shape.")
        print("Using most common baseline schema for validation.")

    return expected


def validate_new_month_columns(month, tab, skiprows, file_path, tab_dict):
    expected = baseline_columns(tab_dict)
    candidate = {"tab": tab, "skiprows": skiprows, "file": file_path}
    actual = read_columns_from_entry(candidate)

    missing = sorted(expected - actual)
    extra = sorted(actual - expected)

    print("\nExpected columns (baseline):")
    print(sorted(expected))
    print("\nActual columns (new file):")
    print(sorted(actual))

    if not missing and not extra:
        print(f"\nValidation passed for {month}: columns match baseline.")
        return True

    print(f"\nValidation failed for {month}:")
    print(f"Missing columns: {missing}")
    print(f"Extra columns: {extra}")
    return False


def print_last_10_months():
    _, month_list = reload_config()
    last_10 = sorted(month_list, reverse=True)[:10]
    print("\nLast 10 configured months:")
    for m in last_10:
        print(m)


def remove_month():
    tab_dict, month_list = reload_config()
    print("\nConfigured months:")
    print(", ".join(sorted(month_list)))
    month = input("Enter month to remove (YYYY-MM): ").strip()
    if month not in tab_dict:
        print(f"Month not found: {month}")
        return

    confirm = input(f"Remove {month} from tab_dict? (y/n): ").strip().lower()
    if confirm not in {"y", "yes"}:
        print("Cancelled.")
        return

    del tab_dict[month]
    write_config(tab_dict)
    print(f"Removed {month} from UPDATE_ams_config.py")


def add_new_month():
    tab_dict, _ = reload_config()

    month = input("Month (YYYY-MM): ").strip()
    if not re.fullmatch(r"\d{4}-\d{2}", month):
        print("Invalid month format. Use YYYY-MM.")
        return
    if month in tab_dict:
        print(f"{month} already exists in tab_dict.")
        return

    tab = input('Tab name [default USE_main]: ').strip() or "USE_main"
    skiprows_raw = input("Rows to skip [default 1]: ").strip() or "1"
    try:
        skiprows = int(skiprows_raw)
    except ValueError:
        print("Rows to skip must be an integer.")
        return

    file_path = input("Full file path: ").strip()
    if not file_path:
        print("File path is required.")
        return
    if not Path(file_path).exists():
        print(f"File not found: {file_path}")
        return

    print("\nRunning column validation before adding...")
    ok = validate_new_month_columns(month, tab, skiprows, file_path, tab_dict)
    if not ok:
        print("Not added due to column mismatch.")
        return

    tab_dict[month] = {"tab": tab, "skiprows": skiprows, "file": file_path}
    write_config(tab_dict)
    print(f"Added {month} to UPDATE_ams_config.py")


def run_ams_full():
    print("Running full amazon_ams process...")
    subprocess.run([sys.executable, str(Path(__file__).with_name("main.py"))], check=True)


def run_ams_incremental():
    tab_dict, month_list = reload_config()
    if not month_list:
        print("No months in config.")
        return

    default_month = sorted(month_list)[-1]
    month = input(f"Month to process [default {default_month}]: ").strip() or default_month
    if month not in tab_dict:
        print(f"Month not found in tab_dict: {month}")
        return

    print(f"Processing incremental month: {month}")
    asin_mapping = load_asin_mapping()
    item_df = upload_item()

    df_month = load_monthly_data(tab_dict[month], asin_mapping, month)
    if not item_df.empty:
        df_month = pd.merge(df_month, item_df, on="ISBN", how="left")

    output_pickle = Path(__file__).with_name("combined_amazon_ads_by_asin.pkl")
    output_excel = Path(__file__).with_name("combined_amazon_ads_by_asin.xlsx")

    if output_pickle.exists():
        existing = pd.read_pickle(output_pickle)
        if "period" in existing.columns:
            existing = existing[existing["period"] != month]
        combined = pd.concat([existing, df_month], ignore_index=True)
    else:
        combined = df_month

    combined.to_pickle(output_pickle)
    combined.to_excel(output_excel, index=False)
    print(f"Updated outputs with {month}:")
    print(f"- {output_pickle}")
    print(f"- {output_excel}")


def run_ams():
    print("\nRun AMS")
    print()
    print("    1. Incremental (new/additional month only)")
    print("    2. Full rerun (all configured months)")
    print("    3. Back")
    print()
    choice = input("Choose an option: ").strip().lower()
    if choice == "1":
        run_ams_incremental()
    elif choice == "2":
        run_ams_full()
    elif choice in {"3", "back", "b"}:
        return
    else:
        print("Invalid choice.")


def main():
    while True:
        print("\nAmazon AMS Manager")
        print()
        print("    1. Show last 10 configured months")
        print("    2. Remove month from tab_dict")
        print("    3. Add/process new month (validate before add)")
        print("    4. Run AMS processing")
        print("    5. Back to launcher")
        print()
        choice = input("Choose an option: ").strip().lower()

        try:
            if choice == "1":
                print_last_10_months()
            elif choice == "2":
                remove_month()
            elif choice == "3":
                add_new_month()
            elif choice == "4":
                run_ams()
            elif choice in {"5", "back", "b", "exit", "quit", "q"}:
                return
            else:
                print("Invalid choice.")
        except Exception as e:
            print(f"Operation failed: {e}")


if __name__ == "__main__":
    main()
