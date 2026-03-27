import shutil
import sys
import re
import argparse
import getpass
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from paths import process_paths
from bn_rolling_reports.isbn_utils import normalize_isbn, normalize_isbn_series
from shared.db.connection import get_connection
from shared.db.query_runner import fetch_data_from_db


DESTINATION_FOLDER = process_paths.CONSOLIDATED_INVENTORY_VERTICALIZATION_FOLDER
DEFAULT_SOURCE_FOLDER = Path(r"F:\2026\Oracle Reports")
ORG_CODES = {"cbc", "hbg", "cbp"}
PICKLE_FILE = DESTINATION_FOLDER / "ConsolidatedInventory.pkl"
SHAREPOINT_FOLDER_SUFFIX = Path(
    r"OneDrive - chroniclebooks.com\CB Fabric and Power BI Project - Excel Data and Reports Needed for Power BI"
)
KNOWN_SHAREPOINT_FOLDERS = {
    "rbh": Path(
        r"C:\Users\rbh\OneDrive - chroniclebooks.com\CB Fabric and Power BI Project - Excel Data and Reports Needed for Power BI"
    ),
    "sdm": Path(
        r"C:\Users\sdm\OneDrive - chroniclebooks.com\CB Fabric and Power BI Project - Excel Data and Reports Needed for Power BI"
    ),
}
PUBLISHER_CACHE_DIR = DESTINATION_FOLDER / "cache"
PUBLISHER_LOOKUP_SQL = """
SELECT
    i.ITEM_TITLE AS ISBN,
    i.SHORT_TITLE AS Title,
    i.PUBLISHER_CODE AS Publisher
FROM
    ebs.item i
WHERE
    i.ISBN IS NOT NULL
    AND i.PUBLISHING_GROUP NOT IN ('MKT', 'ZZZ')
"""


def choose_source_file() -> Path | None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = filedialog.askopenfilename(
            title="Choose the monthly consolidated inventory workbook",
            initialdir=str(DEFAULT_SOURCE_FOLDER),
            filetypes=[
                ("Excel files", "*.xlsx"),
            ],
        )
    finally:
        root.destroy()

    if not selected:
        return None

    return Path(selected)


def choose_consolidated_file() -> Path | None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = filedialog.askopenfilename(
            title="Choose the consolidated inventory workbook to verticalize",
            initialdir=str(DESTINATION_FOLDER),
            filetypes=[
                ("Excel files", "*.xlsx"),
            ],
        )
    finally:
        root.destroy()

    if not selected:
        return None

    return Path(selected)


def prompt_for_period() -> str | None:
    while True:
        print()
        period = input("Enter the month to save as (yyyymm), or type 'back': ").strip()

        if not period:
            print("A period is required.")
            continue

        if period.lower() in {"back", "b", "return", "menu"}:
            return None

        if len(period) != 6 or not period.isdigit():
            print("Invalid period. Please enter the month as yyyymm.")
            continue

        month = int(period[4:])
        if month < 1 or month > 12:
            print("Invalid period. Month must be between 01 and 12.")
            continue

        return period


def prompt_for_month_count() -> int | None:
    while True:
        print()
        raw_value = input("Enter how many months to display, or type 'back': ").strip()

        if not raw_value:
            print("A month count is required.")
            continue

        if raw_value.lower() in {"back", "b", "return", "menu"}:
            return None

        if not raw_value.isdigit():
            print("Invalid entry. Please enter a whole number.")
            continue

        month_count = int(raw_value)
        if month_count <= 0:
            print("Month count must be greater than zero.")
            continue

        return month_count


def prompt_for_row_count() -> int | None:
    while True:
        print()
        raw_value = input("Enter how many rows to display, or type 'back': ").strip()

        if not raw_value:
            print("A row count is required.")
            continue

        if raw_value.lower() in {"back", "b", "return", "menu"}:
            return None

        if not raw_value.isdigit():
            print("Invalid entry. Please enter a whole number.")
            continue

        row_count = int(raw_value)
        if row_count <= 0:
            print("Row count must be greater than zero.")
            continue

        return row_count


def confirm_overwrite(destination: Path) -> bool:
    while True:
        print()
        print(f"The output file already exists: {destination}")
        print()
        print("    1. Overwrite existing file")
        print("    2. Return to prior menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "o", "overwrite"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def add_new_month_file():
    DESTINATION_FOLDER.mkdir(parents=True, exist_ok=True)

    period = prompt_for_period()
    if period is None:
        return

    source_file = choose_source_file()
    if source_file is None:
        print("No file was selected.")
        return

    if not source_file.exists():
        print(f"Selected file was not found: {source_file}")
        return

    destination = DESTINATION_FOLDER / f"All_Consolidated_Inventories_{period}.xlsx"

    print()
    print("Consolidate Inventory Verticalization will use these files:")
    print(f"  Source workbook:  {source_file}")
    print(f"  Save period:      {period}")
    print(f"  Output workbook:  {destination}")

    if destination.exists() and not confirm_overwrite(destination):
        return

    shutil.copy2(source_file, destination)
    print()
    print(f"Saved: {destination}")
    inspect_saved_workbook(destination)


def show_last_10_depot_files():
    DESTINATION_FOLDER.mkdir(parents=True, exist_ok=True)

    files = sorted(
        DESTINATION_FOLDER.glob("All_Consolidated_Inventories_*.xlsx"),
        key=lambda path: path.name,
        reverse=True,
    )[:10]

    print()
    print("Last 10 Consolidated Inventories in the Depot:")
    if not files:
        print("  No consolidated inventory files were found.")
        return

    for file_path in files:
        modified = datetime.fromtimestamp(file_path.stat().st_mtime).strftime("%Y-%m-%d %I:%M:%S %p")
        print(f"  {file_path.name}  |  {modified}")


def view_depot_location():
    print()
    print("ConInv Depot location:")
    print(f"  {DESTINATION_FOLDER}")


def run_depot_menu():
    while True:
        print("\nBash Depot Files")
        print()
        print("    1. Add Monthly Consolidated File to Bash Depot")
        print("    2. View Last 10 Consolidated Inventories in the Depot")
        print("    3. View Location of Bash Depot")
        print("    4. Back")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice == "1":
            add_new_month_file()
            continue

        if choice == "2":
            show_last_10_depot_files()
            continue

        if choice == "3":
            view_depot_location()
            continue

        if choice in {"4", "back", "b", "return", "menu"}:
            return

        print("Invalid choice. Please select a valid option.")


def inspect_saved_workbook(workbook_path: Path):
    excel_file = pd.ExcelFile(workbook_path)
    sheet_names = excel_file.sheet_names
    selected_sheet = choose_data_sheet(excel_file)

    print()
    print("Workbook review:")
    print(f"  Workbook:        {workbook_path}")
    print(f"  Worksheet count: {len(sheet_names)}")
    print("  Worksheets:")
    for sheet_name in sheet_names:
        print(f"    {sheet_name}")

    if selected_sheet is None:
        print("  Data worksheet:  Unable to determine automatically")
        return

    df = pd.read_excel(workbook_path, sheet_name=selected_sheet)
    print(f"  Data worksheet:  {selected_sheet}")
    print(f"  Rows loaded:     {len(df):,}")
    print(f"  Columns loaded:  {len(df.columns):,}")


def choose_data_sheet(excel_file: pd.ExcelFile) -> str | None:
    preferred_prefix = "all_consolidated_inventories_v_"
    preferred_matches = [
        sheet_name
        for sheet_name in excel_file.sheet_names
        if sheet_name.lower().startswith(preferred_prefix)
    ]
    if preferred_matches:
        return preferred_matches[0]

    best_sheet = None
    best_score = -1
    for sheet_name in excel_file.sheet_names:
        preview_df = pd.read_excel(excel_file, sheet_name=sheet_name, nrows=25)
        non_empty_columns = int(preview_df.notna().any().sum())
        non_empty_rows = int(preview_df.dropna(how="all").shape[0])
        score = non_empty_columns * 1000 + non_empty_rows
        if score > best_score:
            best_score = score
            best_sheet = sheet_name

    return best_sheet


def normalize_header_value(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def infer_period_from_filename(file_path: Path) -> str | None:
    match = re.search(r"(\d{6})", file_path.stem)
    if not match:
        return None
    return match.group(1)


def find_header_row(raw_df: pd.DataFrame) -> int:
    for row_index in range(min(15, len(raw_df))):
        row_values = [normalize_header_value(value).lower() for value in raw_df.iloc[row_index]]
        if "item" in row_values:
            if row_index + 1 < len(raw_df):
                next_row_values = [
                    normalize_header_value(value).lower() for value in raw_df.iloc[row_index + 1]
                ]
                if "item" in next_row_values:
                    return row_index + 1
            return row_index
    raise ValueError("Unable to find the header row containing 'Item'.")


def build_vertical_dataframe(raw_df: pd.DataFrame) -> pd.DataFrame:
    header_row_index = find_header_row(raw_df)
    if header_row_index < 1:
        raise ValueError("The workbook is missing the organization header row above 'Item'.")

    org_row = raw_df.iloc[header_row_index - 1].copy()
    base_row = raw_df.iloc[header_row_index]
    data_df = raw_df.iloc[header_row_index + 1 :].copy()
    data_df = data_df.reset_index(drop=True)

    # Merged Excel headers often put the org label only in the first column of a span.
    # Carry labels to the right so the paired $ / units columns can both be detected.
    last_org_label = ""
    for column_index in range(raw_df.shape[1]):
        current_label = normalize_header_value(org_row.iloc[column_index]).lower()
        if current_label in ORG_CODES:
            last_org_label = current_label
        elif current_label:
            last_org_label = ""
        elif current_label == "" and last_org_label:
            org_row.iloc[column_index] = last_org_label

    isbn_column = None
    org_column_map: dict[str, dict[str, int]] = {org: {} for org in ORG_CODES}

    for column_index in range(raw_df.shape[1]):
        base_label = normalize_header_value(base_row.iloc[column_index]).lower()
        org_label = normalize_header_value(org_row.iloc[column_index]).lower()

        if base_label == "item":
            isbn_column = column_index
            continue

        if org_label not in ORG_CODES:
            continue

        if base_label == "$":
            org_column_map[org_label]["value"] = column_index
        elif base_label == "units":
            org_column_map[org_label]["inventory"] = column_index

    if isbn_column is None:
        raise ValueError("Unable to find the 'Item' column in the selected worksheet.")

    missing_orgs = [
        org
        for org, columns in org_column_map.items()
        if "value" not in columns or "inventory" not in columns
    ]
    if missing_orgs:
        missing_text = ", ".join(sorted(missing_orgs))
        raise ValueError(f"Unable to find both value and units columns for: {missing_text}")

    base_vertical = pd.DataFrame()
    base_vertical["ISBN"] = data_df.iloc[:, isbn_column]
    base_vertical["ISBN"] = base_vertical["ISBN"].astype("string").str.strip()
    base_vertical = base_vertical[base_vertical["ISBN"].notna() & (base_vertical["ISBN"] != "")]
    base_vertical = base_vertical[~base_vertical["ISBN"].str.lower().eq("grand total")]
    base_vertical = base_vertical[~base_vertical["ISBN"].str.lower().eq("item")]

    vertical_frames = []
    for org in sorted(ORG_CODES):
        org_frame = pd.DataFrame()
        org_frame["ISBN"] = base_vertical["ISBN"]
        org_frame["ORG"] = org
        org_frame["Value"] = pd.to_numeric(
            data_df.iloc[base_vertical.index, org_column_map[org]["value"]],
            errors="coerce",
        )
        org_frame["Inventory"] = pd.to_numeric(
            data_df.iloc[base_vertical.index, org_column_map[org]["inventory"]],
            errors="coerce",
        )
        vertical_frames.append(org_frame)

    vertical_df = pd.concat(vertical_frames, ignore_index=True)
    vertical_df["Value"] = vertical_df["Value"].fillna(0)
    vertical_df["Inventory"] = vertical_df["Inventory"].fillna(0)
    vertical_df = vertical_df[
        (vertical_df["Value"] != 0) | (vertical_df["Inventory"] != 0)
    ].copy()
    vertical_df["ISBN"] = normalize_isbn_series(vertical_df["ISBN"])
    vertical_df = vertical_df[vertical_df["ISBN"].notna()].copy()
    vertical_df = vertical_df.sort_values(["ISBN", "ORG"]).reset_index(drop=True)
    return vertical_df[["ISBN", "ORG", "Value", "Inventory"]]


def build_metrics_dataframe(vertical_df: pd.DataFrame) -> pd.DataFrame:
    summary_rows = [
        {"Metric": "Row Count", "Value": int(len(vertical_df))},
        {"Metric": "Total Value", "Value": float(vertical_df["Value"].sum())},
        {"Metric": "Total Inventory", "Value": float(vertical_df["Inventory"].sum())},
    ]

    org_totals = vertical_df.groupby("ORG")[["Value", "Inventory"]].sum()
    for org, row in org_totals.iterrows():
        summary_rows.append({"Metric": f"Total {org.upper()} Value", "Value": float(row["Value"])})
        summary_rows.append(
            {"Metric": f"Total {org.upper()} Inventory", "Value": float(row["Inventory"])}
        )

    return pd.DataFrame(summary_rows)


def print_publisher_org_summary(vertical_df: pd.DataFrame):
    summary_df = vertical_df.copy()
    summary_df["PublisherGroup"] = summary_df["Publisher"].fillna("").map(
        lambda publisher: "Chronicle" if str(publisher).strip() == "Chronicle" else "DP"
    )
    grouped = (
        summary_df.groupby(["PublisherGroup", "ORG"])[["Value", "Inventory"]]
        .sum()
        .reset_index()
        .sort_values(["PublisherGroup", "ORG"])
    )

    print()
    print("Publisher Summary:")
    for publisher_group in ["Chronicle", "DP"]:
        publisher_rows = grouped[grouped["PublisherGroup"] == publisher_group]
        if publisher_rows.empty:
            continue
        print(f"  {publisher_group}:")
        for _, row in publisher_rows.iterrows():
            print(
                f"    {row['ORG'].upper()} Value: {int(round(row['Value'])):,} | "
                f"{row['ORG'].upper()} Inventory: {int(round(row['Inventory'])):,}"
            )


def check_specific_period():
    if not PICKLE_FILE.exists():
        print()
        print(f"Pickle file not found: {PICKLE_FILE}")
        return

    period = prompt_for_period()
    if period is None:
        return

    df = load_existing_vertical_pickle()
    period_df = df[df["period"].astype(str) == period].copy()
    if period_df.empty:
        print()
        print(f"No rows were found in the pickle for period {period}.")
        return

    publisher_groups = [
        ("Chronicle", period_df[period_df["Publisher"].fillna("").astype(str).str.strip() == "Chronicle"].copy()),
        ("DP", period_df[period_df["Publisher"].fillna("").astype(str).str.strip() != "Chronicle"].copy()),
    ]

    print()
    print(f"Totals for period {period}:")
    for publisher_name, publisher_df in publisher_groups:
        if publisher_df.empty:
            continue

        summary = (
            publisher_df.groupby("ORG")[["Value", "Inventory"]]
            .sum()
            .reset_index()
            .sort_values("ORG")
        )

        total_value = int(round(summary["Value"].sum()))
        total_inventory = int(round(summary["Inventory"].sum()))
        total_uc_avg = (
            summary["Value"].sum() / summary["Inventory"].sum()
            if summary["Inventory"].sum()
            else 0
        )

        print()
        print(f"  {publisher_name}:")
        print("  " + "-" * 61)
        print(f"  {'ORG':<8}{'Value':>14}{'Inventory':>16}{'U/C Avg':>13}")
        print("  " + "-" * 61)
        for _, row in summary.iterrows():
            org = str(row["ORG"]).upper()
            value = int(round(row["Value"]))
            inventory = int(round(row["Inventory"]))
            uc_avg = (row["Value"] / row["Inventory"]) if row["Inventory"] else 0
            print(f"  {org:<8}{value:>14,}{inventory:>16,}{uc_avg:>13.2f}")
        print("  " + "-" * 61)
        print(f"  {'TOTAL':<8}{total_value:>14,}{total_inventory:>16,}{total_uc_avg:>13.2f}")


def preview_verticalized_coninv():
    if not PICKLE_FILE.exists():
        print()
        print(f"Pickle file not found: {PICKLE_FILE}")
        return

    period = prompt_for_period()
    if period is None:
        return

    row_count = prompt_for_row_count()
    if row_count is None:
        return

    df = load_existing_vertical_pickle()
    period_df = df[df["period"].astype(str) == period].copy()
    if period_df.empty:
        print()
        print(f"No rows were found in the pickle for period {period}.")
        return

    period_df = period_df.sort_values(["ISBN", "ORG"]).reset_index(drop=True)
    preview_df = period_df.head(row_count).copy()

    print()
    print(f"Top {len(preview_df):,} rows of verticalized ConInv for SharePoint ({period}):")
    print(preview_df.to_string(index=False))


def check_last_n_months():
    if not PICKLE_FILE.exists():
        print()
        print(f"Pickle file not found: {PICKLE_FILE}")
        return

    df = load_existing_vertical_pickle()
    if df.empty:
        print()
        print("The pickle file is empty.")
        return

    df = df.copy()
    df["period"] = df["period"].astype(str)
    df = df[df["Publisher"].fillna("").astype(str).str.strip() == "Chronicle"].copy()
    if df.empty:
        print()
        print("No Chronicle rows were found in the pickle.")
        return

    month_count = prompt_for_month_count()
    if month_count is None:
        return

    recent_periods = sorted(df["period"].unique().tolist())[-month_count:]
    recent_df = df[df["period"].isin(recent_periods)].copy()

    summary_rows = []
    for period in sorted(recent_periods, reverse=True):
        period_df = recent_df[recent_df["period"] == period].copy()
        row = {
            "period": period,
            "row_count": int(len(period_df)),
        }
        for org in ["cbc", "hbg", "cbp"]:
            org_df = period_df[period_df["ORG"].astype(str).str.lower() == org]
            org_value = float(org_df["Value"].sum())
            org_inventory = float(org_df["Inventory"].sum())
            row[f"{org}_value"] = int(round(org_value))
            row[f"{org}_inventory"] = int(round(org_inventory))
            row[f"{org}_uc"] = (org_value / org_inventory) if org_inventory else 0
        total_value = float(period_df["Value"].sum())
        total_inventory = float(period_df["Inventory"].sum())
        row["total_value"] = int(round(total_value))
        row["total_inventory"] = int(round(total_inventory))
        row["total_uc"] = (total_value / total_inventory) if total_inventory else 0
        summary_rows.append(row)

    print()
    print(f"Last {len(recent_periods)} months (CB Only):")
    print("  " + "-" * 160)
    print(
        f"  {'Period':<8}"
        f"{'Rows':>8}"
        f"{'CBC Val':>12}"
        f"{'HBG Val':>12}"
        f"{'CBP Val':>12}"
        f"{'CBC Inv':>12}"
        f"{'HBG Inv':>12}"
        f"{'CBP Inv':>12}"
        f"{'CBC U/C':>12}"
        f"{'HBG U/C':>12}"
        f"{'CBP U/C':>12}"
        f"{'Total Val':>12}"
        f"{'Total Inv':>12}"
        f"{'Total U/C':>12}"
    )
    print("  " + "-" * 160)
    for row in summary_rows:
        print(
            f"  {row['period']:<8}"
            f"{row['row_count']:>8,}"
            f"{row['cbc_value']:>12,}"
            f"{row['hbg_value']:>12,}"
            f"{row['cbp_value']:>12,}"
            f"{row['cbc_inventory']:>12,}"
            f"{row['hbg_inventory']:>12,}"
            f"{row['cbp_inventory']:>12,}"
            f"{row['cbc_uc']:>12.2f}"
            f"{row['hbg_uc']:>12.2f}"
            f"{row['cbp_uc']:>12.2f}"
            f"{row['total_value']:>12,}"
            f"{row['total_inventory']:>12,}"
            f"{row['total_uc']:>12.2f}"
        )


def check_inventory_no_value_rows():
    if not PICKLE_FILE.exists():
        print()
        print(f"Pickle file not found: {PICKLE_FILE}")
        return

    period = prompt_for_period()
    if period is None:
        return

    df = load_existing_vertical_pickle()
    period_df = df[df["period"].astype(str) == period].copy()
    if period_df.empty:
        print()
        print(f"No rows were found in the pickle for period {period}.")
        return

    period_df = period_df[
        period_df["Publisher"].fillna("").astype(str).str.strip() == "Chronicle"
    ].copy()
    period_df = period_df[
        (pd.to_numeric(period_df["Inventory"], errors="coerce").fillna(0) != 0)
        & (pd.to_numeric(period_df["Value"], errors="coerce").fillna(0) == 0)
    ].copy()

    if period_df.empty:
        print()
        print(f"No Chronicle rows with inventory and no value were found for period {period}.")
        return

    period_df["Inventory"] = pd.to_numeric(period_df["Inventory"], errors="coerce").fillna(0)
    period_df["Value"] = pd.to_numeric(period_df["Value"], errors="coerce").fillna(0)
    period_df = period_df.sort_values(["Inventory", "ISBN", "ORG"], ascending=[False, True, True])

    print()
    print(f"Chronicle rows with inventory but no value for period {period}:")
    print("  " + "-" * 100)
    print(f"  {'ISBN':<16}{'Title':<44}{'ORG':<8}{'Inventory':>16}{'Value':>16}")
    print("  " + "-" * 100)
    for _, row in period_df.iterrows():
        title = ""
        if pd.notna(row.get("Title")):
            title = str(row["Title"])[:44]
        print(
            f"  {str(row['ISBN']):<16}"
            f"{title:<44}"
            f"{str(row['ORG']).upper():<8}"
            f"{int(round(row['Inventory'])):>16,}"
            f"{int(round(row['Value'])):>16,}"
        )


def check_value_no_inventory_rows():
    if not PICKLE_FILE.exists():
        print()
        print(f"Pickle file not found: {PICKLE_FILE}")
        return

    period = prompt_for_period()
    if period is None:
        return

    df = load_existing_vertical_pickle()
    period_df = df[df["period"].astype(str) == period].copy()
    if period_df.empty:
        print()
        print(f"No rows were found in the pickle for period {period}.")
        return

    period_df = period_df[
        period_df["Publisher"].fillna("").astype(str).str.strip() == "Chronicle"
    ].copy()
    period_df = period_df[
        (pd.to_numeric(period_df["Value"], errors="coerce").fillna(0) != 0)
        & (pd.to_numeric(period_df["Inventory"], errors="coerce").fillna(0) == 0)
    ].copy()

    if period_df.empty:
        print()
        print(f"No Chronicle rows with value and no inventory were found for period {period}.")
        return

    period_df["Inventory"] = pd.to_numeric(period_df["Inventory"], errors="coerce").fillna(0)
    period_df["Value"] = pd.to_numeric(period_df["Value"], errors="coerce").fillna(0)
    period_df = period_df.sort_values(["Value", "ISBN", "ORG"], ascending=[False, True, True])

    print()
    print(f"Chronicle rows with value but no inventory for period {period}:")
    print("  " + "-" * 100)
    print(f"  {'ISBN':<16}{'Title':<44}{'ORG':<8}{'Value':>16}{'Inventory':>16}")
    print("  " + "-" * 100)
    for _, row in period_df.iterrows():
        title = ""
        if pd.notna(row.get("Title")):
            title = str(row["Title"])[:44]
        print(
            f"  {str(row['ISBN']):<16}"
            f"{title:<44}"
            f"{str(row['ORG']).upper():<8}"
            f"{int(round(row['Value'])):>16,}"
            f"{int(round(row['Inventory'])):>16,}"
        )


def check_negative_rows():
    if not PICKLE_FILE.exists():
        print()
        print(f"Pickle file not found: {PICKLE_FILE}")
        return

    period = prompt_for_period()
    if period is None:
        return

    df = load_existing_vertical_pickle()
    period_df = df[df["period"].astype(str) == period].copy()
    if period_df.empty:
        print()
        print(f"No rows were found in the pickle for period {period}.")
        return

    period_df = period_df[
        period_df["Publisher"].fillna("").astype(str).str.strip() == "Chronicle"
    ].copy()
    period_df["Inventory"] = pd.to_numeric(period_df["Inventory"], errors="coerce").fillna(0)
    period_df["Value"] = pd.to_numeric(period_df["Value"], errors="coerce").fillna(0)
    period_df = period_df[
        (period_df["Value"] < 0) | (period_df["Inventory"] < 0)
    ].copy()

    if period_df.empty:
        print()
        print(f"No Chronicle rows with negative value or inventory were found for period {period}.")
        return

    period_df["SortMagnitude"] = period_df[["Value", "Inventory"]].abs().max(axis=1)
    period_df = period_df.sort_values(
        ["SortMagnitude", "ISBN", "ORG"], ascending=[False, True, True]
    )

    print()
    print(f"Chronicle rows with negative value or inventory for period {period}:")
    print("  " + "-" * 100)
    print(f"  {'ISBN':<16}{'Title':<44}{'ORG':<8}{'Value':>16}{'Inventory':>16}")
    print("  " + "-" * 100)
    for _, row in period_df.iterrows():
        title = ""
        if pd.notna(row.get("Title")):
            title = str(row["Title"])[:44]
        print(
            f"  {str(row['ISBN']):<16}"
            f"{title:<44}"
            f"{str(row['ORG']).upper():<8}"
            f"{int(round(row['Value'])):>16,}"
            f"{int(round(row['Inventory'])):>16,}"
        )


def run_invobs_from_depot():
    period = prompt_for_period()
    if period is None:
        return

    print()
    print("Running Consolidate Inventory for the INVOBS...")
    print(f"  Period:           {period}")
    print(f"  Source pickle:    {PICKLE_FILE}")

    script_path = REPO_ROOT / "invobs_consolidated_inventory" / "main.py"
    try:
        import subprocess

        subprocess.run(
            [
                str(REPO_ROOT / "venv" / "Scripts" / "python"),
                str(script_path),
                "--period",
                str(period),
            ],
            check=True,
        )
        print("The INVOBS consolidated inventory file is now ready.")
    except subprocess.CalledProcessError:
        print("An error occurred while running invobs_consolidated_inventory/main.py.")


def coerce_vertical_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    typed_df = df.copy()
    typed_df["period"] = typed_df["period"].astype("string")
    typed_df["ISBN"] = typed_df["ISBN"].astype("string")
    typed_df["Title"] = typed_df["Title"].astype("string")
    typed_df["ORG"] = typed_df["ORG"].astype("string")
    typed_df["Publisher"] = typed_df["Publisher"].astype("string")
    typed_df["Value"] = pd.to_numeric(typed_df["Value"], errors="coerce").fillna(0)
    typed_df["Inventory"] = pd.to_numeric(typed_df["Inventory"], errors="coerce").fillna(0)
    return typed_df


def get_publisher_cache_path(today: datetime | None = None) -> Path:
    stamp = (today or datetime.now()).strftime("%Y%m%d")
    return PUBLISHER_CACHE_DIR / f"inventory_publishers_{stamp}.pkl"


@lru_cache(maxsize=1)
def load_publisher_lookup() -> pd.DataFrame:
    cache_path = get_publisher_cache_path()
    if cache_path.exists():
        cached_df = pd.read_pickle(cache_path)
        required_columns = {"ISBN", "Title", "Publisher"}
        if required_columns.issubset(set(cached_df.columns)):
            return cached_df[["ISBN", "Title", "Publisher"]].copy()

    engine = get_connection()
    publisher_df = fetch_data_from_db(engine, PUBLISHER_LOOKUP_SQL)
    if publisher_df.empty:
        return pd.DataFrame(columns=["ISBN", "Title", "Publisher"])

    required_columns = {"ISBN", "Title", "Publisher"}
    if not required_columns.issubset(set(publisher_df.columns)):
        raise ValueError("Publisher lookup query must return ISBN, Title, and Publisher columns.")

    publisher_df = publisher_df.copy()
    publisher_df["ISBN"] = publisher_df["ISBN"].map(normalize_isbn)
    publisher_df["Title"] = publisher_df["Title"].astype("string").str.strip()
    publisher_df["Publisher"] = publisher_df["Publisher"].astype("string").str.strip()
    publisher_df = publisher_df[publisher_df["ISBN"].notna()].copy()
    publisher_df = publisher_df.drop_duplicates(subset=["ISBN"], keep="first")
    publisher_df = publisher_df[["ISBN", "Title", "Publisher"]].reset_index(drop=True)

    cache_path.parent.mkdir(parents=True, exist_ok=True)
    publisher_df.to_pickle(cache_path)
    return publisher_df


def attach_publishers(df: pd.DataFrame) -> pd.DataFrame:
    publisher_lookup = load_publisher_lookup()
    if publisher_lookup.empty:
        enriched_df = df.copy()
        enriched_df["Title"] = pd.Series(pd.NA, index=enriched_df.index, dtype="string")
        enriched_df["Publisher"] = pd.Series(pd.NA, index=enriched_df.index, dtype="string")
        return enriched_df

    enriched_df = df.merge(publisher_lookup, how="left", on="ISBN")
    return enriched_df


def load_existing_vertical_pickle() -> pd.DataFrame:
    if not PICKLE_FILE.exists():
        empty_df = pd.DataFrame(
            columns=["period", "ISBN", "Title", "Publisher", "ORG", "Value", "Inventory"]
        )
        return coerce_vertical_dtypes(empty_df)

    existing_df = pd.read_pickle(PICKLE_FILE)
    expected_columns = ["period", "ISBN", "Title", "Publisher", "ORG", "Value", "Inventory"]
    if "Publisher" not in existing_df.columns or "Title" not in existing_df.columns:
        existing_df = attach_publishers(
            existing_df[["period", "ISBN", "ORG", "Value", "Inventory"]]
        )
        existing_df.insert(0, "period", existing_df.pop("period"))
    missing_columns = [column for column in expected_columns if column not in existing_df.columns]
    if missing_columns:
        missing_text = ", ".join(missing_columns)
        raise ValueError(
            f"Existing pickle is missing expected columns: {missing_text}. "
            f"File: {PICKLE_FILE}"
        )

    return coerce_vertical_dtypes(existing_df[expected_columns].copy())


def save_vertical_pickle(month_df: pd.DataFrame):
    existing_df = load_existing_vertical_pickle()
    month_df = coerce_vertical_dtypes(month_df)
    period = str(month_df["period"].iloc[0])
    existing_df = existing_df[existing_df["period"].astype(str) != period].copy()

    if existing_df.empty:
        combined_df = month_df.copy()
    else:
        combined_df = pd.concat([existing_df, month_df], ignore_index=True)

    combined_df = combined_df.sort_values(["period", "ISBN", "ORG"]).reset_index(drop=True)
    combined_df.to_pickle(PICKLE_FILE)
    return combined_df


def update_vertical_file():
    DESTINATION_FOLDER.mkdir(parents=True, exist_ok=True)

    source_file = choose_consolidated_file()
    if source_file is None:
        print("No file was selected.")
        return

    if not source_file.exists():
        print(f"Selected file was not found: {source_file}")
        return

    period = infer_period_from_filename(source_file)
    if period is None:
        period = prompt_for_period()
        if period is None:
            return

    process_consolidated_file(source_file, period)


def show_last_10_periods_in_sharepoint_file():
    if not PICKLE_FILE.exists():
        print()
        print(f"Pickle file not found: {PICKLE_FILE}")
        return

    df = load_existing_vertical_pickle()
    periods = sorted(df["period"].astype(str).unique().tolist(), reverse=True)[:10]

    print()
    print("Last 10 periods in the SharePoint file:")
    if not periods:
        print("  No periods were found.")
        return

    for period in periods:
        print(f"  {period}")


def refresh_verticalized_coninv_to_sharepoint():
    if not PICKLE_FILE.exists():
        print()
        print(f"Source pickle not found: {PICKLE_FILE}")
        return

    username = getpass.getuser().strip().lower()
    sharepoint_destination_folder = KNOWN_SHAREPOINT_FOLDERS.get(username)
    if sharepoint_destination_folder is None:
        sharepoint_destination_folder = Path(r"C:\Users") / username / SHAREPOINT_FOLDER_SUFFIX

    if not sharepoint_destination_folder.exists():
        desktop_pickle_file = Path.home() / "Desktop" / "Consolidated_Inventory.pkl"
        shutil.copy2(PICKLE_FILE, desktop_pickle_file)
        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            messagebox.showwarning(
                "CB Fabric Folder Not Found",
                "I don't see your link to the 'CB Fabric' folder.\n\n"
                f"I saved a copy here:\n{desktop_pickle_file}\n\n"
                "You'll have to upload it yourself.",
            )
        finally:
            root.destroy()

        print()
        print("SharePoint destination folder was not found.")
        print(f"Saved a manual-upload copy to: {desktop_pickle_file}")
        return

    sharepoint_pickle_file = sharepoint_destination_folder / "Consolidated_Inventory.pkl"
    shutil.copy2(PICKLE_FILE, sharepoint_pickle_file)

    print()
    print("Refreshed verticalized ConInv to SharePoint:")
    print(f"  Source:       {PICKLE_FILE}")
    print(f"  Destination:  {sharepoint_pickle_file}")


def run_verticalize_menu():
    while True:
        print("\nVerticalize ConInv in Bash Depot and Add to pickle file")
        print()
        print("    1. Verticalize file from Bash Depot and append to pickle file")
        print("    2. See Last 10 Periods In SharePoint file")
        print("    3. Back")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice == "1":
            update_vertical_file()
            continue

        if choice == "2":
            show_last_10_periods_in_sharepoint_file()
            continue

        if choice in {"3", "back", "b", "return", "menu"}:
            return

        print("Invalid choice. Please select a valid option.")


def process_consolidated_file(source_file: Path, period: str):
    excel_file = pd.ExcelFile(source_file)
    selected_sheet = choose_data_sheet(excel_file)
    if selected_sheet is None:
        raise ValueError("Unable to determine the data worksheet.")

    print()
    print("Updating vertical inventory file...")
    print(f"  Source workbook:  {source_file}")
    print(f"  Data worksheet:   {selected_sheet}")

    raw_df = pd.read_excel(source_file, sheet_name=selected_sheet, header=None)
    vertical_df = build_vertical_dataframe(raw_df)
    metrics_df = build_metrics_dataframe(vertical_df)
    output_df = vertical_df[["ISBN", "ORG", "Value", "Inventory"]].copy()
    output_df.insert(0, "period", period)
    output_df = attach_publishers(output_df)
    output_df = output_df[["period", "ISBN", "Title", "Publisher", "ORG", "Value", "Inventory"]]
    combined_df = save_vertical_pickle(output_df)

    print()
    print(f"Updated pickle file: {PICKLE_FILE}")
    print(f"Rows in current period: {len(output_df):,}")
    print(f"Rows in all periods: {len(combined_df):,}")
    print()
    print("Metrics:")
    for _, row in metrics_df.iterrows():
        metric_name = row["Metric"]
        metric_value = row["Value"]
        print(f"  {metric_name}: {int(round(metric_value)):,}")
    print_publisher_org_summary(output_df)


def run_menu():
    while True:
        print("\nConsolidated Inventory Manager (ConInv)")
        print()
        print("    1. Add New ConInv Report to Bash Depot")
        print("    2. Verticalize ConInv from Bash Depot into Pickle")
        print("    3. Copy ConInv Pickle to SharePoint")
        print("    4. Summarize/View Top N Rows of ConInv (Pickle file)")
        print("    5. Summarize a Specific Verticalized Month (Period) of the ConInv (Pickle file)")
        print("    6. Summarize Last N Months (CB Only) of ConInv (Pickle file)")
        print("    7. Summarize Rows of ConInv (Pickle file) with Inventory but No Value (CB Only) e.g. CDUs")
        print("    8. Summarize Rows of ConInv (Pickle file) with Values but No Inventory (CB Only)")
        print("    9. Summarize Chronicle Rows of ConInv with Negative Values or Negative Inventory")
        print("    10. Consolidate Inventory for the INVOBS")
        print("    99. Exit")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice == "1":
            run_depot_menu()
            continue

        if choice == "2":
            run_verticalize_menu()
            continue

        if choice == "3":
            refresh_verticalized_coninv_to_sharepoint()
            continue

        if choice == "4":
            preview_verticalized_coninv()
            continue

        if choice == "5":
            check_specific_period()
            continue

        if choice == "6":
            check_last_n_months()
            continue

        if choice == "7":
            check_inventory_no_value_rows()
            continue

        if choice == "8":
            check_value_no_inventory_rows()
            continue

        if choice == "9":
            check_negative_rows()
            continue

        if choice == "10":
            run_invobs_from_depot()
            continue

        if choice in {"99", "exit", "quit", "q", "back", "b"}:
            return

        print("Invalid choice. Please select a valid option.")


def main():
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--source-file")
    parser.add_argument("--period")
    args, _ = parser.parse_known_args()

    if args.source_file:
        source_file = Path(args.source_file)
        period = args.period or infer_period_from_filename(source_file)
        if period is None:
            raise ValueError("A period is required when it cannot be inferred from the file name.")
        process_consolidated_file(source_file, period)
        return

    run_menu()


if __name__ == "__main__":
    main()
