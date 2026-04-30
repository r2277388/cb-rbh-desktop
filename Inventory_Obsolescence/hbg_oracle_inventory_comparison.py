from __future__ import annotations

import argparse
import re
from datetime import datetime
from pathlib import Path
from tkinter import Tk, filedialog

import pandas as pd

try:
    from paths import process_paths
    from shared.db import fetch_data_from_db, get_connection
except ImportError:
    import sys

    sys.path.append(str(Path(__file__).resolve().parents[1]))
    from paths import process_paths
    from shared.db import fetch_data_from_db, get_connection


MOON_REQUIRED_COLUMNS = ["ISBN", "Ending Balance"]
DEFAULT_MOON_HEADERS = [
    "ISBN",
    "Total Received",
    "Ending Balance",
    "Useable Balance",
    "Indy Other",
    "Chain Stock",
    "WMS Good",
    "WMS QA Hold",
    "WMS Mixed",
    "As Is",
    "Indy Allocated Qty",
    "Soft Allocated",
    "Frozen",
    "Units Backorded",
    "Lines Backordered",
    "Reserve Qty",
    "Consingment",
    "CC Allocated",
    "CC Receipts",
    "Stickered Qty",
    "Hold Qty",
    "Credit Hold Qty",
    "Carton Qty",
    "USA Price",
    "Canada Price",
    "Pub Status",
    "Indy Non Default Non Stickered",
    "NQ",
    "AI",
    "JX",
    "BY",
    "EC",
    "EF",
    "EG",
    "PR",
    "RP",
    "SS",
    "WO",
    "BL",
    "CP",
    "CT",
    "LV",
    "MD",
    "MK",
    "MS",
    "MT",
    "NF",
    "PY",
    "S",
    "SG",
    "ST",
    "XR",
    "Z",
]
CONINV_PICKLE_FILE = process_paths.CONSOLIDATED_INVENTORY_VERTICALIZATION_FOLDER / "ConsolidatedInventory.pkl"

CHRONICLE_TITLE_SQL = """
SELECT
    i.ITEM_TITLE AS ITEM,
    i.SHORT_TITLE Title,
    i.PUBLISHING_GROUP pgrp,
    i.AMORTIZATION_DATE
FROM ebs.item i
WHERE i.PUBLISHER_CODE = 'Chronicle'
    AND i.ISBN IS NOT NULL
    AND i.SHORT_TITLE IS NOT NULL
    AND i.PRODUCT_TYPE IN ('BK', 'FT', 'CP', 'RP')
"""


def normalize_isbn(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    digits = re.sub(r"\D", "", str(value).strip())
    if not digits:
        return ""
    if len(digits) < 13:
        return digits.zfill(13)
    if len(digits) > 13 and digits.startswith("0"):
        return digits[-13:]
    return digits[:13]


def parse_number(value: object) -> float:
    if value is None or pd.isna(value):
        return 0.0
    text = str(value).replace(",", "").strip()
    if text in {"", "-"}:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def prompt_for_period() -> str | None:
    while True:
        value = input("Enter the ConInv period to compare, e.g. 202603, or type 'back': ").strip()
        if value.lower() in {"back", "b", "return", "menu"}:
            return None
        if re.fullmatch(r"\d{6}", value) and 1 <= int(value[4:]) <= 12:
            return value
        print("Invalid period. Use yyyymm, for example 202603.")


def choose_moon_file() -> Path | None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = filedialog.askopenfilename(
            title="Choose Molly's HBG Moon inventory .ok file",
            initialdir=str(process_paths.INVENTORY_OBSOLESCENCE_FOLDER),
            filetypes=[("Moon inventory files", "*.ok"), ("All files", "*.*")],
        )
    finally:
        root.destroy()
    return Path(selected) if selected else None


def read_moon_headers() -> list[str]:
    workbook = process_paths.HBG_MOON_HEADER_WORKBOOK
    if workbook.exists():
        header_df = pd.read_excel(workbook, sheet_name=0, header=None, nrows=2, dtype=object)
        headers = [str(value).strip() for value in header_df.iloc[1].tolist()]
    else:
        headers = DEFAULT_MOON_HEADERS
    if any(column not in headers for column in MOON_REQUIRED_COLUMNS):
        raise ValueError(
            f"Moon header workbook must include columns: {', '.join(MOON_REQUIRED_COLUMNS)}"
        )
    return headers


def load_moon_inventory(moon_file: Path) -> pd.DataFrame:
    if not moon_file.exists():
        raise FileNotFoundError(f"Moon inventory file not found: {moon_file}")
    headers = read_moon_headers()
    usecols = [headers.index(column) for column in MOON_REQUIRED_COLUMNS]
    raw = pd.read_csv(
        moon_file,
        sep="^",
        header=None,
        usecols=usecols,
        dtype=object,
        encoding="unicode_escape",
        engine="python",
    )
    raw.columns = MOON_REQUIRED_COLUMNS
    moon = pd.DataFrame(
        {
            "ISBN": raw["ISBN"].map(normalize_isbn),
            "Moon Ending Balance": raw["Ending Balance"].map(parse_number),
        }
    )
    moon = moon[moon["ISBN"].ne("")].copy()
    moon = moon.groupby("ISBN", as_index=False)["Moon Ending Balance"].sum()
    return moon


def load_hbg_coninv(period: str) -> pd.DataFrame:
    if not CONINV_PICKLE_FILE.exists():
        raise FileNotFoundError(f"ConInv pickle not found: {CONINV_PICKLE_FILE}")
    coninv = pd.read_pickle(CONINV_PICKLE_FILE)
    required = {"period", "ISBN", "ORG", "Inventory"}
    missing = sorted(required - set(coninv.columns))
    if missing:
        raise ValueError(f"ConInv pickle is missing columns: {', '.join(missing)}")
    coninv_columns = ["period", "ISBN", "ORG", "Inventory"]
    coninv_columns.extend(column for column in ["Value", "Title", "Publisher", "PGRP"] if column in coninv.columns)
    coninv = pd.DataFrame({column: list(coninv[column]) for column in coninv_columns})
    hbg = coninv[
        coninv["period"].astype(str).eq(str(period))
        & coninv["ORG"].astype(str).str.strip().str.upper().eq("HBG")
    ].copy()
    if hbg.empty:
        raise ValueError(f"No HBG ConInv rows found for period {period}.")
    hbg["ISBN"] = hbg["ISBN"].map(normalize_isbn)
    hbg["Oracle HBG Inventory"] = hbg["Inventory"].map(parse_number)
    if "Value" in hbg.columns:
        hbg["HBG Value"] = hbg["Value"].map(parse_number)
    else:
        hbg["HBG Value"] = 0.0
    keep_cols = ["ISBN", "Oracle HBG Inventory"]
    keep_cols.append("HBG Value")
    for column in ["Title", "Publisher", "PGRP"]:
        if column in hbg.columns:
            keep_cols.append(column)
    hbg = hbg[hbg["ISBN"].ne("")]
    aggregate = {"Oracle HBG Inventory": "sum", "HBG Value": "sum"}
    for column in ["Title", "Publisher", "PGRP"]:
        if column in hbg.columns:
            aggregate[column] = "first"
    hbg = hbg[keep_cols].groupby("ISBN", as_index=False).agg(aggregate)
    hbg["HBG Unit Value"] = 0.0
    hbg.loc[hbg["Oracle HBG Inventory"].ne(0), "HBG Unit Value"] = (
        hbg.loc[hbg["Oracle HBG Inventory"].ne(0), "HBG Value"]
        / hbg.loc[hbg["Oracle HBG Inventory"].ne(0), "Oracle HBG Inventory"]
    )
    return hbg


def load_chronicle_titles() -> pd.DataFrame:
    print("Refreshing Chronicle title filter from SQL...")
    engine = get_connection()
    titles = fetch_data_from_db(engine, CHRONICLE_TITLE_SQL)
    required = {"ITEM", "Title", "pgrp", "AMORTIZATION_DATE"}
    missing = sorted(required - set(titles.columns))
    if missing:
        raise ValueError(f"Chronicle title SQL is missing columns: {', '.join(missing)}")
    titles = titles.rename(columns={"ITEM": "ISBN", "pgrp": "PGRP", "AMORTIZATION_DATE": "PubDate"})
    titles["ISBN"] = titles["ISBN"].map(normalize_isbn)
    titles = titles[titles["ISBN"].ne("")].copy()
    titles["Publisher"] = "Chronicle"
    return titles.drop_duplicates("ISBN")[["ISBN", "Title", "Publisher", "PGRP", "PubDate"]]


def build_comparison(moon: pd.DataFrame, hbg: pd.DataFrame, chronicle_titles: pd.DataFrame) -> pd.DataFrame:
    comparison = hbg.merge(moon, on="ISBN", how="outer", indicator=True)
    comparison["Oracle HBG Inventory"] = comparison["Oracle HBG Inventory"].fillna(0)
    comparison["Moon Ending Balance"] = comparison["Moon Ending Balance"].fillna(0)
    comparison["Variance"] = comparison["Moon Ending Balance"] - comparison["Oracle HBG Inventory"]
    comparison["HBG Unit Value"] = comparison["HBG Unit Value"].fillna(0)
    comparison["Dollar Variance"] = comparison["Variance"] * comparison["HBG Unit Value"]
    denominator = comparison[["Oracle HBG Inventory", "Moon Ending Balance"]].abs().max(axis=1)
    comparison["Variance %"] = 0.0
    comparison.loc[denominator.ne(0), "Variance %"] = (
        comparison.loc[denominator.ne(0), "Variance"].abs() / denominator[denominator.ne(0)]
    )
    comparison["Status"] = "Match"
    comparison.loc[comparison["_merge"].eq("right_only"), "Status"] = "Only in Moon"
    comparison.loc[comparison["_merge"].eq("left_only"), "Status"] = "Only in Oracle HBG"
    comparison.loc[comparison["Variance"].ne(0) & comparison["Status"].eq("Match"), "Status"] = "Variance"
    comparison = comparison.drop(columns=["Title", "Publisher", "PGRP"], errors="ignore")
    comparison = comparison.merge(chronicle_titles, on="ISBN", how="inner")
    comparison = comparison[comparison["Variance"].ne(0)].copy()
    comparison["Meaningful Variance"] = (
        comparison["Variance"].abs().gt(100)
        & comparison["Variance %"].gt(0.05)
    )
    preferred = [
        "Status",
        "ISBN",
        "Title",
        "Publisher",
        "PGRP",
        "PubDate",
        "Oracle HBG Inventory",
        "Moon Ending Balance",
        "Variance",
        "HBG Unit Value",
        "Dollar Variance",
        "Variance %",
        "Meaningful Variance",
    ]
    for column in preferred:
        if column not in comparison.columns:
            comparison[column] = pd.NA
    return comparison[preferred].sort_values("Dollar Variance", ascending=False).reset_index(drop=True)


def output_path_for(period: str) -> Path:
    folder = process_paths.HBG_ORACLE_COMPARISON_OUTPUT_FOLDER
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f"{period} - HBG vs Oracle Inventory Comparison.xlsx"


def save_report(comparison: pd.DataFrame, period: str, moon_file: Path) -> Path:
    output = output_path_for(period)
    summary = pd.DataFrame(
        [
            {"Metric": "Period", "Value": period},
            {"Metric": "Moon file", "Value": str(moon_file)},
            {"Metric": "Moon file modified", "Value": datetime.fromtimestamp(moon_file.stat().st_mtime)},
            {"Metric": "ConInv pickle", "Value": str(CONINV_PICKLE_FILE)},
            {"Metric": "Title filter", "Value": "Chronicle SQL title list"},
            {"Metric": "Rows", "Value": len(comparison)},
            {"Metric": "Variance rows", "Value": int(comparison["Variance"].ne(0).sum())},
            {"Metric": "Meaningful variance rows", "Value": int(comparison["Meaningful Variance"].sum())},
            {"Metric": "Total Oracle HBG Inventory", "Value": float(comparison["Oracle HBG Inventory"].sum())},
            {"Metric": "Total Moon Ending Balance", "Value": float(comparison["Moon Ending Balance"].sum())},
            {"Metric": "Total Variance", "Value": float(comparison["Variance"].sum())},
            {"Metric": "Total Dollar Variance", "Value": float(comparison["Dollar Variance"].sum())},
        ]
    )
    try:
        writer_context = pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy")
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output = output.with_name(f"{output.stem} - {timestamp}{output.suffix}")
        writer_context = pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy")

    with writer_context as writer:
        comparison.to_excel(writer, sheet_name="Comparison", index=False)
        comparison[comparison["Meaningful Variance"]].to_excel(writer, sheet_name="Meaningful Variances", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        percent_format = writer.book.add_format({"num_format": "0.0%"})
        money_format = writer.book.add_format({"num_format": '$#,##0.00;[Red]($#,##0.00);-'})
        unit_format = writer.book.add_format({"num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'})
        for sheet_name, worksheet in writer.sheets.items():
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, worksheet.dim_rowmax or 0, worksheet.dim_colmax or 0)
            worksheet.set_column("A:A", 18)
            worksheet.set_column("B:B", 15)
            worksheet.set_column("C:C", 42)
            worksheet.set_column("D:E", 16)
            worksheet.set_column("F:F", 12)
            worksheet.set_column("G:I", 18, unit_format)
            worksheet.set_column("J:K", 16, money_format)
            worksheet.set_column("L:L", 12, percent_format)
    print(f"Saved HBG vs Oracle inventory comparison: {output}")
    return output


def run_comparison(moon_file: Path | None = None, period: str | None = None) -> Path | None:
    if period is None:
        period = prompt_for_period()
        if period is None:
            return None
    if moon_file is None:
        moon_file = choose_moon_file()
        if moon_file is None:
            print("No Moon inventory file selected.")
            return None
    print()
    print("HBG vs Oracle Inventory Comparison")
    print(f"  Period:      {period}")
    print(f"  Moon file:   {moon_file}")
    print(f"  Header file: {process_paths.HBG_MOON_HEADER_WORKBOOK}")
    print(f"  ConInv file: {CONINV_PICKLE_FILE}")
    moon = load_moon_inventory(moon_file)
    hbg = load_hbg_coninv(period)
    chronicle_titles = load_chronicle_titles()
    comparison = build_comparison(moon, hbg, chronicle_titles)
    output = save_report(comparison, period, moon_file)
    print(f"  Moon rows:       {len(moon):,}")
    print(f"  Oracle HBG rows: {len(hbg):,}")
    print(f"  Chronicle titles:{len(chronicle_titles):,}")
    print(f"  Report rows:     {len(comparison):,}")
    print(f"  Variance rows:   {int(comparison['Variance'].ne(0).sum()):,}")
    print(f"  Meaningful rows: {int(comparison['Meaningful Variance'].sum()):,}")
    return output


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compare HBG Moon inventory to Oracle HBG ConInv inventory.")
    parser.add_argument("--moon-file", type=Path, help="Path to Molly's .ok Moon inventory file.")
    parser.add_argument("--period", help="ConInv period to compare, e.g. 202603.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    run_comparison(args.moon_file, args.period)


if __name__ == "__main__":
    main()
