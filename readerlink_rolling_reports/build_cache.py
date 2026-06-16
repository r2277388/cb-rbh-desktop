from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
SHARED_BASE_DIR = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Readerlink")
CACHE_DIR = SHARED_BASE_DIR / "cache"

SALES_SOURCE_DIRS = [
    Path(r"F:\ANALYSIS\Finance\DataWarehouse\Weekly reports\2026\Readerlink"),
    Path(r"F:\ANALYSIS\Finance\DataWarehouse\Weekly reports\2025\Readerlink"),
    Path(r"F:\ANALYSIS\Finance\DataWarehouse\Weekly reports\Data Archive\2024\Readerlink"),
]
INVENTORY_SOURCE_DIR = (
    Path(r"F:\ANALYSIS\Finance\DataWarehouse\Weekly reports\2026\Readerlink\Inventory")
)
HISTORICAL_ROLLING_FILE = BASE_DIR / "Week 23 - 2026 Rolling Readerlink (060626).xlsx"

SALES_COLUMNS = [
    "MASTER CHAIN",
    "EAN",
    "CY OB UNITS",
    "CY RET UNITS",
    "CY NET UNITS",
    "CY POS UNITS",
]
OPTIONAL_SALES_COLUMNS = ["TITLE", "AUTHOR NAME", "ONSALE DATE", "MSRP"]
INVENTORY_COLUMNS = ["ITEM", "RDS DC INVENTORY OH", "RDS OPEN PO QUANTITY"]
OPTIONAL_INVENTORY_COLUMNS = ["TITLE", "AUTHOR NAME", "MSRP", "ONSALE DATE"]

FILENAME_DATE_RE = re.compile(r"\b(\d{2})\s+(\d{4})\s+(\d{6})\s+Readerlink\b", re.IGNORECASE)
ROLLING_SHEETS_TO_SKIP = {"pgrp_key"}
PRE_MODERN_HISTORY_CUTOFF = pd.Timestamp("2024-01-01")
EARLIEST_HISTORICAL_WEEK = pd.Timestamp("2014-01-01")


def normalized_header(value) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(value).upper())


SALES_HEADER_ALIASES = {
    "master_chain": {"MASTERCHAIN", "CHAIN"},
    "isbn": {"EAN", "ITEMNUMBER"},
    "cy_ob_units": {"CYOBUNITS", "OUTBOUNDUNITS"},
    "cy_ret_units": {"CYRETUNITS", "RETURNSUNITS"},
    "cy_net_units": {"CYNETUNITS", "NETUNITS"},
    "cy_pos_units": {"CYPOSUNITS", "POSUNITS"},
    "title": {"TITLE"},
    "author_name": {"AUTHORNAME", "AUTHOR"},
    "onsale_date": {"ONSALEDATE"},
    "msrp": {"MSRP", "COVERPRICE"},
}


@dataclass(frozen=True)
class SourceFile:
    path: Path
    week_number: int
    year: int
    week_end: pd.Timestamp


def normalize_isbn(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().replace("-", "").replace(" ", "")
    if text.endswith(".0"):
        text = text[:-2]
    digits = "".join(ch for ch in text if ch.isdigit())
    if not digits:
        return ""
    return digits.zfill(13)[-13:]


def parse_readerlink_filename(path: Path) -> SourceFile | None:
    match = FILENAME_DATE_RE.search(path.name)
    if not match:
        return None
    week_number = int(match.group(1))
    year = int(match.group(2))
    date_token = match.group(3)
    week_end = pd.Timestamp(datetime.strptime(date_token, "%m%d%y"))
    return SourceFile(path=path, week_number=week_number, year=year, week_end=week_end)


def find_sales_files() -> list[SourceFile]:
    files: list[SourceFile] = []
    for folder in SALES_SOURCE_DIRS:
        if not folder.exists():
            print(f"Warning: missing sales folder: {folder}")
            continue
        for path in folder.glob("*.xlsx"):
            if path.name.startswith("~$"):
                continue
            if "inventory" in path.name.lower():
                continue
            parsed = parse_readerlink_filename(path)
            if parsed is not None:
                files.append(parsed)
    return sorted(files, key=lambda item: (item.week_end, item.path.name))


def find_inventory_files() -> list[SourceFile]:
    if not INVENTORY_SOURCE_DIR.exists():
        print(f"Warning: missing inventory folder: {INVENTORY_SOURCE_DIR}")
        return []
    files: list[SourceFile] = []
    for path in INVENTORY_SOURCE_DIR.glob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        if "inventory" not in path.name.lower():
            continue
        parsed = parse_readerlink_filename(path)
        if parsed is not None:
            files.append(parsed)
    return sorted(files, key=lambda item: (item.week_end, item.path.name))


def detect_sales_sheet_and_header(path: Path) -> tuple[str, int, dict[str, str]]:
    xls = pd.ExcelFile(path)
    for sheet_name in xls.sheet_names:
        preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=12)
        for row_index in range(len(preview.index)):
            raw_headers = preview.iloc[row_index].tolist()
            alias_to_column: dict[str, str] = {}
            for raw_header in raw_headers:
                normalized = normalized_header(raw_header)
                if not normalized or normalized == "NAN":
                    continue
                for canonical, aliases in SALES_HEADER_ALIASES.items():
                    if normalized in aliases:
                        alias_to_column[canonical] = raw_header
            required = {
                "master_chain",
                "isbn",
                "cy_ob_units",
                "cy_ret_units",
                "cy_net_units",
                "cy_pos_units",
            }
            if required.issubset(alias_to_column):
                return sheet_name, row_index, alias_to_column
    raise ValueError(f"Could not find Readerlink sales columns in any sheet of {path}")


def read_sales_file(source: SourceFile) -> pd.DataFrame:
    sheet_name, header_row, alias_to_column = detect_sales_sheet_and_header(source.path)
    df = pd.read_excel(source.path, sheet_name=sheet_name, header=header_row, dtype=object)
    rename_map = {
        source_column: canonical
        for canonical, source_column in alias_to_column.items()
        if source_column in df.columns
    }
    keep_columns = list(rename_map)
    df = df[keep_columns].copy().rename(columns=rename_map)
    df["isbn"] = df["isbn"].map(normalize_isbn)
    df = df[df["isbn"] != ""]
    df["master_chain"] = df["master_chain"].astype("string").str.strip().fillna("")
    for column in ["cy_ob_units", "cy_ret_units", "cy_net_units", "cy_pos_units"]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)
    if "msrp" in df.columns:
        df["msrp"] = pd.to_numeric(df["msrp"], errors="coerce")

    df["week_end"] = source.week_end
    df["readerlink_week"] = source.week_number
    df["readerlink_year"] = source.year
    df["source_file"] = source.path.name

    group_columns = [
        "week_end",
        "readerlink_week",
        "readerlink_year",
        "master_chain",
        "isbn",
        "source_file",
    ]
    metadata_columns = [column for column in ["title", "author_name", "onsale_date", "msrp"] if column in df.columns]
    numeric_columns = ["cy_ob_units", "cy_ret_units", "cy_net_units", "cy_pos_units"]
    grouped = df.groupby(group_columns, as_index=False, dropna=False)[numeric_columns].sum()
    if metadata_columns:
        metadata = df[group_columns + metadata_columns].drop_duplicates(subset=group_columns, keep="first")
        grouped = grouped.merge(metadata, on=group_columns, how="left")
    return grouped


def build_sales_cache() -> pd.DataFrame:
    files = find_sales_files()
    if not files:
        raise FileNotFoundError("No Readerlink sales source files found.")

    frames = []
    for index, source in enumerate(files, start=1):
        print(f"Reading sales {index:>3}/{len(files)}: {source.path.name}")
        frames.append(read_sales_file(source))

    sales = pd.concat(frames, ignore_index=True)
    sales = sales.sort_values(["week_end", "master_chain", "isbn"]).reset_index(drop=True)
    return sales


def read_inventory_file(source: SourceFile) -> pd.DataFrame:
    df = pd.read_excel(source.path, sheet_name="Export", dtype={"ITEM": str})
    missing = [column for column in INVENTORY_COLUMNS if column not in df.columns]
    if missing:
        raise ValueError(f"{source.path} is missing required inventory columns: {missing}")

    keep_columns = INVENTORY_COLUMNS + [
        column for column in OPTIONAL_INVENTORY_COLUMNS if column in df.columns
    ]
    df = df[keep_columns].copy()
    df = df.rename(
        columns={
            "ITEM": "isbn",
            "RDS DC INVENTORY OH": "rds_dc_inventory_oh",
            "RDS OPEN PO QUANTITY": "rds_open_po_quantity",
            "TITLE": "title",
            "AUTHOR NAME": "author_name",
            "MSRP": "msrp",
            "ONSALE DATE": "onsale_date",
        }
    )
    df["isbn"] = df["isbn"].map(normalize_isbn)
    df = df[df["isbn"] != ""]
    for column in ["rds_dc_inventory_oh", "rds_open_po_quantity"]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)
    if "msrp" in df.columns:
        df["msrp"] = pd.to_numeric(df["msrp"], errors="coerce")

    df["week_end"] = source.week_end
    df["readerlink_week"] = source.week_number
    df["readerlink_year"] = source.year
    df["source_file"] = source.path.name

    group_columns = ["week_end", "readerlink_week", "readerlink_year", "isbn", "source_file"]
    numeric_columns = ["rds_dc_inventory_oh", "rds_open_po_quantity"]
    metadata_columns = [column for column in ["title", "author_name", "onsale_date", "msrp"] if column in df.columns]
    grouped = df.groupby(group_columns, as_index=False, dropna=False)[numeric_columns].sum()
    if metadata_columns:
        metadata = df[group_columns + metadata_columns].drop_duplicates(subset=group_columns, keep="first")
        grouped = grouped.merge(metadata, on=group_columns, how="left")
    return grouped.sort_values(["week_end", "isbn"]).reset_index(drop=True)


def build_latest_inventory_cache() -> pd.DataFrame:
    files = find_inventory_files()
    if not files:
        raise FileNotFoundError("No Readerlink inventory source files found.")
    latest = files[-1]
    print(f"Reading latest inventory: {latest.path.name}")
    return read_inventory_file(latest)


def rolling_sheet_to_master_chain(sheet_name: str) -> str:
    if sheet_name == "Summary - All Accounts":
        return "ALL ACCOUNTS"
    return sheet_name.strip()


def build_historical_pos_cache() -> pd.DataFrame:
    if not HISTORICAL_ROLLING_FILE.exists():
        raise FileNotFoundError(HISTORICAL_ROLLING_FILE)

    xls = pd.ExcelFile(HISTORICAL_ROLLING_FILE)
    frames = []
    for sheet_name in xls.sheet_names:
        if sheet_name in ROLLING_SHEETS_TO_SKIP:
            continue
        print(f"Reading historical rolling sheet: {sheet_name}")
        df = pd.read_excel(xls, sheet_name=sheet_name, header=2)
        if "ISBN 13" not in df.columns:
            continue

        date_columns = [
            column
            for column in df.columns
            if isinstance(column, datetime)
            and EARLIEST_HISTORICAL_WEEK <= pd.Timestamp(column) < PRE_MODERN_HISTORY_CUTOFF
            and pd.Timestamp(column).weekday() == 5
        ]
        if not date_columns:
            continue

        id_columns = [column for column in ["ISBN 13", "TITLE", "PRICE", "SHIP"] if column in df.columns]
        history = df[id_columns + date_columns].copy()
        history = history.rename(
            columns={
                "ISBN 13": "isbn",
                "TITLE": "title",
                "PRICE": "price",
                "SHIP": "ship_date",
            }
        )
        history["isbn"] = history["isbn"].map(normalize_isbn)
        history = history[history["isbn"] != ""]
        melted = history.melt(
            id_vars=[column for column in ["isbn", "title", "price", "ship_date"] if column in history.columns],
            value_vars=date_columns,
            var_name="week_end",
            value_name="cy_pos_units",
        )
        melted["cy_pos_units"] = pd.to_numeric(melted["cy_pos_units"], errors="coerce").fillna(0)
        melted = melted[melted["cy_pos_units"] != 0]
        melted["week_end"] = pd.to_datetime(melted["week_end"])
        melted["master_chain"] = rolling_sheet_to_master_chain(sheet_name)
        melted["source_file"] = HISTORICAL_ROLLING_FILE.name
        melted["source_sheet"] = sheet_name
        frames.append(melted)

    if not frames:
        return pd.DataFrame(
            columns=[
                "week_end",
                "master_chain",
                "isbn",
                "cy_pos_units",
                "source_file",
                "source_sheet",
            ]
        )

    history = pd.concat(frames, ignore_index=True)
    history = history.sort_values(["week_end", "master_chain", "isbn"]).reset_index(drop=True)
    return history


def save_cache(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    output = df.copy()
    for column in ["week_end", "onsale_date", "ship_date"]:
        if column in output.columns:
            output[column] = pd.to_datetime(output[column], errors="coerce")
    for column in [
        "master_chain",
        "isbn",
        "source_file",
        "source_sheet",
        "title",
        "author_name",
    ]:
        if column in output.columns:
            output[column] = output[column].astype("string")
    df = output
    df.to_parquet(path, index=False)
    print(f"Saved {len(df):,} rows: {path}")


def build_combined_pos_history(sales: pd.DataFrame, history: pd.DataFrame) -> pd.DataFrame:
    modern_columns = [
        "week_end",
        "master_chain",
        "isbn",
        "cy_pos_units",
        "source_file",
    ]
    modern = sales[modern_columns + [column for column in ["title", "msrp"] if column in sales.columns]].copy()
    modern["source_sheet"] = "Export/Data"
    combined = pd.concat([history, modern], ignore_index=True, sort=False)
    return combined.sort_values(["week_end", "master_chain", "isbn"]).reset_index(drop=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build Readerlink rolling report source caches.")
    parser.add_argument(
        "--skip-history",
        action="store_true",
        help="Skip pre-2024 history extraction from the old rolling workbook.",
    )
    args = parser.parse_args()

    sales = build_sales_cache()
    save_cache(sales, CACHE_DIR / "readerlink_weekly_sales.parquet")

    inventory = build_latest_inventory_cache()
    save_cache(inventory, CACHE_DIR / "readerlink_latest_inventory.parquet")

    if not args.skip_history:
        history = build_historical_pos_cache()
        save_cache(history, CACHE_DIR / "readerlink_pre2024_pos_history.parquet")
        combined_pos = build_combined_pos_history(sales, history)
        save_cache(combined_pos, CACHE_DIR / "readerlink_pos_history.parquet")

    print("Readerlink cache build complete.")


if __name__ == "__main__":
    main()
