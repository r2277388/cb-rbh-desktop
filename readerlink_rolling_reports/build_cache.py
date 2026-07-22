from __future__ import annotations

import argparse
import re
import sys
import warnings
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
REPO_ROOT = BASE_DIR.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from shared.bookscan_calendar import bookscan_week

warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
    module="openpyxl.styles.stylesheet",
)

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
STORE_INVENTORY_SOURCE_DIR = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Weekly reports\2026\Readerlink\OH_Store_TitlePerformanceReport"
)
HISTORICAL_ROLLING_FILE = BASE_DIR / "Week 23 - 2026 Rolling Readerlink (060626).xlsx"
WEEKLY_SALES_CACHE = CACHE_DIR / "readerlink_weekly_sales.parquet"
LATEST_INVENTORY_CACHE = CACHE_DIR / "readerlink_latest_inventory.parquet"
INVENTORY_HISTORY_CACHE = CACHE_DIR / "readerlink_inventory_history.parquet"
STORE_INVENTORY_CACHE = CACHE_DIR / "readerlink_store_inventory_history.parquet"
PRE2024_HISTORY_CACHE = CACHE_DIR / "readerlink_pre2024_pos_history.parquet"
POS_HISTORY_CACHE = CACHE_DIR / "readerlink_pos_history.parquet"

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

FILENAME_DATE_RE = re.compile(
    r"\b(\d{2})\s+(\d{4})\s+(\d{6})\s+Readerlink(?:\b|_)", re.IGNORECASE
)
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
    date_token = match.group(3)
    week_end = pd.Timestamp(datetime.strptime(date_token, "%m%d%y"))
    bookscan = bookscan_week(week_end)
    week_number = bookscan.week
    year = bookscan.year
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


def find_store_inventory_files() -> list[SourceFile]:
    if not STORE_INVENTORY_SOURCE_DIR.exists():
        print(f"Warning: missing store inventory folder: {STORE_INVENTORY_SOURCE_DIR}")
        return []
    files = []
    for path in STORE_INVENTORY_SOURCE_DIR.glob("*.xlsx"):
        if path.name.startswith("~$"):
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


def load_cached_sales() -> pd.DataFrame | None:
    if not WEEKLY_SALES_CACHE.exists():
        return None
    df = pd.read_parquet(WEEKLY_SALES_CACHE)
    df["week_end"] = pd.to_datetime(df["week_end"])
    return df


def build_incremental_sales_cache(files: list[SourceFile], rebuild_all: bool) -> pd.DataFrame:
    current_source_by_week = {
        source.week_end.normalize(): source.path.name
        for source in files
    }
    cached = None if rebuild_all else load_cached_sales()
    if cached is None:
        print("Sales cache: no reusable cache found; reading all sales source files.")
        files_to_read = files
    else:
        cached_sources = set(cached["source_file"].astype("string").dropna())
        cached_weeks = set(pd.to_datetime(cached["week_end"]).dt.normalize().dropna())
        files_to_read = [source for source in files if source.week_end.normalize() not in cached_weeks]
        print(f"Sales cache: {len(cached_sources):,} source file(s) already cached.")

    if not files_to_read:
        print("Sales cache: no new sales files to read.")
        sales = cached.copy()
        sales["week_end"] = pd.to_datetime(sales["week_end"])
        sales["source_file"] = (
            sales["week_end"].dt.normalize().map(current_source_by_week).fillna(sales["source_file"])
        )
        return (
            sales.sort_values(["week_end", "source_file"])
            .drop_duplicates(subset=["week_end", "master_chain", "isbn"], keep="last")
            .sort_values(["week_end", "master_chain", "isbn"])
            .reset_index(drop=True)
        )

    frames = []
    for index, source in enumerate(files_to_read, start=1):
        print(f"Reading sales {index:>3}/{len(files_to_read)}: {source.path.name}")
        frames.append(read_sales_file(source))

    new_sales = pd.concat(frames, ignore_index=True)
    sales = new_sales if cached is None else pd.concat([cached, new_sales], ignore_index=True, sort=False)
    sales["week_end"] = pd.to_datetime(sales["week_end"])
    sales["source_file"] = sales["week_end"].dt.normalize().map(current_source_by_week).fillna(sales["source_file"])
    sales = sales.drop_duplicates(
        subset=["week_end", "master_chain", "isbn"],
        keep="last",
    )
    return sales.sort_values(["week_end", "master_chain", "isbn"]).reset_index(drop=True)


def read_inventory_file(source: SourceFile) -> pd.DataFrame:
    workbook = pd.ExcelFile(source.path)
    if "Export" in workbook.sheet_names:
        df = pd.read_excel(workbook, sheet_name="Export", dtype={"ITEM": str})
    elif "rem dup" in workbook.sheet_names:
        df = pd.read_excel(workbook, sheet_name="rem dup", dtype={"EAN": str}).rename(
            columns={
                "EAN": "ITEM",
                "DC_OH_Tot": "RDS DC INVENTORY OH",
                "DC_OO_Tot": "RDS OPEN PO QUANTITY",
            }
        )
        print(f"Reading legacy Readerlink inventory layout: {source.path.name} [rem dup]")
    else:
        raise ValueError(
            f"Could not find a supported Readerlink inventory sheet in {source.path}"
        )
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


def build_incremental_inventory_history_cache(
    files: list[SourceFile], rebuild_all: bool
) -> pd.DataFrame:
    cached = None
    if INVENTORY_HISTORY_CACHE.exists() and not rebuild_all:
        cached = pd.read_parquet(INVENTORY_HISTORY_CACHE)
        cached["week_end"] = pd.to_datetime(cached["week_end"])
    cached_weeks = (
        set(pd.to_datetime(cached["week_end"]).dt.normalize().dropna())
        if cached is not None
        else set()
    )
    files_to_read = [source for source in files if source.week_end.normalize() not in cached_weeks]
    if files_to_read:
        frames = [read_inventory_file(source) for source in files_to_read]
        new_rows = pd.concat(frames, ignore_index=True)
        history = new_rows if cached is None else pd.concat([cached, new_rows], ignore_index=True)
    elif cached is not None:
        print("DC inventory history cache: no new inventory files to read.")
        history = cached.copy()
    else:
        history = pd.DataFrame()
    if history.empty:
        return history
    history["week_end"] = pd.to_datetime(history["week_end"])
    return (
        history.drop_duplicates(subset=["week_end", "isbn"], keep="last")
        .sort_values(["week_end", "isbn"])
        .reset_index(drop=True)
    )


def read_store_inventory_file(source: SourceFile) -> pd.DataFrame:
    required = ["MASTER CHAIN", "EAN", "TOTAL OH UNITS"]
    df = pd.read_excel(
        source.path,
        sheet_name="Export",
        usecols=required,
        dtype={"EAN": str},
    )
    df = df.rename(
        columns={
            "MASTER CHAIN": "master_chain",
            "EAN": "isbn",
            "TOTAL OH UNITS": "store_oh_units",
        }
    )
    df["master_chain"] = df["master_chain"].astype("string").str.strip()
    df["isbn"] = df["isbn"].map(normalize_isbn)
    df["store_oh_units"] = pd.to_numeric(
        df["store_oh_units"], errors="coerce"
    ).fillna(0)
    df = df[df["isbn"].ne("")]
    df = df.groupby(["master_chain", "isbn"], as_index=False)["store_oh_units"].sum()
    df["week_end"] = source.week_end
    df["readerlink_week"] = source.week_number
    df["readerlink_year"] = source.year
    df["source_file"] = source.path.name
    return df[
        [
            "week_end",
            "readerlink_week",
            "readerlink_year",
            "master_chain",
            "isbn",
            "store_oh_units",
            "source_file",
        ]
    ].sort_values(["week_end", "master_chain", "isbn"]).reset_index(drop=True)


def build_incremental_store_inventory_cache(
    files: list[SourceFile], rebuild_all: bool
) -> pd.DataFrame:
    cached = None
    if STORE_INVENTORY_CACHE.exists() and not rebuild_all:
        cached = pd.read_parquet(STORE_INVENTORY_CACHE)
        cached["week_end"] = pd.to_datetime(cached["week_end"])
    cached_weeks = (
        set(pd.to_datetime(cached["week_end"]).dt.normalize().dropna())
        if cached is not None
        else set()
    )
    files_to_read = [source for source in files if source.week_end.normalize() not in cached_weeks]
    if not files_to_read:
        print("Store inventory cache: no new store inventory files to read.")
        if cached is not None:
            return cached.copy()
        return pd.DataFrame(
            columns=[
                "week_end",
                "readerlink_week",
                "readerlink_year",
                "master_chain",
                "isbn",
                "store_oh_units",
                "source_file",
            ]
        )
    frames = [read_store_inventory_file(source) for source in files_to_read]
    new_rows = pd.concat(frames, ignore_index=True)
    history = new_rows if cached is None else pd.concat([cached, new_rows], ignore_index=True)
    history["week_end"] = pd.to_datetime(history["week_end"])
    return (
        history.drop_duplicates(
            subset=["week_end", "master_chain", "isbn"], keep="last"
        )
        .sort_values(["week_end", "master_chain", "isbn"])
        .reset_index(drop=True)
    )


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


def load_or_build_history(skip_history: bool, rebuild_all: bool) -> pd.DataFrame:
    if skip_history:
        if PRE2024_HISTORY_CACHE.exists():
            print(f"Pre-2024 history: using existing cache: {PRE2024_HISTORY_CACHE}")
            return pd.read_parquet(PRE2024_HISTORY_CACHE)
        print("Pre-2024 history: skipped and no existing cache found.")
        return pd.DataFrame(columns=["week_end", "master_chain", "isbn", "cy_pos_units", "source_file", "source_sheet"])

    if PRE2024_HISTORY_CACHE.exists() and not rebuild_all:
        print(f"Pre-2024 history: using existing cache: {PRE2024_HISTORY_CACHE}")
        return pd.read_parquet(PRE2024_HISTORY_CACHE)

    print(f"Pre-2024 history: reading historical workbook: {HISTORICAL_ROLLING_FILE}")
    history = build_historical_pos_cache()
    save_cache(history, PRE2024_HISTORY_CACHE)
    return history


def print_cache_plan(
    sales_files: list[SourceFile],
    inventory_files: list[SourceFile],
    store_inventory_files: list[SourceFile],
    rebuild_all: bool,
) -> None:
    print("")
    print("Readerlink cache inputs")
    print("-----------------------")
    print(f"Sales source folders:")
    for folder in SALES_SOURCE_DIRS:
        print(f"  - {folder}")
    print(f"Sales filenames found: {len(sales_files):,}")
    if sales_files:
        print(f"Latest sales file: {sales_files[-1].path.name} ({sales_files[-1].week_end:%m/%d/%Y})")
    print(f"Inventory source folder: {INVENTORY_SOURCE_DIR}")
    print(f"Inventory filenames found: {len(inventory_files):,}")
    if inventory_files:
        print(f"Latest inventory file: {inventory_files[-1].path.name} ({inventory_files[-1].week_end:%m/%d/%Y})")
    print(f"Store inventory source folder: {STORE_INVENTORY_SOURCE_DIR}")
    print(f"Store inventory filenames found: {len(store_inventory_files):,}")
    if store_inventory_files:
        latest_store = store_inventory_files[-1]
        print(f"Latest store inventory file: {latest_store.path.name} ({latest_store.week_end:%m/%d/%Y})")
    print(f"Existing sales cache: {'yes' if WEEKLY_SALES_CACHE.exists() else 'no'}")
    print(f"Existing pre-2024 history cache: {'yes' if PRE2024_HISTORY_CACHE.exists() else 'no'}")
    print(f"Mode: {'full rebuild' if rebuild_all else 'incremental weekly update'}")
    if not rebuild_all:
        print("Normal weekly run: prior sales history is read from cache; only uncached weekly sales files are opened.")
        print("Normal weekly run: uncached DC inventory and store-OH weeks are added to history; the latest DC compatibility cache is also refreshed.")
    print("")


def print_recent_week_summary(
    pos_history: pd.DataFrame,
    latest_inventory: pd.DataFrame | None = None,
    week_count: int = 4,
) -> None:
    if pos_history.empty:
        print(f"Last {week_count} cached weeks: no cached POS history found.")
        return

    data = pos_history.copy()
    data["week_end"] = pd.to_datetime(data["week_end"])
    data["cy_pos_units"] = pd.to_numeric(data["cy_pos_units"], errors="coerce").fillna(0)
    last_weeks = sorted(data["week_end"].dropna().unique())[-week_count:]
    summary = (
        data[data["week_end"].isin(last_weeks)]
        .groupby("week_end", as_index=False)
        .agg(
            pos_units=("cy_pos_units", "sum"),
            isbn_count=("isbn", "nunique"),
            chain_count=("master_chain", "nunique"),
            source_files=("source_file", "nunique"),
        )
        .sort_values("week_end", ascending=False)
    )

    print("")
    print(f"Last {week_count} cached Readerlink weeks")
    print("------------------------------")
    for row in summary.itertuples(index=False):
        bs = bookscan_week(row.week_end)
        print(
            f"{pd.Timestamp(row.week_end):%m/%d/%Y} | "
            f"Week {bs.week:02d} - {bs.year} | "
            f"POS units: {row.pos_units:,.0f} | "
            f"ISBNs: {row.isbn_count:,} | "
            f"Chains: {row.chain_count:,} | "
            f"Sales source files: {row.source_files:,}"
        )
    if latest_inventory is not None and not latest_inventory.empty:
        latest_inventory = latest_inventory.copy()
        latest_inventory["rds_dc_inventory_oh"] = pd.to_numeric(
            latest_inventory["rds_dc_inventory_oh"], errors="coerce"
        ).fillna(0)
        latest_inventory["rds_open_po_quantity"] = pd.to_numeric(
            latest_inventory["rds_open_po_quantity"], errors="coerce"
        ).fillna(0)
        print(
            "Latest inventory totals | "
            f"OH: {latest_inventory['rds_dc_inventory_oh'].sum():,.0f} | "
            f"OO: {latest_inventory['rds_open_po_quantity'].sum():,.0f} | "
            f"ISBNs: {latest_inventory['isbn'].nunique():,}"
        )
    print("")


def confirm_latest_files(
    sales_files: list[SourceFile],
    inventory_files: list[SourceFile],
    store_inventory_files: list[SourceFile],
) -> bool:
    latest_sales = sales_files[-1]
    latest_inventory = inventory_files[-1]
    sales_cached = False
    if WEEKLY_SALES_CACHE.exists():
        cached_sales = pd.read_parquet(WEEKLY_SALES_CACHE, columns=["week_end"])
        cached_weeks = set(pd.to_datetime(cached_sales["week_end"]).dt.normalize().dropna())
        sales_cached = latest_sales.week_end.normalize() in cached_weeks

    print("")
    print("Readerlink files selected for cache update")
    print("------------------------------------------")
    print(f"Latest sales file: {latest_sales.path.name} ({latest_sales.week_end:%m/%d/%Y})")
    print(f"Latest inventory file: {latest_inventory.path.name} ({latest_inventory.week_end:%m/%d/%Y})")
    latest_store = store_inventory_files[-1]
    print(f"Latest store inventory file: {latest_store.path.name} ({latest_store.week_end:%m/%d/%Y})")
    print(f"Latest sales week already cached: {'yes' if sales_cached else 'no'}")
    print("The cache update will add uncached sales and store-OH weeks, and refresh DC inventory from the latest file.")
    while True:
        try:
            choice = input("Add/update Readerlink cache with these files? (y/n): ").strip().lower()
        except EOFError:
            choice = "n"
        if choice in {"y", "yes"}:
            return True
        if choice in {"n", "no"}:
            print("Readerlink cache update skipped. Returning to the Readerlink menu.")
            return False
        print("Invalid choice. Please enter y or n.")


def run_add_new_week(skip_history: bool, rebuild_all: bool) -> bool:
    sales_files = find_sales_files()
    inventory_files = find_inventory_files()
    store_inventory_files = find_store_inventory_files()
    if not sales_files:
        raise FileNotFoundError("No Readerlink sales source files found.")
    if not inventory_files:
        raise FileNotFoundError("No Readerlink inventory source files found.")
    if not store_inventory_files:
        raise FileNotFoundError("No Readerlink store inventory source files found.")

    print_cache_plan(sales_files, inventory_files, store_inventory_files, rebuild_all)
    if not confirm_latest_files(sales_files, inventory_files, store_inventory_files):
        return False

    sales = build_incremental_sales_cache(sales_files, rebuild_all)
    save_cache(sales, WEEKLY_SALES_CACHE)

    inventory_history = build_incremental_inventory_history_cache(
        inventory_files, rebuild_all
    )
    save_cache(inventory_history, INVENTORY_HISTORY_CACHE)
    latest_inventory_week = pd.to_datetime(inventory_history["week_end"]).max()
    inventory = inventory_history[
        pd.to_datetime(inventory_history["week_end"]).eq(latest_inventory_week)
    ].copy()
    save_cache(inventory, LATEST_INVENTORY_CACHE)

    store_inventory = build_incremental_store_inventory_cache(
        store_inventory_files, rebuild_all
    )
    save_cache(store_inventory, STORE_INVENTORY_CACHE)

    history = load_or_build_history(skip_history, rebuild_all)
    combined_pos = build_combined_pos_history(sales, history)
    save_cache(combined_pos, POS_HISTORY_CACHE)
    print_recent_week_summary(combined_pos, inventory, week_count=4)

    print("Readerlink cache update complete.")
    return True


def run_show_last_weeks() -> None:
    if not POS_HISTORY_CACHE.exists():
        raise FileNotFoundError(f"Readerlink POS cache not found: {POS_HISTORY_CACHE}")
    pos_history = pd.read_parquet(POS_HISTORY_CACHE)
    latest_inventory = pd.read_parquet(LATEST_INVENTORY_CACHE) if LATEST_INVENTORY_CACHE.exists() else None
    print_recent_week_summary(pos_history, latest_inventory, week_count=4)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build Readerlink rolling report source caches.")
    parser.add_argument(
        "--add-new-week",
        action="store_true",
        help="Preview latest Readerlink files, confirm, then add uncached weeks to cache.",
    )
    parser.add_argument(
        "--show-last-weeks",
        action="store_true",
        help="Show totals for the last 4 cached Readerlink weeks without reading source Excel files.",
    )
    parser.add_argument(
        "--skip-history",
        action="store_true",
        help="Skip pre-2024 history extraction from the old rolling workbook.",
    )
    parser.add_argument(
        "--rebuild-all",
        action="store_true",
        help="Re-read every Readerlink sales workbook and rebuild the pre-2024 history cache.",
    )
    args = parser.parse_args()

    if args.show_last_weeks:
        run_show_last_weeks()
        return

    if not run_add_new_week(args.skip_history, args.rebuild_all):
        raise SystemExit(10)


if __name__ == "__main__":
    main()
