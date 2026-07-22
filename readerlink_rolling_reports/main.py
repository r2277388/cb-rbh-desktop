from __future__ import annotations

import argparse
import sys
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd

warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
    module="openpyxl.styles.stylesheet",
)

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from amazon_rolling_reports.functions import fetch_data_from_db, get_connection
from shared.bookscan_calendar import bookscan_week


BASE_DIR = Path(__file__).resolve().parent
SHARED_BASE_DIR = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Readerlink")
CACHE_DIR = SHARED_BASE_DIR / "cache"
OUTPUT_DIR = Path(r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Readerlink")
HBG_INVENTORY_FILE = Path(r"G:\OPS\Inventory\Daily\Finance_Only\Inventory Detail.xlsx")
FROZEN_QUANTITIES_FILE = Path(r"G:\OPS\Inventory\Daily\Finance_Only\Frozen Quantities.xlsx")

POS_HISTORY_CACHE = CACHE_DIR / "readerlink_pos_history.parquet"
LATEST_INVENTORY_CACHE = CACHE_DIR / "readerlink_latest_inventory.parquet"
INVENTORY_HISTORY_CACHE = CACHE_DIR / "readerlink_inventory_history.parquet"
STORE_INVENTORY_CACHE = CACHE_DIR / "readerlink_store_inventory_history.parquet"
METADATA_HISTORY_CACHE = CACHE_DIR / "readerlink_metadata_history.parquet"
HBG_INVENTORY_HISTORY_CACHE = CACHE_DIR / "readerlink_hbg_inventory_history.parquet"
FREEZES_HISTORY_CACHE = CACHE_DIR / "readerlink_freezes_history.parquet"
STORE_INVENTORY_DIR = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Weekly reports\2026\Readerlink\OH_Store_TitlePerformanceReport"
)
STORE_INVENTORY_GLOB = "*.xlsx"
WEEKLY_REPORT_YEARS = 5
ACTIVE_ROW_YEARS = 3

TRACKED_CHAINS = [
    "AAFES",
    "BJS",
    "COSTCO",
    "FRED MEYER",
    "HUDSON",
    "KROGER",
    "MEIJER",
    "PARADIES",
    "SamsClub",
    "TARGET",
    "Walmart",
    "WHOLE FOODS",
]
CHAIN_ALIASES = {
    "TARGET STORES": "TARGET",
    "TARGET": "TARGET",
    "SAMS CLUB": "SamsClub",
    "SAM'S CLUB": "SamsClub",
    "SAMS": "SamsClub",
    "SAMSCLUB.COM": "SamsClub",
    "BJS": "BJS",
    "COSTCO": "COSTCO",
    "FRED MEYER": "FRED MEYER",
    "HUDSON": "HUDSON",
    "KROGER": "KROGER",
    "MEIJER": "MEIJER",
    "PARADIES": "PARADIES",
    "AAFES": "AAFES",
    "WALMART": "Walmart",
    "WM.COM": "Walmart",
    "WHOLE FOODS": "WHOLE FOODS",
}

PG_GROUPING_KEY = {
    ("CHL", "BK"): "CBKids",
    ("PTC", "BK"): "CBRPG",
    ("RID", "BK"): "CBRPG",
    ("GAM", "BK"): "CBRPG",
    ("FWN", "BK"): "CBBook",
    ("ART", "BK"): "CBBook",
    ("ENT", "BK"): "CBBook",
    ("LIF", "BK"): "CBBook",
    ("CCB", "BK"): "CBBook",
    ("CPB", "BK"): "CBBook",
    ("CPA", "BK"): "CBBook",
    ("BAR-LIF", "BK"): "CBBook",
    ("BAR-ENT", "BK"): "CBBook",
    ("CHL", "FT"): "CBKids",
    ("PTC", "FT"): "CBRPG",
    ("RID", "FT"): "CBRPG",
    ("GAM", "FT"): "CBRPG",
    ("FWN", "FT"): "CBGift",
    ("ART", "FT"): "CBGift",
    ("ENT", "FT"): "CBGift",
    ("LIF", "FT"): "CBGift",
    ("CCB", "FT"): "CBGift",
    ("CPB", "FT"): "CBGift",
    ("CPA", "FT"): "CBGift",
    ("BAR-LIF", "FT"): "CBGift",
    ("BAR-ENT", "FT"): "CBGift",
}


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


def bookscan_year_start(week_end: pd.Timestamp) -> pd.Timestamp:
    year = week_end.year
    jan1 = pd.Timestamp(year=year, month=1, day=1)
    days_to_sunday = (6 - jan1.weekday()) % 7
    first_sunday = jan1 + pd.Timedelta(days=days_to_sunday)
    if week_end < first_sunday:
        return bookscan_year_start(pd.Timestamp(year=year - 1, month=12, day=31))
    return first_sunday


def readerlink_week(date_value: pd.Timestamp) -> int:
    return bookscan_week(date_value).week


def canonical_chain(value: str) -> str:
    text = str(value).strip().upper()
    return CHAIN_ALIASES.get(text, text)


def sheet_chain_filter_name(value: str) -> str:
    canonical = canonical_chain(value)
    return canonical if canonical in TRACKED_CHAINS else "All Other Accounts"


def load_metadata() -> pd.DataFrame:
    query = """
    SELECT
        CASE
            WHEN i.PUBLISHER_CODE = 'Quadrille Publishing Limited' THEN 'Quadrille'
            ELSE i.PUBLISHER_CODE
        END AS Pub,
        i.PRODUCT_TYPE AS pt,
        i.FORMAT AS ft,
        CASE
            WHEN LEFT(i.PUBLISHING_GROUP,3) = 'BAR' THEN 'BAR'
            ELSE i.PUBLISHING_GROUP
        END AS pgrp,
        i.ITEM_TITLE AS ISBN,
        i.SHORT_TITLE AS Title,
        i.PRICE_AMOUNT AS Price,
        CONVERT(varchar(10), COALESCE(i.AMORTIZATION_DATE, osd.OSD), 110) AS PubDate
    FROM ebs.item i
    LEFT JOIN (
        SELECT
            tt.ean13 AS ISBN,
            MAX(tt.active_datevalue) AS OSD
        FROM tmm.cb_Import_Title_Tasks tt
        WHERE tt.date_desc = 'On Sale Date'
          AND tt.active_datevalue IS NOT NULL
          AND tt.printingnumber = 1
          AND tt.activeind = '1'
        GROUP BY tt.ean13
    ) osd
        ON osd.ISBN = i.ITEM_TITLE
    WHERE i.PRODUCT_TYPE IN ('BK','FT','CP','RP','DI')
      AND i.PUBLISHING_GROUP NOT IN ('MKT', 'ZZZ')
      AND i.PUBLISHER_CODE NOT IN (
            'Benefit','AFO LLC','Glam Media','PQ Blackwell','PRINCETON','AMMO Books',
            'San Francisco Art Institute','FareArts','Sager','In Active','Driscolls',
            'Impossible Foods','Moleskine'
      );
    """
    df = fetch_data_from_db(get_connection(), query)
    df["ISBN"] = df["ISBN"].map(normalize_isbn)
    df = df[df["ISBN"] != ""]
    df = df.drop_duplicates(subset="ISBN", keep="first")
    return df


def pg_grouping(row: pd.Series) -> str:
    pub = str(row.get("Pub", "")).strip()
    pgrp = str(row.get("pgrp", "")).strip()
    pt = str(row.get("pt", "")).strip()
    if pub != "Chronicle":
        return pub
    return PG_GROUPING_KEY.get((pgrp, pt), "")


def load_hbg_inventory() -> pd.DataFrame:
    df = pd.read_excel(
        HBG_INVENTORY_FILE,
        usecols=["ISBN", "Available To Sell"],
        dtype={"ISBN": str},
    )
    df["ISBN"] = df["ISBN"].map(normalize_isbn)
    df = df[df["ISBN"] != ""]
    df["Available To Sell"] = pd.to_numeric(df["Available To Sell"], errors="coerce").fillna(0)
    df = df.groupby("ISBN", as_index=False)[["Available To Sell"]].sum()
    return df.rename(columns={"Available To Sell": "HBG_Avail"})


def load_readerlink_freezes() -> pd.DataFrame:
    df = pd.read_excel(
        FROZEN_QUANTITIES_FILE,
        dtype={"ISBN": str},
    )
    required_cols = {"ISBN", "Reason", "Requestor", "Current Stock Freeze", "Future Reserve Qty"}
    missing_cols = required_cols - set(df.columns)
    if missing_cols:
        missing = ", ".join(sorted(missing_cols))
        raise ValueError(f"Missing expected Frozen Quantities column(s): {missing}")

    df["ISBN"] = df["ISBN"].map(normalize_isbn)
    reason = df["Reason"].astype("string").str.strip().str.upper()
    requestor = df["Requestor"].astype("string").str.strip().str.upper().str.replace(".", "", regex=False)
    df = df[
        df["ISBN"].ne("")
        & reason.eq("READERLINK")
        & requestor.isin({"TRACY", "TRACY V", "TRACY VEGA"})
    ].copy()
    for column in ["Current Stock Freeze", "Future Reserve Qty"]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)
    df["RL Freezes"] = df["Current Stock Freeze"] + df["Future Reserve Qty"]
    return df.groupby("ISBN", as_index=False)[["RL Freezes"]].sum()


def load_or_create_weekly_snapshot(
    cache_path: Path,
    week: pd.Timestamp,
    live_loader,
) -> pd.DataFrame:
    snapshot_week = pd.Timestamp(week).normalize()
    history = pd.read_parquet(cache_path) if cache_path.exists() else pd.DataFrame()
    if not history.empty and "snapshot_week" in history.columns:
        history["snapshot_week"] = pd.to_datetime(history["snapshot_week"]).dt.normalize()
        cached = history[history["snapshot_week"].eq(snapshot_week)].copy()
        if not cached.empty:
            return cached.drop(columns=["snapshot_week"])

    current = live_loader().copy()
    snapshot = current.copy()
    snapshot.insert(0, "snapshot_week", snapshot_week)
    if not history.empty:
        history = history[~history["snapshot_week"].eq(snapshot_week)]
        snapshot = pd.concat([history, snapshot], ignore_index=True, sort=False)
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    snapshot.to_parquet(cache_path, index=False)
    print(f"Archived Readerlink weekly input snapshot: {cache_path.name} ({snapshot_week:%m/%d/%Y})")
    return current


def load_readerlink_inventory(week: pd.Timestamp | None = None) -> pd.DataFrame:
    cache_path = INVENTORY_HISTORY_CACHE if INVENTORY_HISTORY_CACHE.exists() else LATEST_INVENTORY_CACHE
    df = pd.read_parquet(cache_path)
    if week is not None and "week_end" in df.columns:
        df["week_end"] = pd.to_datetime(df["week_end"]).dt.normalize()
        selected = df[df["week_end"].eq(pd.Timestamp(week).normalize())]
        if not selected.empty:
            df = selected.copy()
    df = df.rename(
        columns={
            "isbn": "ISBN",
            "rds_dc_inventory_oh": "OH_DC",
            "rds_open_po_quantity": "OO",
        }
    )
    df["ISBN"] = df["ISBN"].map(normalize_isbn)
    return df.groupby("ISBN", as_index=False)[["OH_DC", "OO"]].sum()


def latest_store_inventory_file() -> Path:
    files = [
        path for path in STORE_INVENTORY_DIR.glob(STORE_INVENTORY_GLOB)
        if not path.name.startswith("~$")
    ]
    if not files:
        raise FileNotFoundError(
            f"No Readerlink store inventory workbook found in {STORE_INVENTORY_DIR}"
        )
    return max(files, key=lambda path: path.stat().st_mtime)


def load_store_inventory(
    source_path: Path | None = None, week: pd.Timestamp | None = None
) -> pd.DataFrame:
    if source_path is None and STORE_INVENTORY_CACHE.exists():
        df = pd.read_parquet(STORE_INVENTORY_CACHE)
        df["week_end"] = pd.to_datetime(df["week_end"])
        requested_week = pd.Timestamp(week).normalize() if week is not None else df["week_end"].max()
        selected = df[df["week_end"].eq(requested_week)]
        df = (selected if not selected.empty else df[df["week_end"].eq(df["week_end"].max())]).copy()
        df = df.rename(
            columns={
                "master_chain": "MASTER CHAIN",
                "isbn": "EAN",
                "store_oh_units": "TOTAL OH UNITS",
            }
        )
    else:
        path = source_path or latest_store_inventory_file()
        required = ["MASTER CHAIN", "EAN", "TOTAL OH UNITS"]
        df = pd.read_excel(
            path, sheet_name="Export", usecols=required, dtype={"EAN": str}
        )
    df["ISBN"] = df["EAN"].map(normalize_isbn)
    df["sheet_chain"] = df["MASTER CHAIN"].map(sheet_chain_filter_name)
    df["OH_Store"] = pd.to_numeric(df["TOTAL OH UNITS"], errors="coerce").fillna(0)
    df = df[df["ISBN"] != ""]
    return df.groupby(["sheet_chain", "ISBN"], as_index=False)[["OH_Store"]].sum()


def load_pos_history() -> pd.DataFrame:
    df = pd.read_parquet(POS_HISTORY_CACHE)
    df["ISBN"] = df["isbn"].map(normalize_isbn)
    df = df[df["ISBN"] != ""].copy()
    df["week_end"] = pd.to_datetime(df["week_end"])
    df["source_sheet"] = df.get("source_sheet", "").astype("string").fillna("")
    df["sheet_chain"] = df["master_chain"].map(sheet_chain_filter_name)
    df.loc[df["source_sheet"].eq("Summary - All Accounts"), "sheet_chain"] = "Summary - All Accounts"
    df["cy_pos_units"] = pd.to_numeric(df["cy_pos_units"], errors="coerce").fillna(0)
    return df


def history_columns(pos_history: pd.DataFrame) -> list[str]:
    dates = sorted(pos_history["week_end"].dropna().unique(), reverse=True)
    return [pd.Timestamp(value).strftime("%m-%d-%Y") for value in dates]


def report_week_columns(pos_history: pd.DataFrame, latest_week: pd.Timestamp) -> list[str]:
    cutoff = pd.Timestamp(year=latest_week.year - WEEKLY_REPORT_YEARS, month=1, day=1)
    dates = sorted(
        [
            pd.Timestamp(value)
            for value in pos_history["week_end"].dropna().unique()
            if pd.Timestamp(value) >= cutoff
        ],
        reverse=True,
    )
    return [value.strftime("%m-%d-%Y") for value in dates]


def older_year_columns(pos_history: pd.DataFrame, latest_week: pd.Timestamp) -> list[str]:
    cutoff = pd.Timestamp(year=latest_week.year - WEEKLY_REPORT_YEARS, month=1, day=1)
    years = sorted(
        {
            pd.Timestamp(value).year
            for value in pos_history["week_end"].dropna().unique()
            if pd.Timestamp(value) < cutoff
        },
        reverse=True,
    )
    return [str(year) for year in years]


def build_sheet_history(pos_history: pd.DataFrame, sheet_name: str, all_week_cols: list[str]) -> pd.DataFrame:
    if sheet_name == "Summary - All Accounts":
        source = pos_history[
            pos_history["source_sheet"].eq("Export/Data")
            | pos_history["sheet_chain"].eq("Summary - All Accounts")
        ].copy()
    elif sheet_name == "All Other Accounts":
        source = pos_history[pos_history["sheet_chain"] == "All Other Accounts"].copy()
    else:
        source = pos_history[pos_history["sheet_chain"] == sheet_name].copy()

    if source.empty:
        return pd.DataFrame(columns=["ISBN"] + all_week_cols)

    source["week_label"] = source["week_end"].dt.strftime("%m-%d-%Y")
    pivot = (
        source.groupby(["ISBN", "week_label"], as_index=False)["cy_pos_units"]
        .sum()
        .pivot(index="ISBN", columns="week_label", values="cy_pos_units")
        .fillna(0)
    )
    pivot = pivot.reindex(columns=all_week_cols, fill_value=0)
    return pivot.reset_index()


def build_report_sheet(
    sheet_name: str,
    pos_history: pd.DataFrame,
    metadata: pd.DataFrame,
    hbg_inventory: pd.DataFrame,
    readerlink_freezes: pd.DataFrame,
    readerlink_inventory: pd.DataFrame,
    store_inventory: pd.DataFrame,
    all_week_cols: list[str],
    report_week_cols: list[str],
    older_year_cols: list[str],
    latest_week: pd.Timestamp,
) -> pd.DataFrame:
    history = build_sheet_history(pos_history, sheet_name, all_week_cols)
    inventory_isbns = readerlink_inventory[
        pd.to_numeric(readerlink_inventory["OH_DC"], errors="coerce").fillna(0).ne(0)
        | pd.to_numeric(readerlink_inventory["OO"], errors="coerce").fillna(0).ne(0)
    ][["ISBN"]].drop_duplicates()
    history = pd.concat([history, inventory_isbns], ignore_index=True, sort=False)
    if sheet_name == "Summary - All Accounts":
        sheet_store_inventory = store_inventory.groupby("ISBN", as_index=False)[["OH_Store"]].sum()
        history = pd.concat(
            [history, sheet_store_inventory.loc[sheet_store_inventory["OH_Store"].ne(0), ["ISBN"]]],
            ignore_index=True,
            sort=False,
        )
    else:
        sheet_store_inventory = store_inventory[store_inventory["sheet_chain"].eq(sheet_name)][
            ["ISBN", "OH_Store"]
        ]
        history = pd.concat(
            [history, sheet_store_inventory.loc[sheet_store_inventory["OH_Store"].ne(0), ["ISBN"]]],
            ignore_index=True,
            sort=False,
        )
    history["ISBN"] = history["ISBN"].map(normalize_isbn)
    history = history[history["ISBN"] != ""].drop_duplicates(subset="ISBN", keep="first")
    for column in all_week_cols:
        if column not in history.columns:
            history[column] = 0
    history[all_week_cols] = history[all_week_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

    report = history.merge(metadata, on="ISBN", how="left")
    report = report[report["Pub"].notna() & report["Pub"].astype(str).str.strip().ne("")]
    report = report.merge(hbg_inventory, on="ISBN", how="left")
    report = report.merge(readerlink_freezes, on="ISBN", how="left")
    report = report.merge(readerlink_inventory, on="ISBN", how="left")
    report = report.merge(sheet_store_inventory, on="ISBN", how="left")

    for column in ["HBG_Avail", "RL Freezes", "OH_DC", "OH_Store", "OO", "Price"]:
        report[column] = pd.to_numeric(report.get(column, 0), errors="coerce").fillna(0)
    report["OH_ALL"] = report["OH_DC"] + report["OH_Store"]

    for column in ["Pub", "pt", "ft", "pgrp", "Title", "PubDate"]:
        if column not in report.columns:
            report[column] = ""
        report[column] = report[column].fillna("")

    latest_label = latest_week.strftime("%m-%d-%Y")
    prior_label = (latest_week - pd.Timedelta(weeks=1)).strftime("%m-%d-%Y")
    last_52 = [
        col
        for col in all_week_cols
        if latest_week - pd.Timedelta(weeks=51) <= pd.to_datetime(col, format="%m-%d-%Y") <= latest_week
    ]
    last_6 = [
        col
        for col in all_week_cols
        if latest_week - pd.Timedelta(weeks=5) <= pd.to_datetime(col, format="%m-%d-%Y") <= latest_week
    ]

    current_year_start = bookscan_year_start(latest_week)
    current_bs = bookscan_week(latest_week)
    tytd_cols = [
        col
        for col in all_week_cols
        if current_year_start <= pd.to_datetime(col, format="%m-%d-%Y") <= latest_week
    ]
    lytd_cols = [
        col
        for col in all_week_cols
        if bookscan_week(pd.to_datetime(col, format="%m-%d-%Y")).year == current_bs.year - 1
        and bookscan_week(pd.to_datetime(col, format="%m-%d-%Y")).week <= current_bs.week
    ]
    ly_fy_cols = [
        col
        for col in all_week_cols
        if bookscan_week(pd.to_datetime(col, format="%m-%d-%Y")).year == current_bs.year - 1
    ]
    active_cutoff = latest_week - pd.DateOffset(years=ACTIVE_ROW_YEARS)
    active_sales_cols = [
        col
        for col in all_week_cols
        if active_cutoff <= pd.to_datetime(col, format="%m-%d-%Y") <= latest_week
    ]
    account_activity_cutoff = latest_week - pd.DateOffset(years=5)
    account_activity_cols = [
        col
        for col in all_week_cols
        if account_activity_cutoff <= pd.to_datetime(col, format="%m-%d-%Y") <= latest_week
    ]

    for year_label in older_year_cols:
        year_cols = [
            col
            for col in all_week_cols
            if pd.to_datetime(col, format="%m-%d-%Y").year == int(year_label)
        ]
        report[year_label] = report[year_cols].sum(axis=1) if year_cols else 0

    report["PG_Grouping"] = report.apply(pg_grouping, axis=1)
    report["W52"] = report[last_52].sum(axis=1)
    report["6Wk Avg"] = (report[last_6].sum(axis=1) / 6.0).round(2)
    report["TYTD"] = report[tytd_cols].sum(axis=1)
    report["LYTD"] = report[lytd_cols].sum(axis=1)
    report["YTD Var"] = report["TYTD"] - report["LYTD"]
    report["TYTD Val"] = report["TYTD"] * report["Price"]
    report["LYTD Val"] = report["LYTD"] * report["Price"]
    report["YTD Val Var"] = report["TYTD Val"] - report["LYTD Val"]
    report["LY_FY"] = report[ly_fy_cols].sum(axis=1)
    report["LTD"] = report[all_week_cols].sum(axis=1)
    report["OH_Avg"] = (report["OH_DC"] / report["6Wk Avg"].where(report["6Wk Avg"].ne(0))).fillna(0).round(2)
    report["OH+OO_Avg"] = (
        (report["OH_DC"] + report["OO"]) / report["6Wk Avg"].where(report["6Wk Avg"].ne(0))
    ).fillna(0).round(2)
    if latest_label in report.columns and prior_label in report.columns:
        report["PW %"] = (
            (report[latest_label] - report[prior_label])
            / report[prior_label].where(report[prior_label].ne(0))
        ).fillna(0)
    else:
        report["PW %"] = 0

    if sheet_name == "Summary - All Accounts":
        active_mask = (
            report["OH_DC"].ne(0)
            | report["OH_Store"].ne(0)
            | report["OO"].ne(0)
            | report[active_sales_cols].sum(axis=1).ne(0)
        )
    else:
        active_mask = (
            report["OH_Store"].ne(0)
            | report[account_activity_cols].sum(axis=1).ne(0)
        )
    report = report[active_mask].copy()

    base_cols = [
        "Pub",
        "pt",
        "ft",
        "pgrp",
        "PG_Grouping",
        "ISBN",
        "Title",
        "Price",
        "PubDate",
        "HBG_Avail",
        "RL Freezes",
        "OH_DC",
        "OH_Store",
        "OH_ALL",
        "OO",
        "OH_Avg",
        "OH+OO_Avg",
        "W52",
        "6Wk Avg",
        "TYTD",
        "LYTD",
        "YTD Var",
        "TYTD Val",
        "LYTD Val",
        "YTD Val Var",
        "LY_FY",
        "LTD",
        "PW %",
    ]
    report = report[base_cols + report_week_cols + older_year_cols]
    report = report.sort_values(by=latest_label, ascending=False).reset_index(drop=True)
    return report


def output_file_for(latest_week: pd.Timestamp) -> Path:
    bookscan = bookscan_week(latest_week)
    return OUTPUT_DIR / f"Week {bookscan.week:02d} - {bookscan.year} Rolling Readerlink ({latest_week.strftime('%m%d%y')}).xlsx"


def print_report_inputs(pos_history: pd.DataFrame, latest_week: pd.Timestamp, output_path: Path) -> None:
    week_count = pos_history["week_end"].nunique()
    isbn_count = pos_history["ISBN"].nunique()
    source_count = pos_history["source_file"].nunique() if "source_file" in pos_history.columns else 0
    latest_sales_sources = (
        pos_history.loc[pos_history["week_end"].eq(latest_week), "source_file"]
        .dropna()
        .astype(str)
        .drop_duplicates()
        .sort_values()
        .tolist()
    )
    inventory_source = ""
    if LATEST_INVENTORY_CACHE.exists():
        inventory_cache = pd.read_parquet(LATEST_INVENTORY_CACHE, columns=["source_file", "week_end"])
        if not inventory_cache.empty:
            latest_inventory_week = pd.to_datetime(inventory_cache["week_end"]).max()
            inventory_files = (
                inventory_cache.loc[pd.to_datetime(inventory_cache["week_end"]).eq(latest_inventory_week), "source_file"]
                .dropna()
                .astype(str)
                .drop_duplicates()
                .sort_values()
                .tolist()
            )
            inventory_source = ", ".join(inventory_files)

    bookscan = bookscan_week(latest_week)
    print("")
    print("Readerlink report inputs")
    print("------------------------")
    print(f"Latest sales week in cache: {latest_week:%m/%d/%Y} / Week {bookscan.week:02d} - {bookscan.year}")
    print(f"POS cache: {POS_HISTORY_CACHE}")
    print(f"POS cache coverage: {week_count:,} week(s), {isbn_count:,} ISBN(s), {source_count:,} source file(s)")
    print(f"Latest sales source file(s): {', '.join(latest_sales_sources) if latest_sales_sources else 'none found'}")
    print(f"Latest Readerlink inventory cache: {LATEST_INVENTORY_CACHE}")
    print(f"Latest Readerlink inventory source file(s): {inventory_source or 'none found'}")
    print(f"Latest Readerlink store inventory source: {latest_store_inventory_file()}")
    print(f"HBG_Avail source: {HBG_INVENTORY_FILE}")
    print(f"RL Freezes source: {FROZEN_QUANTITIES_FILE}")
    print("Title metadata source: SQL ebs.item with On Sale Date fallback")
    print(f"Output workbook: {output_path}")
    print("")


def write_workbook(sheets: dict[str, pd.DataFrame], output_path: Path, latest_week: pd.Timestamp) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    header_row = 4
    data_start_row = header_row + 1
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        week_header_fmt_a = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#D8E4BC", "border": 1})
        week_header_fmt_b = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#F2DCDB", "border": 1})
        year_header_fmt = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#E4DFEC", "border": 1})
        base_header_fmt = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#B8CCE4", "border": 1})
        weeknum_fmt = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#DDD9C4", "border": 1})
        total_label_fmt = workbook.add_format({"bold": True, "align": "left", "bg_color": "#CCC0DA", "border": 1})
        accounting_format = '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
        total_num_fmt = workbook.add_format({"bold": True, "num_format": accounting_format, "bg_color": "#E4DFEC", "border": 1})
        total_dec_fmt = workbook.add_format({"bold": True, "num_format": '#,##0.00', "bg_color": "#E4DFEC", "border": 1})
        total_pct_fmt = workbook.add_format({"bold": True, "num_format": '0.0%', "bg_color": "#E4DFEC", "border": 1})
        int_fmt = workbook.add_format({"num_format": accounting_format})
        plain_int_fmt = workbook.add_format({"num_format": '#,##0'})
        dec_fmt = workbook.add_format({"num_format": '#,##0.00'})
        pct_fmt = workbook.add_format({"num_format": '0.0%'})
        isbn_fmt = workbook.add_format({"num_format": '0'})
        title_fmt = workbook.add_format({"bold": False, "align": "center_across", "valign": "vcenter", "bg_color": "#C4BD97", "font_size": 16})
        unit_band_fmt = workbook.add_format({"bold": True, "align": "center_across", "valign": "vcenter", "bg_color": "#FCD5B4", "border": 1})
        value_band_fmt = workbook.add_format({"bold": True, "align": "center_across", "valign": "vcenter", "bg_color": "#B7DEE8", "border": 1})
        criteria_title_fmt = workbook.add_format({"bold": True, "font_size": 14, "bg_color": "#B8CCE4", "border": 1})
        criteria_header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAD3", "border": 1})
        criteria_text_fmt = workbook.add_format({"text_wrap": True, "valign": "top", "border": 1})

        for sheet_name, df in sheets.items():
            if sheet_name == "Criteria":
                df.to_excel(writer, sheet_name=sheet_name, startrow=2, index=False)
                ws = writer.sheets[sheet_name]
                ws.set_tab_color("#403151")
                ws.write(0, 0, "Readerlink Rolling Report Criteria", criteria_title_fmt)
                for col_idx, col_name in enumerate(df.columns):
                    ws.write(2, col_idx, col_name, criteria_header_fmt)
                ws.set_column(0, 0, 28)
                ws.set_column(1, 1, 120, criteria_text_fmt)
                ws.freeze_panes(3, 0)
                continue

            safe_sheet = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_sheet, startrow=header_row, index=False)
            ws = writer.sheets[safe_sheet]
            if sheet_name == "Summary - All Accounts":
                ws.set_tab_color("#0F243E")
            last_row = header_row + len(df)
            last_col = len(df.columns) - 1

            ws.write(0, 3, sheet_name, title_fmt)
            ws.write(1, 3, f"Rolling Readerlink POS - Week Ending: {latest_week.strftime('%B %d, %Y')}", title_fmt)
            for col_idx in range(4, 7):
                ws.write_blank(0, col_idx, None, title_fmt)
                ws.write_blank(1, col_idx, None, title_fmt)

            ws.write(0, 8, "Total", total_label_fmt)
            ws.write(1, 8, "Subtotal", total_label_fmt)
            col_positions = {col_name: idx for idx, col_name in enumerate(df.columns)}
            pre_week_cols = col_positions.get(latest_week.strftime("%m-%d-%Y"), len(df.columns))
            unit_cols = ["TYTD", "LYTD", "YTD Var"]
            value_cols = ["TYTD Val", "LYTD Val", "YTD Val Var"]
            if all(col in col_positions for col in unit_cols):
                start_col = col_positions[unit_cols[0]]
                ws.write(header_row - 1, start_col, "Unit", unit_band_fmt)
                for col_idx in range(start_col + 1, col_positions[unit_cols[-1]] + 1):
                    ws.write_blank(header_row - 1, col_idx, None, unit_band_fmt)
            if all(col in col_positions for col in value_cols):
                start_col = col_positions[value_cols[0]]
                ws.write(header_row - 1, start_col, "$ Value", value_band_fmt)
                for col_idx in range(start_col + 1, col_positions[value_cols[-1]] + 1):
                    ws.write_blank(header_row - 1, col_idx, None, value_band_fmt)
            latest_label = latest_week.strftime("%m-%d-%Y")
            prior_label = (latest_week - pd.Timedelta(weeks=1)).strftime("%m-%d-%Y")
            for col_idx, col_name in enumerate(df.columns):
                parsed = pd.to_datetime(col_name, format="%m-%d-%Y", errors="coerce")
                if col_idx < pre_week_cols:
                    fmt = base_header_fmt
                elif not pd.isna(parsed):
                    fmt = week_header_fmt_a if (col_idx - pre_week_cols) % 2 == 0 else week_header_fmt_b
                else:
                    fmt = year_header_fmt
                ws.write(header_row, col_idx, col_name, fmt)
                if col_idx >= pre_week_cols and not pd.isna(parsed):
                    ws.write(header_row - 1, col_idx, readerlink_week(parsed), fmt)

                if col_idx >= 9:
                    col_letter_start = _xl_cell(data_start_row, col_idx)
                    col_letter_end = _xl_cell(last_row, col_idx)
                    total_formula = f"=SUM({col_letter_start}:{col_letter_end})"
                    subtotal_formula = f"=SUBTOTAL(9,{col_letter_start}:{col_letter_end})"
                    value = float(pd.to_numeric(df[col_name], errors="coerce").fillna(0).sum()) if col_name in df else 0
                    if col_name == "PW %":
                        fmt_total = total_pct_fmt
                    elif col_name == "Price":
                        fmt_total = total_dec_fmt
                    else:
                        fmt_total = total_num_fmt

                    if col_name == "OH_Avg":
                        oh_total = _xl_cell(0, col_positions["OH_DC"])
                        avg_total = _xl_cell(0, col_positions["6Wk Avg"])
                        oh_subtotal = _xl_cell(1, col_positions["OH_DC"])
                        avg_subtotal = _xl_cell(1, col_positions["6Wk Avg"])
                        total_formula = f'=IFERROR({oh_total}/{avg_total},0)'
                        subtotal_formula = f'=IFERROR({oh_subtotal}/{avg_subtotal},0)'
                        total_6wk = pd.to_numeric(df["6Wk Avg"], errors="coerce").fillna(0).sum()
                        value = 0 if total_6wk == 0 else float(pd.to_numeric(df["OH_DC"], errors="coerce").fillna(0).sum() / total_6wk)
                    elif col_name == "OH+OO_Avg":
                        oh_total = _xl_cell(0, col_positions["OH_DC"])
                        oo_total = _xl_cell(0, col_positions["OO"])
                        avg_total = _xl_cell(0, col_positions["6Wk Avg"])
                        oh_subtotal = _xl_cell(1, col_positions["OH_DC"])
                        oo_subtotal = _xl_cell(1, col_positions["OO"])
                        avg_subtotal = _xl_cell(1, col_positions["6Wk Avg"])
                        total_formula = f'=IFERROR(({oh_total}+{oo_total})/{avg_total},0)'
                        subtotal_formula = f'=IFERROR(({oh_subtotal}+{oo_subtotal})/{avg_subtotal},0)'
                        total_6wk = pd.to_numeric(df["6Wk Avg"], errors="coerce").fillna(0).sum()
                        value = (
                            0
                            if total_6wk == 0
                            else float(
                                (
                                    pd.to_numeric(df["OH_DC"], errors="coerce").fillna(0).sum()
                                    + pd.to_numeric(df["OO"], errors="coerce").fillna(0).sum()
                                )
                                / total_6wk
                            )
                        )
                    elif col_name == "PW %" and latest_label in col_positions and prior_label in col_positions:
                        latest_total = _xl_cell(0, col_positions[latest_label])
                        prior_total = _xl_cell(0, col_positions[prior_label])
                        latest_subtotal = _xl_cell(1, col_positions[latest_label])
                        prior_subtotal = _xl_cell(1, col_positions[prior_label])
                        total_formula = f'=IFERROR(({latest_total}-{prior_total})/{prior_total},0)'
                        subtotal_formula = f'=IFERROR(({latest_subtotal}-{prior_subtotal})/{prior_subtotal},0)'
                        prior_sum = pd.to_numeric(df[prior_label], errors="coerce").fillna(0).sum()
                        value = (
                            0
                            if prior_sum == 0
                            else float(
                                (
                                    pd.to_numeric(df[latest_label], errors="coerce").fillna(0).sum()
                                    - prior_sum
                                )
                                / prior_sum
                            )
                        )
                    ws.write_formula(0, col_idx, total_formula, fmt_total, value)
                    ws.write_formula(1, col_idx, subtotal_formula, fmt_total, value)

            ws.write(header_row - 1, pre_week_cols - 1, "WeekNum", weeknum_fmt)
            ws.autofilter(header_row, 0, last_row, last_col)
            ws.freeze_panes(data_start_row, 0)
            ws.set_column(0, 4, 10)
            ws.set_column(5, 5, 13.5, isbn_fmt)
            ws.set_column(6, 6, 42)
            ws.set_column(7, 7, 9, dec_fmt)
            ws.set_column(8, 8, 11)
            ws.set_column(9, 9, 11, plain_int_fmt)
            ws.set_column(10, last_col, 11, int_fmt)
            ws.set_column(col_positions["PW %"], col_positions["PW %"], 9, pct_fmt)

    print(f"Saved Readerlink rolling report: {output_path}")


def ordered_sheets_by_tytd(sheets: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    ordered = {"Summary - All Accounts": sheets["Summary - All Accounts"]}
    remaining = [
        (name, df)
        for name, df in sheets.items()
        if name != "Summary - All Accounts"
    ]
    remaining.sort(
        key=lambda item: pd.to_numeric(item[1].get("TYTD", pd.Series(dtype=float)), errors="coerce").fillna(0).sum(),
        reverse=True,
    )
    ordered.update(remaining)
    return ordered


def build_criteria_sheet(latest_week: pd.Timestamp) -> pd.DataFrame:
    rows = [
        (
            "Report week",
            f"Latest week is taken from the max week_end in the Readerlink POS cache. Current build week ending: {latest_week.strftime('%B %d, %Y')}.",
        ),
        (
            "Sales source",
            "Weekly Readerlink sales workbooks from the 2026/2025 Readerlink folders and the 2024 archive folder. Modern files use the Export/Data layout; 2024 archive files use the LAST WEEK layout.",
        ),
        (
            "Sales metric",
            "Weekly units come from CY POS UNITS. ISBN is normalized from EAN/ITEMNUMBER to a 13-digit text ISBN.",
        ),
        (
            "Inventory source",
            f"Readerlink OH_DC and OO use the report week's snapshot in {INVENTORY_HISTORY_CACHE}. OH_DC = RDS DC INVENTORY OH; OO = RDS OPEN PO QUANTITY. {LATEST_INVENTORY_CACHE} is retained as a latest-week compatibility cache.",
        ),
        (
            "Store inventory source",
            f"OH_Store comes from the latest weekly snapshot in {STORE_INVENTORY_CACHE}, built from Excel workbooks in {STORE_INVENTORY_DIR}. It is grouped by MASTER CHAIN and EAN for account tabs, and across all master chains for Summary - All Accounts. OH_ALL = OH_DC + OH_Store.",
        ),
        (
            "HBG availability",
            f"HBG_Avail uses the report week's snapshot in {HBG_INVENTORY_HISTORY_CACHE}, originally read from {HBG_INVENTORY_FILE}, field Available To Sell, grouped by ISBN.",
        ),
        (
            "RL Freezes",
            f"RL Freezes uses the report week's snapshot in {FREEZES_HISTORY_CACHE}, originally read from {FROZEN_QUANTITIES_FILE}. Included rows require Reason = Readerlink and Requestor in Tracy, Tracy V, Tracy Vega, or Tracy V. Value = Current Stock Freeze + Future Reserve Qty, grouped by ISBN.",
        ),
        (
            "Title metadata",
            f"Pub, pt, ft, pgrp, ISBN, Title, Price, and PubDate use the report week's snapshot in {METADATA_HISTORY_CACHE}, originally queried from ebs.item. PubDate uses AMORTIZATION_DATE with fallback to the On Sale Date task table when AMORTIZATION_DATE is blank.",
        ),
        (
            "Metadata filter",
            "Only EBS items with PRODUCT_TYPE in BK, FT, CP, RP, DI are included. Publishing groups MKT and ZZZ are excluded, along with the publisher exclusion list used by the shared rolling-report metadata query.",
        ),
        (
            "Unknown ISBNs",
            "Readerlink ISBNs not found in the EBS item query are left out of the report, even if Readerlink sent sales or inventory for them.",
        ),
        (
            "Chain tabs",
            "Summary - All Accounts is always first. Remaining account tabs are sorted by total TYTD units descending. Criteria is always last.",
        ),
        (
            "Account rollups",
            "SAMS, SAMS CLUB, SAM'S CLUB, and SAMSCLUB.COM roll into SamsClub. WALMART and WM.COM roll into Walmart. Untracked accounts roll into All Other Accounts.",
        ),
        (
            "Inventory on tabs",
            "Readerlink OH_DC, OO, HBG_Avail, and RL Freezes are ISBN-level values repeated on every account tab. OH_Store is specific to each Master Chain; the summary uses total store OH across all chains.",
        ),
        (
            "Displayed weekly history",
            f"The report displays complete weekly years only using the current rolling setting of {WEEKLY_REPORT_YEARS} years. Older weeks stay in cache and are shown as annual aggregate columns.",
        ),
        (
            "Row inclusion",
            f"Summary rows are included when the ISBN is in EBS and has any OH_DC, OH_Store, OO, or sales in the last {ACTIVE_ROW_YEARS} years. Account tabs include ISBNs with that account's OH_Store or sales activity in the last 5 years. Cache data is retained even when rows are filtered from the report.",
        ),
        (
            "Calculations",
            "W52 is the latest 52 weeks. 6Wk Avg is the latest 6 weeks divided by 6. Existing coverage calculations remain DC-based: OH_Avg = OH_DC / 6Wk Avg and OH+OO_Avg = (OH_DC + OO) / 6Wk Avg. PW % = current week versus prior week.",
        ),
        (
            "YTD logic",
            "TYTD, LYTD, LY_FY, and related value fields follow the shared Bookscan calendar logic used by the other rolling reports.",
        ),
        (
            "Caches",
            f"Shared Readerlink caches live in {CACHE_DIR}. The cache keeps all collected Readerlink data for future report builds.",
        ),
    ]
    return pd.DataFrame(rows, columns=["Criteria", "Rule / Source"])


def _xl_cell(row: int, col: int) -> str:
    from xlsxwriter.utility import xl_rowcol_to_cell

    return xl_rowcol_to_cell(row, col)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build the Readerlink rolling report workbook.")
    parser.add_argument("--output", type=Path, help="Optional output workbook path.")
    args = parser.parse_args()

    pos_history = load_pos_history()
    latest_week = pos_history["week_end"].max()
    week_cols = history_columns(pos_history)
    display_week_cols = report_week_columns(pos_history, latest_week)
    display_year_cols = older_year_columns(pos_history, latest_week)
    output_path = args.output or output_file_for(latest_week)
    print_report_inputs(pos_history, latest_week, output_path)
    metadata = load_or_create_weekly_snapshot(
        METADATA_HISTORY_CACHE, latest_week, load_metadata
    )
    hbg_inventory = load_or_create_weekly_snapshot(
        HBG_INVENTORY_HISTORY_CACHE, latest_week, load_hbg_inventory
    )
    readerlink_freezes = load_or_create_weekly_snapshot(
        FREEZES_HISTORY_CACHE, latest_week, load_readerlink_freezes
    )
    readerlink_inventory = load_readerlink_inventory(latest_week)
    store_inventory = load_store_inventory(week=latest_week)

    sheets: dict[str, pd.DataFrame] = {}
    for sheet_name in ["Summary - All Accounts"] + TRACKED_CHAINS + ["All Other Accounts"]:
        print(f"Building sheet: {sheet_name}")
        sheets[sheet_name] = build_report_sheet(
            sheet_name,
            pos_history,
            metadata,
            hbg_inventory,
            readerlink_freezes,
            readerlink_inventory,
            store_inventory,
            week_cols,
            display_week_cols,
            display_year_cols,
            latest_week,
        )

    ordered_sheets = ordered_sheets_by_tytd(sheets)
    ordered_sheets["Criteria"] = build_criteria_sheet(latest_week)
    write_workbook(ordered_sheets, output_path, latest_week)


if __name__ == "__main__":
    main()
