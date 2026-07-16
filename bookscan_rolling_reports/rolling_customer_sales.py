from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import Tk, filedialog

import pandas as pd

sys.path.append(str(Path(__file__).resolve().parents[1]))

from amazon_rolling_reports.functions import build_column_totals, save_to_excel
from bn_rolling_reports.isbn_utils import normalize_isbn_series
from shared.bookscan_calendar import bookscan_parts, bookscan_week
try:
    from .rolling_paths import (
        bookscan_dp_folders,
        bookscan_rolling_folder,
        inventory_cache_file,
        inventory_detail_workbook,
        local_review_dir,
        manual_missing_weeks_file,
        metadata_cache_file,
        sales_cache_file,
    )
    from .rolling_queries import (
        DISTINCT_WEEKS_QUERY,
        LATEST_WEEK_QUERY,
        MISSING_WEEKS_QUERY,
        RECENT_WEEK_SUMMARY_QUERY,
        SOURCE_METADATA_QUERY,
        build_distinct_weeks_since_query,
        build_sales_query,
    )
except ImportError:
    from rolling_paths import (
        bookscan_dp_folders,
        bookscan_rolling_folder,
        inventory_cache_file,
        inventory_detail_workbook,
        local_review_dir,
        manual_missing_weeks_file,
        metadata_cache_file,
        sales_cache_file,
    )
    from rolling_queries import (
        DISTINCT_WEEKS_QUERY,
        LATEST_WEEK_QUERY,
        MISSING_WEEKS_QUERY,
        RECENT_WEEK_SUMMARY_QUERY,
        SOURCE_METADATA_QUERY,
        build_distinct_weeks_since_query,
        build_sales_query,
    )
from shared.db.connection import get_connection
from shared.db.query_runner import fetch_data_from_db


WEEKLY_HISTORY_START = pd.Timestamp("2023-01-01")
YEARLY_HISTORY_START = 2007
YEARLY_HISTORY_END = 2022
DEFAULT_HISTORY_WORKBOOK_GLOB = "*Rolling Bookscan*.xlsx"
REFRESH_LOOKBACK_WEEKS = 1
METADATA_BATCH_SIZE = 500
PUBLISHER_NORMALIZATION = {
    "Quadrille Publishing Limited": "Quadrille",
}
PUBLISHER_EXCLUSIONS = (
    "Benefit",
    "AFO LLC",
    "Glam Media",
    "PQ Blackwell",
    "PRINCETON",
    "AMMO Books",
    "San Francisco Art Institute",
    "FareArts",
    "Sager",
    "In Active",
    "Driscolls",
    "Impossible Foods",
    "Moleskine",
)
PUBLISHER_EXCLUSION_KEYS = {value.strip().casefold() for value in PUBLISHER_EXCLUSIONS}
PUBLISHER_EXCLUSION_KEYS.add("princeton")


@dataclass
class RollingBuildResult:
    output_file: Path
    latest_week: pd.Timestamp
    sales_rows: int
    inventory_rows: int
    report_shape: tuple[int, int]
    dp_files_saved: int = 0


@dataclass
class DpSaveResult:
    source_file: Path
    latest_week: pd.Timestamp
    dp_files_saved: int


@dataclass
class CacheRefreshResult:
    refresh_mode: str
    latest_sql_week: pd.Timestamp | None
    sales_cache_week: pd.Timestamp | None
    inventory_cache_week: pd.Timestamp | None
    expected_next_week: pd.Timestamp | None
    missing_week_count: int
    missing_weeks: list[pd.Timestamp]
    sales_rows: int
    inventory_rows: int


@dataclass
class WeekCheckResult:
    min_week: pd.Timestamp | None
    max_week: pd.Timestamp | None
    missing_weeks: list[pd.Timestamp]
    latest_week_rows: list[dict[str, object]]


@dataclass
class DeltaWeekStatus:
    latest_sql_week: pd.Timestamp | None
    latest_cache_week: pd.Timestamp | None
    expected_next_week: pd.Timestamp | None
    missing_weeks: list[pd.Timestamp]


def _expected_next_week(week: pd.Timestamp | None) -> pd.Timestamp | None:
    if week is None:
        return None
    return pd.Timestamp(week) + pd.Timedelta(days=7)


def _ensure_cache_dir() -> None:
    sales_cache_file.parent.mkdir(parents=True, exist_ok=True)


def _load_parquet_or_empty(cache_file: Path) -> pd.DataFrame:
    if cache_file.exists():
        return pd.read_parquet(cache_file)
    return pd.DataFrame()


def _save_parquet(df: pd.DataFrame, cache_file: Path) -> None:
    _ensure_cache_dir()
    df.to_parquet(cache_file, index=False)


def resolve_inventory_detail_workbook(
    inventory_detail_workbook_override: str | Path | None = None,
) -> Path:
    if inventory_detail_workbook_override:
        override_path = Path(inventory_detail_workbook_override)
        if override_path.exists():
            return override_path
        raise FileNotFoundError(f"Inventory detail workbook not found: {override_path}")

    if inventory_detail_workbook.exists():
        return inventory_detail_workbook

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = filedialog.askopenfilename(
            title="Choose the current Inventory Detail workbook",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
    finally:
        root.destroy()

    if not selected:
        raise FileNotFoundError(
            "Inventory detail workbook was not accessible and no replacement file was selected."
        )
    return Path(selected)


def _normalize_publisher(value):
    if pd.isna(value):
        return value
    return PUBLISHER_NORMALIZATION.get(str(value), value)


def _filter_allowed_publishers(df: pd.DataFrame, column: str = "Pub") -> pd.DataFrame:
    if df.empty or column not in df.columns:
        return df
    publisher_keys = df[column].astype("string").str.strip().str.casefold()
    return df[df[column].isna() | ~publisher_keys.isin(PUBLISHER_EXCLUSION_KEYS)].copy()


def _isbn10_to_isbn13(value: str) -> str | None:
    clean = str(value).strip().replace("-", "").replace(" ", "").upper()
    if len(clean) != 10 or not clean[:9].isdigit():
        return None
    body = "978" + clean[:9]
    checksum = sum((1 if idx % 2 == 0 else 3) * int(ch) for idx, ch in enumerate(body))
    check_digit = (10 - (checksum % 10)) % 10
    return body + str(check_digit)


def normalize_bookscan_isbn(value: object) -> str | None:
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)) and not pd.isna(value):
        if float(value).is_integer():
            clean = str(int(value))
        else:
            clean = str(value)
    else:
        clean = str(value).strip()
        if clean.endswith(".0") and clean[:-2].isdigit():
            clean = clean[:-2]
        else:
            numeric = pd.to_numeric(clean, errors="coerce")
            if pd.notna(numeric) and float(numeric).is_integer():
                clean = str(int(numeric))
    clean = clean.replace("-", "").replace(" ", "")
    if not clean:
        return None
    if len(clean) == 10:
        return _isbn10_to_isbn13(clean)
    series = pd.Series([clean], dtype="string")
    normalized = normalize_isbn_series(series)
    return normalized.iloc[0]


def normalize_bookscan_isbn_series(series: pd.Series) -> pd.Series:
    normalized = series.map(normalize_bookscan_isbn)
    return normalized.astype("object")


def _format_bookscan_output_filename(week_ending: pd.Timestamp) -> str:
    return (
        f"Week {bookscan_week(week_ending).week:02d} - {bookscan_week(week_ending).year} "
        f"Rolling Bookscan ({week_ending:%m%d%y}).xlsx"
    )


def _parse_week_from_output_filename(path: Path) -> pd.Timestamp | None:
    match = re.search(r"\((\d{6})\)\.xlsx$", path.name)
    if not match:
        return None
    return pd.Timestamp(datetime.strptime(match.group(1), "%m%d%y").date())


def find_latest_saved_main_report() -> Path | None:
    candidates = sorted(bookscan_rolling_folder.glob("*Rolling Bookscan (*).xlsx"))
    if not candidates:
        return None
    return max(candidates, key=lambda path: path.stat().st_mtime)


def load_saved_main_report(source_file: str | Path | None = None) -> tuple[pd.DataFrame, pd.Timestamp, Path]:
    workbook_path = Path(source_file) if source_file else find_latest_saved_main_report()
    if workbook_path is None or not workbook_path.exists():
        raise FileNotFoundError("No saved main Bookscan report was found to create DP versions from.")

    report_df = pd.read_excel(workbook_path, header=5)
    latest_week = _parse_week_from_output_filename(workbook_path)
    if latest_week is None:
        date_columns = [
            pd.to_datetime(column, format="%m-%d-%Y", errors="coerce")
            for column in report_df.columns
            if isinstance(column, str) and len(column) == 10 and column.count("-") == 2
        ]
        valid_dates = [value for value in date_columns if not pd.isna(value)]
        if not valid_dates:
            raise ValueError(f"Could not determine the Bookscan week from {workbook_path.name}.")
        latest_week = max(valid_dates)
    return report_df, pd.Timestamp(latest_week), workbook_path


def find_default_history_workbook() -> Path | None:
    candidates = sorted(Path(__file__).resolve().parent.glob(DEFAULT_HISTORY_WORKBOOK_GLOB))
    if not candidates:
        return None
    return max(candidates, key=lambda path: path.stat().st_mtime)


def get_latest_sql_week() -> pd.Timestamp | None:
    engine = get_connection()
    result = fetch_data_from_db(engine, LATEST_WEEK_QUERY)
    if result.empty or "latest_week" not in result.columns or pd.isna(result.at[0, "latest_week"]):
        return None
    return pd.Timestamp(result.at[0, "latest_week"])


def get_latest_cache_week() -> pd.Timestamp | None:
    cached = _load_parquet_or_empty(sales_cache_file)
    if cached.empty or "Week" not in cached.columns:
        return None
    cached["Week"] = pd.to_datetime(cached["Week"], errors="coerce")
    if "HistoryType" in cached.columns:
        cached = cached[cached["HistoryType"].astype(str) == "weekly"].copy()
    latest = cached["Week"].dropna().max()
    if pd.isna(latest):
        return None
    return pd.Timestamp(latest)


def check_source_weeks() -> WeekCheckResult:
    engine = get_connection()
    week_result = fetch_data_from_db(engine, DISTINCT_WEEKS_QUERY)
    if week_result.empty:
        return WeekCheckResult(None, None, [], [])
    week_result["Week"] = pd.to_datetime(week_result["Week"])
    weeks = sorted(week_result["Week"].dropna().tolist())
    missing_result = fetch_data_from_db(engine, MISSING_WEEKS_QUERY)
    missing = []
    if not missing_result.empty and "missing_week" in missing_result.columns:
        missing_result["missing_week"] = pd.to_datetime(
            missing_result["missing_week"], errors="coerce"
        )
        missing = sorted(missing_result["missing_week"].dropna().tolist())
    recent_result = fetch_data_from_db(engine, RECENT_WEEK_SUMMARY_QUERY)
    recent_rows = recent_result.to_dict("records") if not recent_result.empty else []
    return WeekCheckResult(weeks[0], weeks[-1], missing, recent_rows)


def get_delta_week_status(
    latest_cache_week: pd.Timestamp | None = None,
    latest_sql_week: pd.Timestamp | None = None,
) -> DeltaWeekStatus:
    cache_week = pd.Timestamp(latest_cache_week) if latest_cache_week is not None else get_latest_cache_week()
    sql_week = pd.Timestamp(latest_sql_week) if latest_sql_week is not None else get_latest_sql_week()
    expected_next_week = _expected_next_week(cache_week)

    if expected_next_week is None or sql_week is None or expected_next_week > sql_week:
        return DeltaWeekStatus(sql_week, cache_week, expected_next_week, [])

    engine = get_connection()
    result = fetch_data_from_db(
        engine, build_distinct_weeks_since_query(expected_next_week.strftime("%Y-%m-%d"))
    )
    if result.empty:
        expected = pd.date_range(start=expected_next_week, end=sql_week, freq="7D")
        return DeltaWeekStatus(sql_week, cache_week, expected_next_week, list(expected))

    result["Week"] = pd.to_datetime(result["Week"], errors="coerce")
    weeks = pd.DatetimeIndex(sorted(result["Week"].dropna().tolist()))
    expected = pd.date_range(start=expected_next_week, end=sql_week, freq="7D")
    missing = list(expected.difference(weeks))
    return DeltaWeekStatus(sql_week, cache_week, expected_next_week, missing)


def refresh_manual_missing_weeks_cache(workbook_path: str | Path) -> pd.DataFrame:
    manual_df = pd.read_excel(workbook_path, sheet_name=0, header=5)
    week_columns = [column for column in manual_df.columns if hasattr(column, "strftime")]
    if not week_columns:
        raise ValueError("The Bookscan workbook did not contain any weekly date columns.")

    required = ["Pub", "PT", "CAT", "PGR", "ISBN 13", "TITLE", "PRICE", "SHIP"]
    missing = [column for column in required if column not in manual_df.columns]
    if missing:
        raise ValueError(f"The Bookscan workbook is missing required columns: {missing}")

    supplemental = manual_df.loc[:, required + week_columns].copy()
    supplemental = supplemental.rename(
        columns={
            "Pub": "Pub",
            "PT": "PT",
            "CAT": "CAT",
            "PGR": "pgrp",
            "ISBN 13": "ISBN",
            "TITLE": "Title",
            "PRICE": "Price",
            "SHIP": "PubDate",
        }
    )
    supplemental["ISBN"] = normalize_bookscan_isbn_series(supplemental["ISBN"].astype("string"))
    supplemental = supplemental[supplemental["ISBN"].notna()].copy()
    supplemental["Pub"] = supplemental["Pub"].map(_normalize_publisher)
    for column in ["Pub", "PT", "CAT", "pgrp", "Title"]:
        supplemental[column] = supplemental[column].where(supplemental[column].isna(), supplemental[column].astype(str))
    supplemental["Price"] = pd.to_numeric(supplemental["Price"], errors="coerce")
    supplemental["PubDate"] = pd.to_datetime(supplemental["PubDate"], errors="coerce")

    long_df = supplemental.melt(
        id_vars=["Pub", "PT", "CAT", "pgrp", "ISBN", "Title", "Price", "PubDate"],
        value_vars=week_columns,
        var_name="Week",
        value_name="qty",
    )
    long_df["Week"] = pd.to_datetime(long_df["Week"], errors="coerce")
    long_df["qty"] = pd.to_numeric(long_df["qty"], errors="coerce").fillna(0)
    long_df = long_df[(long_df["Week"].notna()) & (long_df["qty"] != 0)].copy()
    long_df["qty"] = long_df["qty"].astype(int)
    long_df = (
        long_df.groupby(["Week", "ISBN", "Pub", "PT", "CAT", "pgrp", "Title", "Price", "PubDate"], as_index=False)["qty"]
        .sum()
        .sort_values(["Week", "ISBN"])
        .reset_index(drop=True)
    )
    _save_parquet(long_df, manual_missing_weeks_file)
    return long_df


def _load_manual_missing_weeks() -> pd.DataFrame:
    if not manual_missing_weeks_file.exists():
        return pd.DataFrame()
    supplemental = pd.read_parquet(manual_missing_weeks_file)
    if supplemental.empty:
        return supplemental
    supplemental["Week"] = pd.to_datetime(supplemental["Week"])
    supplemental["ISBN"] = normalize_bookscan_isbn_series(supplemental["ISBN"].astype("string"))
    supplemental["PubDate"] = pd.to_datetime(supplemental["PubDate"], errors="coerce")
    supplemental["qty"] = pd.to_numeric(supplemental["qty"], errors="coerce").fillna(0).astype(int)
    if "Pub" in supplemental.columns:
        supplemental["Pub"] = supplemental["Pub"].map(_normalize_publisher)
    return supplemental[supplemental["ISBN"].notna()].copy()


def _apply_manual_missing_weeks(sales_df: pd.DataFrame) -> pd.DataFrame:
    supplemental = _load_manual_missing_weeks()
    if supplemental.empty:
        return sales_df

    renamed = supplemental.rename(columns={"Pub": "Publisher"}).copy()
    for column in ["PT", "CAT", "pgrp", "Title", "Price", "PubDate"]:
        if column not in renamed.columns:
            renamed[column] = None
    renamed["HistoryType"] = "weekly"

    combined = pd.concat([sales_df.copy(), renamed], ignore_index=True, sort=False)
    combined["Week"] = pd.to_datetime(combined["Week"])
    combined["ISBN"] = normalize_bookscan_isbn_series(combined["ISBN"].astype("string"))
    combined["qty"] = pd.to_numeric(combined["qty"], errors="coerce").fillna(0).astype(int)
    combined = combined.drop_duplicates(subset=["Week", "ISBN"], keep="last")
    combined = combined.sort_values(["Week", "ISBN"]).reset_index(drop=True)
    return combined


def refresh_sales_cache(
    full_refresh: bool = False,
    refresh_lookback_weeks: int = REFRESH_LOOKBACK_WEEKS,
) -> pd.DataFrame:
    refresh_lookback_weeks = max(int(refresh_lookback_weeks), 0)
    cached = _load_parquet_or_empty(sales_cache_file)
    if not cached.empty:
        if "HistoryType" not in cached.columns:
            full_refresh = True
        else:
            cached["HistoryType"] = cached["HistoryType"].astype(str)
            cached["Week"] = pd.to_datetime(cached["Week"])
            cached = cached[
                ~(
                    (cached["HistoryType"] == "weekly")
                    & (cached["Week"] < WEEKLY_HISTORY_START)
                )
            ].copy()
    start_date = "2023-01-01"
    if not full_refresh and not cached.empty:
        weekly_cached = cached[cached["HistoryType"] == "weekly"].copy()
        if not weekly_cached.empty:
            start_date = (
                weekly_cached["Week"].max() - pd.Timedelta(weeks=refresh_lookback_weeks)
            ).strftime("%Y-%m-%d")

    engine = get_connection()
    fetched = fetch_data_from_db(engine, build_sales_query(start_date))
    if fetched.empty:
        if cached.empty:
            raise ValueError("The Bookscan sales query returned no rows.")
        return cached

    fetched["Week"] = pd.to_datetime(fetched["Week"])
    fetched["ISBN"] = normalize_bookscan_isbn_series(fetched["RawISBN"].astype("string"))
    fetched = fetched[fetched["ISBN"].notna()].copy()
    fetched["HistoryType"] = fetched["HistoryType"].astype(str)
    fetched["qty"] = pd.to_numeric(fetched["qty"], errors="coerce").fillna(0).astype(int)
    fetched = fetched.groupby(["HistoryType", "Week", "ISBN"], as_index=False)["qty"].sum()

    if full_refresh or cached.empty:
        combined = fetched
    else:
        fetched_min_week = fetched.loc[fetched["HistoryType"] == "weekly", "Week"].min()
        fetched_max_week = fetched.loc[fetched["HistoryType"] == "weekly", "Week"].max()
        cached = cached[
            ~(
                (cached["HistoryType"] == "weekly")
                & cached["Week"].between(fetched_min_week, fetched_max_week)
            )
        ]
        cached = cached[cached["HistoryType"] != "yearly"]
        combined = pd.concat([cached, fetched], ignore_index=True)
    combined = combined.drop_duplicates(subset=["HistoryType", "Week", "ISBN"], keep="last")
    combined = combined.sort_values(["HistoryType", "Week", "ISBN"]).reset_index(drop=True)
    _save_parquet(combined, sales_cache_file)
    return combined


def refresh_inventory_cache(
    force: bool = False,
    inventory_detail_workbook_override: str | Path | None = None,
) -> pd.DataFrame:
    source_workbook = resolve_inventory_detail_workbook(inventory_detail_workbook_override)
    if (
        inventory_cache_file.exists()
        and not force
        and source_workbook.exists()
        and inventory_cache_file.stat().st_mtime >= source_workbook.stat().st_mtime
    ):
        return pd.read_parquet(inventory_cache_file)

    inventory = pd.read_excel(
        source_workbook,
        sheet_name="Inventory Detail",
        usecols=["ISBN", "Available To Sell", "Consignment", "Backorder", "Frozen"],
    )
    inventory["ISBN"] = normalize_bookscan_isbn_series(inventory["ISBN"].astype("string"))
    inventory = inventory[inventory["ISBN"].notna()].copy()
    inventory = inventory.rename(
        columns={
            "Available To Sell": "Available",
            "Consignment": "Consign",
            "Backorder": "B/O",
            "Frozen": "Frozen",
        }
    )
    for column in ["Available", "Consign", "B/O", "Frozen"]:
        inventory[column] = pd.to_numeric(inventory[column], errors="coerce").fillna(0)
    inventory = inventory.groupby("ISBN", as_index=False)[["Available", "Consign", "B/O", "Frozen"]].sum()
    inventory["SnapshotDate"] = pd.Timestamp(datetime.fromtimestamp(source_workbook.stat().st_mtime).date())
    _save_parquet(inventory, inventory_cache_file)
    return inventory


def _last_non_null(series: pd.Series):
    non_null = series.dropna()
    if non_null.empty:
        return None
    return non_null.iloc[-1]


def _build_manual_metadata() -> pd.DataFrame:
    supplemental = _load_manual_missing_weeks()
    if supplemental.empty:
        return pd.DataFrame(columns=["ISBN", "Pub", "PT", "CAT", "pgrp", "Title", "Price", "PubDate"])
    return (
        supplemental.sort_values(["Week", "ISBN"])
        .groupby("ISBN", as_index=False)
        .agg(
            {
                "Pub": _last_non_null,
                "PT": _last_non_null,
                "CAT": _last_non_null,
                "pgrp": _last_non_null,
                "Title": _last_non_null,
                "Price": _last_non_null,
                "PubDate": _last_non_null,
            }
        )
    )


def _fetch_item_metadata_for_isbns(isbns: list[str]) -> pd.DataFrame:
    if not isbns:
        return pd.DataFrame(columns=["ISBN", "Pub", "PT", "CAT", "pgrp", "Title", "Price", "PubDate"])

    frames: list[pd.DataFrame] = []
    try:
        engine = get_connection()
    except Exception:
        return pd.DataFrame(columns=["ISBN", "Pub", "PT", "CAT", "pgrp", "Title", "Price", "PubDate"])
    unique = sorted({isbn for isbn in isbns if isbn})
    for start in range(0, len(unique), METADATA_BATCH_SIZE):
        batch = unique[start : start + METADATA_BATCH_SIZE]
        quoted = ",".join(f"'{isbn}'" for isbn in batch)
        query = f"""
        SELECT
            i.ITEM_TITLE AS ISBN,
            CASE
                WHEN i.PUBLISHER_CODE = 'Quadrille Publishing Limited' THEN 'Quadrille'
                ELSE i.PUBLISHER_CODE
            END AS Pub,
            i.PRODUCT_TYPE AS PT,
            i.FORMAT AS CAT,
            CASE
                WHEN LEFT(i.PUBLISHING_GROUP, 3) = 'BAR' THEN 'BAR'
                ELSE i.PUBLISHING_GROUP
            END AS pgrp,
            i.SHORT_TITLE AS Title,
            i.PRICE_AMOUNT AS Price,
            CAST(i.AMORTIZATION_DATE AS date) AS PubDate
        FROM ebs.item i
        WHERE
            i.ITEM_TITLE IN ({quoted})
            AND i.PUBLISHER_CODE IS NOT NULL
            AND i.PUBLISHER_CODE NOT IN ({",".join(f"'{value}'" for value in PUBLISHER_EXCLUSIONS)})
            AND i.PRODUCT_TYPE IN ('BK', 'FT', 'RP', 'CP', 'DI')
            AND i.PUBLISHING_GROUP NOT IN ('MKT', 'ZZZ');
        """
        try:
            frame = fetch_data_from_db(engine, query)
        except Exception:
            break
        if not frame.empty:
            frames.append(frame)

    if not frames:
        return pd.DataFrame(columns=["ISBN", "Pub", "PT", "CAT", "pgrp", "Title", "Price", "PubDate"])

    metadata = pd.concat(frames, ignore_index=True)
    metadata["ISBN"] = normalize_bookscan_isbn_series(metadata["ISBN"].astype("string"))
    metadata["Pub"] = metadata["Pub"].map(_normalize_publisher)
    for column in ["Pub", "PT", "CAT", "pgrp", "Title"]:
        metadata[column] = metadata[column].where(metadata[column].isna(), metadata[column].astype(str))
    metadata["Price"] = pd.to_numeric(metadata["Price"], errors="coerce")
    metadata["PubDate"] = pd.to_datetime(metadata["PubDate"], errors="coerce")
    metadata = metadata[metadata["ISBN"].notna()].copy()
    metadata = _filter_allowed_publishers(metadata)
    metadata = metadata.drop_duplicates(subset=["ISBN"], keep="first").reset_index(drop=True)
    return metadata


def _fetch_source_metadata() -> pd.DataFrame:
    engine = get_connection()
    metadata = fetch_data_from_db(engine, SOURCE_METADATA_QUERY)
    if metadata.empty:
        return metadata
    metadata["ISBN"] = normalize_bookscan_isbn_series(metadata["ISBN"].astype("string"))
    metadata["Pub"] = metadata["Pub"].map(_normalize_publisher)
    for column in ["Pub", "PT", "CAT", "pgrp", "Title"]:
        metadata[column] = metadata[column].where(metadata[column].isna(), metadata[column].astype(str))
    metadata["Price"] = pd.to_numeric(metadata["Price"], errors="coerce")
    metadata["PubDate"] = pd.to_datetime(metadata["PubDate"], errors="coerce")
    metadata = metadata[metadata["ISBN"].notna()].copy()
    metadata = _filter_allowed_publishers(metadata)
    metadata = metadata.drop_duplicates(subset=["ISBN"], keep="first").reset_index(drop=True)
    return metadata


def refresh_metadata_cache(
    sales_df: pd.DataFrame,
    inventory_df: pd.DataFrame,
    force: bool = False,
    include_manual_missing_weeks: bool = False,
) -> pd.DataFrame:
    cached = pd.DataFrame() if force else _load_parquet_or_empty(metadata_cache_file)
    if force or cached.empty:
        cached = _fetch_source_metadata()
    manual = _build_manual_metadata() if include_manual_missing_weeks else pd.DataFrame()
    if not cached.empty:
        cached["ISBN"] = normalize_bookscan_isbn_series(cached["ISBN"].astype("string"))
        cached["PubDate"] = pd.to_datetime(cached["PubDate"], errors="coerce")
        cached = _filter_allowed_publishers(cached)

    target_isbns = set(sales_df["ISBN"].dropna().astype(str).tolist())
    target_isbns.update(inventory_df["ISBN"].dropna().astype(str).tolist())
    if "ISBN" in manual.columns:
        target_isbns.update(manual["ISBN"].dropna().astype(str).tolist())
    cached_isbns = set(cached["ISBN"].dropna().astype(str).tolist()) if not cached.empty else set()
    fetched = _fetch_item_metadata_for_isbns(sorted(target_isbns - cached_isbns))

    frames = [frame for frame in [cached, fetched, manual] if not frame.empty]
    if not frames:
        return pd.DataFrame(columns=["ISBN", "Pub", "PT", "CAT", "pgrp", "Title", "Price", "PubDate"])

    combined = pd.concat(frames, ignore_index=True, sort=False)
    combined["ISBN"] = normalize_bookscan_isbn_series(combined["ISBN"].astype("string"))
    combined["Pub"] = combined["Pub"].map(_normalize_publisher)
    for column in ["Pub", "PT", "CAT", "pgrp", "Title"]:
        combined[column] = combined[column].where(combined[column].isna(), combined[column].astype(str))
    combined["PubDate"] = pd.to_datetime(combined["PubDate"], errors="coerce")
    combined = combined[combined["ISBN"].notna()].copy()
    combined = _filter_allowed_publishers(combined)
    combined = combined.drop_duplicates(subset=["ISBN"], keep="last").sort_values("ISBN").reset_index(drop=True)
    _save_parquet(combined, metadata_cache_file)
    return combined


def _series_by_isbn(sales_df: pd.DataFrame, mask: pd.Series) -> pd.Series:
    if not mask.any():
        return pd.Series(dtype="float64")
    return sales_df.loc[mask].groupby("ISBN")["qty"].sum()


def _history_columns(df: pd.DataFrame) -> list[str]:
    return [
        column
        for column in df.columns
        if isinstance(column, str)
        and ((len(column) == 10 and column.count("-") == 2) or column.startswith("12-31-"))
    ]


def _prune_zero_history_columns(df: pd.DataFrame) -> pd.DataFrame:
    history_columns = _history_columns(df)
    keep_history = [
        column for column in history_columns if pd.to_numeric(df[column], errors="coerce").fillna(0).sum() != 0
    ]
    keep_columns = [column for column in df.columns if column not in history_columns or column in keep_history]
    return df.loc[:, keep_columns].copy()


def build_report_dataframe(
    sales_df: pd.DataFrame,
    inventory_df: pd.DataFrame,
    metadata_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.Timestamp]:
    if sales_df.empty:
        raise ValueError("Bookscan sales cache is empty.")
    if metadata_df.empty:
        raise ValueError("Bookscan metadata cache is empty.")

    sales_df = sales_df.copy()
    sales_df["Week"] = pd.to_datetime(sales_df["Week"])
    sales_df["ISBN"] = sales_df["ISBN"].astype("string").str.strip()
    sales_df["HistoryType"] = sales_df.get("HistoryType", "weekly").astype(str)
    sales_df["qty"] = pd.to_numeric(sales_df["qty"], errors="coerce").fillna(0).astype(int)
    sales_df = sales_df[sales_df["ISBN"].notna()].copy()

    metadata_df = metadata_df.copy()
    metadata_df["ISBN"] = metadata_df["ISBN"].astype("string").str.strip()
    metadata_df = metadata_df[metadata_df["ISBN"].notna()].copy()
    allowed_isbns = set(metadata_df["ISBN"].astype(str))
    if not allowed_isbns:
        raise ValueError("Bookscan metadata contains no allowed ISBNs.")

    # Only keep titles that resolve to an allowed ebs.item record.
    sales_df = sales_df[sales_df["ISBN"].astype(str).isin(allowed_isbns)].copy()

    inventory_df = inventory_df.copy()
    inventory_df["ISBN"] = inventory_df["ISBN"].astype("string").str.strip()
    inventory_df = inventory_df[inventory_df["ISBN"].notna()].copy()
    inventory_df = inventory_df[inventory_df["ISBN"].astype(str).isin(allowed_isbns)].copy()

    weekly_sales = sales_df[sales_df["HistoryType"] == "weekly"].copy()
    yearly_sales = sales_df[sales_df["HistoryType"] == "yearly"].copy()
    if weekly_sales.empty:
        raise ValueError("No Bookscan weekly sales remain after filtering to allowed ebs.item ISBNs.")
    latest_week = weekly_sales["Week"].max()
    base_index = pd.Index(
        sorted(
            set(weekly_sales["ISBN"].astype(str))
            | set(yearly_sales["ISBN"].astype(str))
            | set(inventory_df["ISBN"].dropna().astype(str))
        )
    )

    qty_by_week = (
        weekly_sales.pivot_table(index="ISBN", columns="Week", values="qty", aggfunc="sum", fill_value=0)
        .sort_index(axis=1, ascending=False)
    )
    qty_by_week.columns = pd.to_datetime(qty_by_week.columns)

    weekly_columns = [week for week in qty_by_week.columns if week >= WEEKLY_HISTORY_START]
    weekly_df = qty_by_week.loc[:, weekly_columns].copy() if weekly_columns else qty_by_week.iloc[:, 0:0].copy()
    weekly_df.columns = [week.strftime("%m-%d-%Y") for week in weekly_columns]
    weekly_df = weekly_df.reindex(base_index, fill_value=0)

    yearly_df = pd.DataFrame(index=base_index)
    if not yearly_sales.empty:
        yearly_wide = (
            yearly_sales.pivot_table(index="ISBN", columns="Week", values="qty", aggfunc="sum", fill_value=0)
            .sort_index(axis=1, ascending=False)
        )
        yearly_wide.columns = [pd.Timestamp(column).strftime("%m-%d-%Y") for column in yearly_wide.columns]
        keep_yearly_columns = [f"12-31-{year}" for year in range(YEARLY_HISTORY_END, YEARLY_HISTORY_START - 1, -1)]
        yearly_df = yearly_wide.reindex(index=base_index, columns=keep_yearly_columns, fill_value=0)

    latest_bookscan = bookscan_week(latest_week)
    bookscan_dates = bookscan_parts(weekly_sales["Week"])
    tytd = _series_by_isbn(weekly_sales, bookscan_dates["BookScanYear"] == latest_bookscan.year)
    lytd = _series_by_isbn(
        weekly_sales,
        (bookscan_dates["BookScanYear"] == latest_bookscan.year - 1)
        & (bookscan_dates["BookScanWeek"] <= latest_bookscan.week),
    )
    ly_fy = _series_by_isbn(weekly_sales, bookscan_dates["BookScanYear"] == latest_bookscan.year - 1)
    ltd = sales_df.groupby("ISBN")["qty"].sum()
    w52 = _series_by_isbn(weekly_sales, weekly_sales["Week"].between(latest_week - pd.Timedelta(weeks=51), latest_week))
    last6 = _series_by_isbn(weekly_sales, weekly_sales["Week"].between(latest_week - pd.Timedelta(weeks=5), latest_week))
    last26 = _series_by_isbn(weekly_sales, weekly_sales["Week"].between(latest_week - pd.Timedelta(weeks=25), latest_week))

    metrics = pd.concat(
        [
            w52.rename("52 WK"),
            (last6 / 6.0).round(2).rename("6Wk Avg"),
            tytd.rename("TYTD"),
            lytd.rename("LYTD"),
            tytd.subtract(lytd, fill_value=0).rename("YTD Var"),
            ly_fy.rename("LY_FY"),
            ltd.rename("LTD"),
            (last26 / 26.0).round(2).rename("26Wk Avg"),
        ],
        axis=1,
    ).fillna(0).reindex(base_index, fill_value=0)

    inventory = inventory_df.drop_duplicates(subset=["ISBN"], keep="last").set_index("ISBN")
    inventory = inventory[["Available", "Consign", "B/O", "Frozen"]].reindex(base_index)

    metadata = metadata_df.copy()
    metadata["PubDate"] = pd.to_datetime(metadata["PubDate"], errors="coerce").dt.strftime("%m-%d-%Y")
    metadata["PubDate"] = metadata["PubDate"].fillna("")
    metadata = metadata.drop_duplicates(subset=["ISBN"], keep="last").set_index("ISBN")
    metadata = metadata.reindex(base_index)

    report = pd.concat([metadata, inventory, metrics, weekly_df, yearly_df], axis=1)
    report["Avg Avail Wk"] = (
        pd.to_numeric(report["Available"], errors="coerce").fillna(0)
        / pd.to_numeric(report["26Wk Avg"], errors="coerce").replace(0, pd.NA)
    )
    report["Avg Avail Wk"] = pd.to_numeric(report["Avg Avail Wk"], errors="coerce").fillna(0).round(2)
    report = report.reset_index().rename(columns={"index": "ISBN"})

    ordered_columns = [
        "Pub",
        "PT",
        "CAT",
        "pgrp",
        "ISBN",
        "Title",
        "Price",
        "PubDate",
        "Available",
        "Consign",
        "B/O",
        "Frozen",
        "Avg Avail Wk",
        "52 WK",
        "6Wk Avg",
        "TYTD",
        "LYTD",
        "YTD Var",
        "LY_FY",
        "LTD",
    ]
    history_columns = [column for column in report.columns if column not in ordered_columns]
    report = report[ordered_columns + history_columns].copy()
    report = report.drop(columns=["26Wk Avg"], errors="ignore")

    for column in ["Pub", "PT", "CAT", "pgrp", "Title", "PubDate"]:
        if column in report.columns:
            report[column] = report[column].fillna("")
    if "Price" in report.columns:
        report["Price"] = pd.to_numeric(report["Price"], errors="coerce")
    for column in ["Available", "Consign", "B/O", "Frozen", "52 WK", "TYTD", "LYTD", "YTD Var", "LY_FY", "LTD"]:
        if column in report.columns:
            report[column] = pd.to_numeric(report[column], errors="coerce").fillna(0).astype(int)
    if "6Wk Avg" in report.columns:
        report["6Wk Avg"] = pd.to_numeric(report["6Wk Avg"], errors="coerce").fillna(0).round(2)

    latest_week_label = latest_week.strftime("%m-%d-%Y")
    report = report.sort_values([latest_week_label, "Pub", "pgrp", "Title", "ISBN"], ascending=[False, True, True, True, True])
    report = report.reset_index(drop=True)
    return report, latest_week


def _build_save_options(report_df: pd.DataFrame, latest_week: pd.Timestamp) -> dict[str, object]:
    date_cols = [column for column in report_df.columns if isinstance(column, str) and len(column) == 10 and column.count("-") == 2]
    yearly_cols = [column for column in report_df.columns if isinstance(column, str) and column.startswith("12-31-")]
    summary_cols = ["Available", "Consign", "B/O", "Frozen", "52 WK", "TYTD", "LYTD", "YTD Var", "LY_FY", "LTD"]
    format_cols = date_cols + yearly_cols + summary_cols
    decimal_cols = ["Price", "Avg Avail Wk", "6Wk Avg"]
    summary_decimal_cols = ["Avg Avail Wk", "6Wk Avg"]
    totals = build_column_totals(report_df, format_cols + summary_decimal_cols)
    header_fill_overrides = {col_idx: "#E6B8B7" for col_idx in range(8, 13)}
    history_top_labels: dict[int, str] = {}
    for col_idx, column in enumerate(report_df.columns):
        if isinstance(column, str) and column.startswith("12-31-"):
            history_top_labels[col_idx] = f"FY {column[-4:]}"
    accounting_format = {"num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'}
    decimal_accounting_format = {"num_format": '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'}
    column_format_overrides = {
        col_idx: {"format": accounting_format}
        for col_idx in range(8, min(20, len(report_df.columns)))
        if col_idx != 12
    }
    if 12 < len(report_df.columns):
        column_format_overrides[12] = {"format": decimal_accounting_format, "width": 10}
    if 14 < len(report_df.columns):
        column_format_overrides[14] = {"format": decimal_accounting_format, "width": 10}
    title_block = {
        "start_row": 1,
        "end_row": 2,
        "start_col": 5,
        "end_col": 5,
        "title": "Rolling Bookscan POS",
        "subtitle": f"Week Ending: {latest_week.strftime('%B %d, %Y')}",
        "merge_cells": False,
        "align": "center",
    }
    return {
        "summary": totals,
        "format_cols": format_cols,
        "decimal_cols": decimal_cols,
        "integer_accounting_no_symbol": True,
        "rolling_main_layout": True,
        "pre_date_column_count": 20,
        "summary_label_col_idx": 7,
        "header_fill_overrides": header_fill_overrides,
        "format_blank_summary_cells": False,
        "title_block": title_block,
        "header_row_override": 5,
        "show_weeknum_label": True,
        "history_top_labels": history_top_labels,
        "weeknum_label_fill": "#C4BD97",
        "column_format_overrides": column_format_overrides,
    }


def save_reports_by_pub(report_df: pd.DataFrame, latest_week: pd.Timestamp) -> int:
    filename = _format_bookscan_output_filename(latest_week)
    saved_count = 0
    for publisher, folder in bookscan_dp_folders.items():
        df_pub = report_df[report_df["Pub"] == publisher].copy()
        if df_pub.empty:
            print(f"No data for {publisher}")
            continue
        df_pub = _prune_zero_history_columns(df_pub)
        folder.mkdir(parents=True, exist_ok=True)
        output_file = folder / filename
        save_to_excel(df_pub, output_file, **_build_save_options(df_pub, latest_week))
        print(f"Saved Bookscan rolling report for {publisher} to {output_file}")
        saved_count += 1
    return saved_count


def save_dp_reports_from_main_workbook(source_file: str | Path | None = None) -> DpSaveResult:
    report_df, latest_week, workbook_path = load_saved_main_report(source_file)
    dp_files_saved = save_reports_by_pub(report_df, latest_week)
    return DpSaveResult(
        source_file=workbook_path,
        latest_week=latest_week,
        dp_files_saved=dp_files_saved,
    )


def build_customer_sales_report(
    refresh_sales: bool = False,
    refresh_inventory: bool = False,
    refresh_manual_cache: bool = False,
    full_refresh: bool = False,
    refresh_lookback_weeks: int = REFRESH_LOOKBACK_WEEKS,
    inventory_detail_workbook_override: str | Path | None = None,
    save_dp: bool = False,
    local_only: bool = False,
    manual_missing_weeks_workbook: str | Path | None = None,
    include_manual_missing_weeks: bool = False,
) -> RollingBuildResult:
    history_workbook = Path(manual_missing_weeks_workbook) if manual_missing_weeks_workbook else find_default_history_workbook()
    if include_manual_missing_weeks and history_workbook and (
        refresh_manual_cache or full_refresh or not manual_missing_weeks_file.exists()
    ):
        refresh_manual_missing_weeks_cache(history_workbook)

    sales_df = (
        refresh_sales_cache(
            full_refresh=full_refresh,
            refresh_lookback_weeks=refresh_lookback_weeks,
        )
        if refresh_sales or full_refresh or not sales_cache_file.exists()
        else _load_parquet_or_empty(sales_cache_file)
    )
    if include_manual_missing_weeks:
        sales_df = _apply_manual_missing_weeks(sales_df)
    inventory_df = (
        refresh_inventory_cache(
            force=(refresh_inventory or full_refresh),
            inventory_detail_workbook_override=inventory_detail_workbook_override,
        )
        if refresh_inventory or full_refresh or not inventory_cache_file.exists()
        else _load_parquet_or_empty(inventory_cache_file)
    )
    metadata_df = refresh_metadata_cache(
        sales_df,
        inventory_df,
        force=full_refresh,
        include_manual_missing_weeks=include_manual_missing_weeks,
    )

    report_df, latest_week = build_report_dataframe(sales_df, inventory_df, metadata_df)
    output_dir = local_review_dir if local_only else bookscan_rolling_folder
    output_file = output_dir / _format_bookscan_output_filename(latest_week)
    output_dir.mkdir(parents=True, exist_ok=True)

    save_to_excel(report_df, output_file, **_build_save_options(report_df, latest_week))
    dp_files_saved = save_reports_by_pub(report_df, latest_week) if save_dp and not local_only else 0
    return RollingBuildResult(
        output_file=output_file,
        latest_week=latest_week,
        sales_rows=len(sales_df),
        inventory_rows=len(inventory_df),
        report_shape=report_df.shape,
        dp_files_saved=dp_files_saved,
    )


def refresh_caches_only(
    refresh_manual_cache: bool = False,
    full_refresh: bool = False,
    refresh_lookback_weeks: int = REFRESH_LOOKBACK_WEEKS,
    inventory_detail_workbook_override: str | Path | None = None,
    manual_missing_weeks_workbook: str | Path | None = None,
    include_manual_missing_weeks: bool = False,
) -> CacheRefreshResult:
    history_workbook = Path(manual_missing_weeks_workbook) if manual_missing_weeks_workbook else find_default_history_workbook()
    if include_manual_missing_weeks and history_workbook and (
        refresh_manual_cache or full_refresh or not manual_missing_weeks_file.exists()
    ):
        refresh_manual_missing_weeks_cache(history_workbook)

    prior_cache_week = None if full_refresh else get_latest_cache_week()
    latest_sql_week = get_latest_sql_week()
    delta_status = get_delta_week_status(
        latest_cache_week=prior_cache_week,
        latest_sql_week=latest_sql_week,
    )
    sales_df = refresh_sales_cache(
        full_refresh=full_refresh,
        refresh_lookback_weeks=refresh_lookback_weeks,
    )
    if include_manual_missing_weeks:
        sales_df = _apply_manual_missing_weeks(sales_df)
    inventory_df = refresh_inventory_cache(
        force=full_refresh,
        inventory_detail_workbook_override=inventory_detail_workbook_override,
    )
    refresh_metadata_cache(
        sales_df,
        inventory_df,
        force=full_refresh,
        include_manual_missing_weeks=include_manual_missing_weeks,
    )
    sales_cache_week = pd.to_datetime(sales_df["Week"]).max() if not sales_df.empty else None
    inventory_cache_week = pd.to_datetime(inventory_df["SnapshotDate"]).max() if not inventory_df.empty else None
    return CacheRefreshResult(
        refresh_mode="full" if full_refresh else "incremental",
        latest_sql_week=latest_sql_week,
        sales_cache_week=sales_cache_week,
        inventory_cache_week=inventory_cache_week,
        expected_next_week=delta_status.expected_next_week,
        missing_week_count=len(delta_status.missing_weeks),
        missing_weeks=delta_status.missing_weeks,
        sales_rows=len(sales_df),
        inventory_rows=len(inventory_df),
    )


def print_week_check(result: WeekCheckResult) -> None:
    print()
    print(
        "Source range: "
        f"{result.min_week.strftime('%Y-%m-%d') if result.min_week is not None else 'None'}"
        " -> "
        f"{result.max_week.strftime('%Y-%m-%d') if result.max_week is not None else 'None'}"
    )
    print(f"Missing weeks: {len(result.missing_weeks)}")
    if result.missing_weeks:
        full_list = ", ".join(week.strftime("%Y-%m-%d") for week in reversed(result.missing_weeks))
        print(f"Missing week list: {full_list}")
    if result.latest_week_rows:
        print("Latest weeks present:")
        filename_width = max(
            len("FILENAME"),
            *(len(str(row.get("Filename", ""))) for row in result.latest_week_rows),
        )
        print(f"    {'WEEK':<10}  {'FILENAME':<{filename_width}}  {'ROW_CNT':>12}  {'SALES_QTY':>12}")
        for row in result.latest_week_rows:
            week = pd.to_datetime(row.get("Week"), errors="coerce")
            week_text = week.strftime("%Y-%m-%d") if pd.notna(week) else ""
            filename = str(row.get("Filename", ""))
            row_count = pd.to_numeric(row.get("Row_Cnt"), errors="coerce")
            sales_qty = pd.to_numeric(row.get("Sales_Qty"), errors="coerce")
            row_count_text = f"{row_count:,.0f}" if pd.notna(row_count) else ""
            sales_qty_text = f"{sales_qty:,.0f}" if pd.notna(sales_qty) else ""
            print(
                f"    {week_text:<10}  {filename:<{filename_width}}  "
                f"{row_count_text:>12}  {sales_qty_text:>12}"
            )


def print_result_summary(result: RollingBuildResult) -> None:
    print()
    print(f"Latest week: {result.latest_week:%Y-%m-%d}")
    print(f"Sales cache rows: {result.sales_rows:,}")
    print(f"Inventory cache rows: {result.inventory_rows:,}")
    print(f"Report shape: {result.report_shape}")
    if result.dp_files_saved:
        print(f"DP files saved: {result.dp_files_saved}")


def print_dp_save_summary(result: DpSaveResult) -> None:
    print()
    print(f"Source main report: {result.source_file}")
    print(f"Latest week: {result.latest_week:%Y-%m-%d}")
    print(f"DP files saved: {result.dp_files_saved}")


def print_cache_refresh_summary(result: CacheRefreshResult) -> None:
    print()
    print(
        "Refresh mode: "
        + ("Full historical rebuild" if result.refresh_mode == "full" else "Latest changes only")
    )
    print(
        "Latest SQL week: "
        f"{result.latest_sql_week.strftime('%Y-%m-%d') if result.latest_sql_week is not None else 'None'}"
    )
    print(
        "Sales cache latest week: "
        f"{result.sales_cache_week.strftime('%Y-%m-%d') if result.sales_cache_week is not None else 'None'}"
    )
    print(
        "Inventory snapshot date: "
        f"{result.inventory_cache_week.strftime('%Y-%m-%d') if result.inventory_cache_week is not None else 'None'}"
    )
    print(
        "Expected next week: "
        f"{result.expected_next_week.strftime('%Y-%m-%d') if result.expected_next_week is not None else 'None'}"
    )
    print(f"Missing SQL weeks detected: {result.missing_week_count}")
    if result.missing_weeks:
        preview = ", ".join(week.strftime("%Y-%m-%d") for week in result.missing_weeks)
        print(f"Missing SQL week list: {preview}")
    print(f"Sales cache rows: {result.sales_rows:,}")
    print(f"Inventory cache rows: {result.inventory_rows:,}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build the Bookscan rolling report.")
    parser.add_argument("--refresh-sales-cache", action="store_true", help="Refresh the sales cache from SQL before building.")
    parser.add_argument("--refresh-inventory-cache", action="store_true", help="Refresh the inventory detail cache before building.")
    parser.add_argument("--refresh-manual-cache", action="store_true", help="Refresh the missing-week supplement cache from the local historical workbook.")
    parser.add_argument(
        "--include-manual-missing-weeks",
        action="store_true",
        help="Include the legacy manual missing-week supplement. Normal builds skip it when SQL has complete weekly coverage.",
    )
    parser.add_argument("--full-refresh", action="store_true", help="Rebuild all Bookscan caches from scratch before building.")
    parser.add_argument(
        "--refresh-lookback-weeks",
        type=int,
        default=REFRESH_LOOKBACK_WEEKS,
        help="How many recent weekly Bookscan weeks to re-pull from SQL when refreshing the sales cache.",
    )
    parser.add_argument(
        "--inventory-detail-workbook",
        help="Optional full path to the current Inventory Detail workbook.",
    )
    parser.add_argument("--save-dp", action="store_true", help="Also save publisher-filtered versions to the DP folders.")
    parser.add_argument("--local-only", action="store_true", help="Save the report to the local review_output folder instead of the network target.")
    parser.add_argument("--manual-missing-weeks-workbook", help="Optional Bookscan workbook used to seed the missing-week supplement cache.")
    parser.add_argument("--check-weeks", action="store_true", help="Check the Bookscan SQL source for missing weeks and exit.")
    parser.add_argument("--refresh-caches-only", action="store_true", help="Refresh the caches and exit without building the report.")
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.check_weeks:
        print_week_check(check_source_weeks())
        return

    if args.refresh_caches_only:
        result = refresh_caches_only(
            refresh_manual_cache=args.refresh_manual_cache,
            full_refresh=args.full_refresh,
            refresh_lookback_weeks=args.refresh_lookback_weeks,
            inventory_detail_workbook_override=args.inventory_detail_workbook,
            manual_missing_weeks_workbook=args.manual_missing_weeks_workbook,
            include_manual_missing_weeks=args.include_manual_missing_weeks,
        )
        print_cache_refresh_summary(result)
        return

    result = build_customer_sales_report(
        refresh_sales=args.refresh_sales_cache,
        refresh_inventory=args.refresh_inventory_cache,
        refresh_manual_cache=args.refresh_manual_cache,
        full_refresh=args.full_refresh,
        refresh_lookback_weeks=args.refresh_lookback_weeks,
        inventory_detail_workbook_override=args.inventory_detail_workbook,
        save_dp=args.save_dp,
        local_only=args.local_only,
        manual_missing_weeks_workbook=args.manual_missing_weeks_workbook,
        include_manual_missing_weeks=args.include_manual_missing_weeks,
    )
    print_result_summary(result)


if __name__ == "__main__":
    main()
