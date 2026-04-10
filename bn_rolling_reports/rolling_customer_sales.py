from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.append(str(Path(__file__).resolve().parents[1]))

from amazon_rolling_reports.functions import build_column_totals, save_to_excel
from inventory_working import build_inventory_working_file
from isbn_utils import normalize_isbn_series
from pos_combiner import (
    build_combined_pos,
    format_output_filename as format_pos_output_filename,
    get_candidate_raw_folders,
    parse_week_ending,
    resolve_raw_folder,
)
from rolling_paths import (
    bn_dp_folders,
    bn_rolling_folder,
    inventory_cache_file,
    local_review_dir,
    manual_missing_weeks_file,
    sales_cache_file,
)
from rolling_queries import build_customer_sales_query
from shared.db.connection import get_connection
from shared.db.query_runner import fetch_data_from_db


WEEKLY_HISTORY_START = pd.Timestamp("2023-01-01")
YEARLY_HISTORY_START = 2007
YEARLY_HISTORY_END = 2022
MANUAL_MISSING_WEEKS = (
    pd.Timestamp("2017-12-09"),
    pd.Timestamp("2018-03-17"),
    pd.Timestamp("2018-03-31"),
    pd.Timestamp("2023-01-21"),
    pd.Timestamp("2023-01-28"),
    pd.Timestamp("2024-03-30"),
    pd.Timestamp("2025-07-05"),
    pd.Timestamp("2026-01-03"),
)


PUBLISHER_NORMALIZATION = {
    "Quadrille Publishing Limited": "Quadrille",
}


@dataclass
class RollingBuildResult:
    output_file: Path
    latest_week: pd.Timestamp
    sales_rows: int
    inventory_rows: int
    report_shape: tuple[int, int]
    dp_files_saved: int = 0


@dataclass
class CacheRefreshResult:
    latest_sql_week: pd.Timestamp | None
    sales_cache_week: pd.Timestamp | None
    inventory_cache_week: pd.Timestamp | None
    sales_rows: int
    inventory_rows: int


def _ensure_cache_dir() -> None:
    sales_cache_file.parent.mkdir(parents=True, exist_ok=True)


def _normalize_publisher_value(value):
    if pd.isna(value):
        return value
    return PUBLISHER_NORMALIZATION.get(str(value), value)


def _load_parquet_or_empty(cache_file: Path) -> pd.DataFrame:
    if cache_file.exists():
        return pd.read_parquet(cache_file)
    return pd.DataFrame()


def _save_parquet(df: pd.DataFrame, cache_file: Path) -> None:
    _ensure_cache_dir()
    df.to_parquet(cache_file, index=False)


def _format_bn_output_filename(week_ending: pd.Timestamp) -> str:
    return (
        f"Week {week_ending.isocalendar().week:02d} - {week_ending:%Y} "
        f"Rolling Barnes & Noble ({week_ending:%m%d%y}).xlsx"
    )


def _manual_week_column_map(columns) -> dict[pd.Timestamp, object]:
    week_map: dict[pd.Timestamp, object] = {}
    target_weeks = set(MANUAL_MISSING_WEEKS)
    for column in columns:
        timestamp = None
        if hasattr(column, "strftime"):
            timestamp = pd.Timestamp(column)
        elif isinstance(column, str) and len(column) == 10 and column.count("-") == 2:
            try:
                timestamp = pd.to_datetime(column, format="%m-%d-%Y")
            except ValueError:
                timestamp = None
        if timestamp is not None and timestamp in target_weeks:
            week_map[timestamp] = column
    return week_map


def refresh_manual_missing_weeks_cache(workbook_path: str | Path) -> pd.DataFrame:
    manual_df = pd.read_excel(workbook_path, sheet_name=0, header=5)
    week_column_map = _manual_week_column_map(manual_df.columns)
    if not week_column_map:
        raise ValueError("The manual workbook did not contain any of the configured missing week columns.")

    required_columns = ["PUB", "PT", "CAT", "PGR", "ISBN 13", "TITLE", "PRICE", "SHIP", "Dept", "Cat"]
    missing = [column for column in required_columns if column not in manual_df.columns]
    if missing:
        raise ValueError(f"The manual workbook is missing required columns: {missing}")

    supplemental = manual_df.loc[:, required_columns + list(week_column_map.values())].copy()
    supplemental = supplemental.rename(
        columns={
            "PUB": "Publisher",
            "PT": "PT",
            "CAT": "CAT",
            "PGR": "pgrp",
            "ISBN 13": "ISBN",
            "TITLE": "Title",
            "PRICE": "Price",
            "SHIP": "PubDate",
            "Dept": "DeptCode",
            "Cat": "SubjectCode",
        }
    )
    supplemental["Publisher"] = supplemental["Publisher"].map(_normalize_publisher_value)
    supplemental["ISBN"] = normalize_isbn_series(supplemental["ISBN"].astype("string"))
    supplemental = supplemental[supplemental["ISBN"].notna()].copy()
    supplemental["Price"] = pd.to_numeric(supplemental["Price"], errors="coerce")
    supplemental["PubDate"] = pd.to_datetime(supplemental["PubDate"], errors="coerce")
    supplemental["DeptCode"] = (
        pd.to_numeric(supplemental["DeptCode"], errors="coerce")
        .fillna(0)
        .astype(int)
        .astype(str)
        .str.zfill(5)
    )
    supplemental["SubjectCode"] = (
        pd.to_numeric(supplemental["SubjectCode"], errors="coerce")
        .fillna(0)
        .astype(int)
        .astype(str)
        .str.zfill(5)
    )

    rename_weeks = {column: week.strftime("%m-%d-%Y") for week, column in week_column_map.items()}
    supplemental = supplemental.rename(columns=rename_weeks)
    long_df = supplemental.melt(
        id_vars=["ISBN", "Title", "Publisher", "PT", "CAT", "pgrp", "SubjectCode", "DeptCode", "Price", "PubDate"],
        value_vars=list(rename_weeks.values()),
        var_name="Week",
        value_name="qty",
    )
    long_df["Week"] = pd.to_datetime(long_df["Week"], format="%m-%d-%Y")
    long_df["qty"] = pd.to_numeric(long_df["qty"], errors="coerce").fillna(0)
    long_df = long_df[long_df["qty"] != 0].copy()
    long_df = (
        long_df.groupby(
            ["Week", "ISBN", "Title", "Publisher", "PT", "CAT", "pgrp", "SubjectCode", "DeptCode", "Price", "PubDate"],
            as_index=False,
        )["qty"]
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
    if "Publisher" in supplemental.columns:
        supplemental["Publisher"] = supplemental["Publisher"].map(_normalize_publisher_value)
    supplemental["Week"] = pd.to_datetime(supplemental["Week"])
    supplemental["PubDate"] = pd.to_datetime(supplemental["PubDate"], errors="coerce")
    supplemental["ISBN"] = normalize_isbn_series(supplemental["ISBN"].astype("string"))
    supplemental["qty"] = pd.to_numeric(supplemental["qty"], errors="coerce").fillna(0).astype(int)
    return supplemental[supplemental["ISBN"].notna()].copy()


def _apply_manual_missing_weeks(sales_df: pd.DataFrame) -> pd.DataFrame:
    supplemental = _load_manual_missing_weeks()
    if supplemental.empty:
        return sales_df

    combined = pd.concat([sales_df.copy(), supplemental], ignore_index=True, sort=False)
    combined["Week"] = pd.to_datetime(combined["Week"])
    combined["PubDate"] = pd.to_datetime(combined["PubDate"], errors="coerce")
    combined["ISBN"] = normalize_isbn_series(combined["ISBN"].astype("string"))
    combined["qty"] = pd.to_numeric(combined["qty"], errors="coerce").fillna(0).astype(int)
    combined = combined.drop_duplicates(
        subset=["Week", "ISBN", "Title", "Publisher", "pgrp", "SubjectCode", "DeptCode"],
        keep="last",
    ).sort_values(["Week", "ISBN"]).reset_index(drop=True)
    return combined


def refresh_sales_cache(full_refresh: bool = False) -> pd.DataFrame:
    cached = _load_parquet_or_empty(sales_cache_file)
    start_date = "2007-01-01"
    if not full_refresh and not cached.empty:
        cached["Week"] = pd.to_datetime(cached["Week"])
        start_date = (cached["Week"].max() + pd.Timedelta(days=1)).strftime("%Y-%m-%d")

    engine = get_connection()
    query = build_customer_sales_query(start_date)
    fetched = fetch_data_from_db(engine, query)

    if fetched.empty:
        if cached.empty:
            raise ValueError("The B&N customer sales query returned no rows.")
        return cached

    fetched["Week"] = pd.to_datetime(fetched["Week"])
    fetched["PubDate"] = pd.to_datetime(fetched["PubDate"], errors="coerce")
    fetched["ISBN"] = normalize_isbn_series(fetched["ISBN"].astype("string"))
    fetched = fetched[fetched["ISBN"].notna()].copy()
    fetched["qty"] = pd.to_numeric(fetched["qty"], errors="coerce").fillna(0).astype(int)
    fetched["Price"] = pd.to_numeric(fetched["Price"], errors="coerce")

    combined = fetched if full_refresh or cached.empty else pd.concat([cached, fetched], ignore_index=True)
    combined = combined.drop_duplicates(
        subset=["Week", "ISBN", "Title", "Publisher", "pgrp", "SubjectCode", "DeptCode"],
        keep="last",
    ).sort_values(["Week", "ISBN"]).reset_index(drop=True)
    _save_parquet(combined, sales_cache_file)
    return combined


def get_latest_sql_week() -> pd.Timestamp | None:
    engine = get_connection()
    query = """
    SELECT MAX(CAST([WEEK] AS date)) AS latest_week
    FROM [CBQ2].[cb].[Sellthrough_Barnes_and_Noble];
    """
    result = fetch_data_from_db(engine, query)
    if result.empty or "latest_week" not in result.columns:
        return None
    latest_value = result.iloc[0]["latest_week"]
    if pd.isna(latest_value):
        return None
    return pd.Timestamp(latest_value)


def get_latest_cache_week() -> pd.Timestamp | None:
    cached = _load_parquet_or_empty(sales_cache_file)
    if cached.empty or "Week" not in cached.columns:
        return None
    cached["Week"] = pd.to_datetime(cached["Week"], errors="coerce")
    latest = cached["Week"].dropna().max()
    if pd.isna(latest):
        return None
    return pd.Timestamp(latest)


def _read_inventory_snapshot(raw_folder: Path) -> pd.DataFrame:
    week_ending = parse_week_ending(raw_folder.name)
    inventory_result = build_inventory_working_file(raw_folder=raw_folder)
    pos_df = pd.read_excel(
        inventory_result.snapshot_file,
        dtype={"ISBN": "string"},
    )
    required_columns = ["ISBN", "OH_Stores", "OO_Stores", "OH_DC", "OO_DC"]
    missing = [col for col in required_columns if col not in pos_df.columns]
    if missing:
        raise ValueError(f"Combined inventory snapshot is missing required columns: {missing}")

    snapshot = pos_df.loc[:, required_columns].copy()
    snapshot["ISBN"] = normalize_isbn_series(snapshot["ISBN"])
    snapshot = snapshot[snapshot["ISBN"].notna()].drop_duplicates(subset=["ISBN"], keep="first")
    snapshot["Week"] = pd.Timestamp(week_ending.date())
    for col in ["OH_Stores", "OO_Stores", "OH_DC", "OO_DC"]:
        snapshot[col] = pd.to_numeric(snapshot[col], errors="coerce").fillna(0).astype(int)
    return snapshot.reset_index(drop=True)


def refresh_inventory_cache(raw_folder: str | Path | None = None, full_refresh: bool = False) -> pd.DataFrame:
    cached = _load_parquet_or_empty(inventory_cache_file)
    folders: list[Path]
    explicit_raw_folder = raw_folder is not None
    if full_refresh:
        folders = get_candidate_raw_folders()
    elif raw_folder:
        folders = [resolve_raw_folder(raw_folder)]
    else:
        folders = [get_candidate_raw_folders()[-1]]

    existing_weeks: set[pd.Timestamp] = set()
    if not cached.empty:
        cached["Week"] = pd.to_datetime(cached["Week"])
        existing_weeks = set(cached["Week"].dropna().tolist())

    snapshots: list[pd.DataFrame] = []
    for folder in folders:
        week_ending = pd.Timestamp(parse_week_ending(folder.name).date())
        if not full_refresh and not explicit_raw_folder and week_ending in existing_weeks:
            continue
        snapshots.append(_read_inventory_snapshot(folder))

    if not snapshots:
        if cached.empty:
            raise ValueError("No B&N inventory snapshots were available to cache.")
        return cached

    combined = (
        pd.concat([cached, *snapshots], ignore_index=True)
        if not cached.empty and not full_refresh
        else pd.concat(snapshots, ignore_index=True)
    )
    combined["Week"] = pd.to_datetime(combined["Week"])
    combined = combined.drop_duplicates(subset=["Week", "ISBN"], keep="last").sort_values(["Week", "ISBN"]).reset_index(drop=True)
    _save_parquet(combined, inventory_cache_file)
    return combined


def _last_non_null(series: pd.Series):
    non_null = series.dropna()
    if non_null.empty:
        return None
    return non_null.iloc[-1]


def _build_base_metadata(sales_df: pd.DataFrame) -> pd.DataFrame:
    metadata_columns = [
        column
        for column in ["Publisher", "PT", "CAT", "pgrp", "Title", "Price", "PubDate", "SubjectCode", "DeptCode"]
        if column in sales_df.columns
    ]
    metadata = (
        sales_df.sort_values(["Week", "ISBN"])
        .groupby("ISBN", as_index=False)
        .agg({column: _last_non_null for column in metadata_columns})
    )
    metadata.rename(columns={"Publisher": "Pub"}, inplace=True)
    metadata["PubDate"] = pd.to_datetime(metadata["PubDate"], errors="coerce").dt.strftime("%m-%d-%Y")
    metadata["PubDate"] = metadata["PubDate"].fillna("")
    return metadata


def _fetch_item_metadata_for_isbns(isbns: list[str]) -> pd.DataFrame:
    if not isbns:
        return pd.DataFrame(columns=["ISBN", "Pub", "PT", "CAT", "pgrp", "Title", "Price", "PubDate"])

    quoted_isbns = ",".join(f"'{isbn}'" for isbn in sorted(set(isbns)))
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
        i.ITEM_TITLE IN ({quoted_isbns})
        AND i.PUBLISHER_CODE IS NOT NULL
        AND i.PUBLISHER_CODE NOT IN (
            'Benefit',
            'AFO LLC',
            'Glam Media',
            'PQ Blackwell',
            'PRINCETON',
            'AMMO Books',
            'San Francisco Art Institute',
            'FareArts',
            'Sager',
            'In Active',
            'Driscolls',
            'Impossible Foods',
            'Moleskine'
        )
        AND i.PRODUCT_TYPE IN ('BK', 'FT', 'RP', 'CP', 'DI')
        AND i.PUBLISHING_GROUP NOT IN ('MKT', 'ZZZ');
    """
    engine = get_connection()
    metadata = fetch_data_from_db(engine, query)
    if metadata.empty:
        return metadata

    metadata["ISBN"] = normalize_isbn_series(metadata["ISBN"].astype("string"))
    metadata["PubDate"] = pd.to_datetime(metadata["PubDate"], errors="coerce").dt.strftime("%m-%d-%Y")
    metadata["PubDate"] = metadata["PubDate"].fillna("")
    return metadata


def _series_by_isbn(sales_df: pd.DataFrame, mask: pd.Series) -> pd.Series:
    if not mask.any():
        return pd.Series(dtype="float64")
    return sales_df.loc[mask].groupby("ISBN")["qty"].sum()


def _history_columns(df: pd.DataFrame) -> list[str]:
    return [
        column
        for column in df.columns
        if isinstance(column, str)
        and (
            (len(column) == 10 and column.count("-") == 2)
            or column.startswith("12-31-")
        )
    ]


def _prune_zero_history_columns(df: pd.DataFrame) -> pd.DataFrame:
    history_columns = _history_columns(df)
    keep_history_columns = [
        column
        for column in history_columns
        if pd.to_numeric(df[column], errors="coerce").fillna(0).sum() != 0
    ]
    keep_columns = [
        column for column in df.columns if column not in history_columns or column in keep_history_columns
    ]
    return df.loc[:, keep_columns].copy()


def build_report_dataframe(sales_df: pd.DataFrame, inventory_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.Timestamp]:
    if sales_df.empty:
        raise ValueError("Sales cache is empty.")

    sales_df = sales_df.copy()
    if "Publisher" in sales_df.columns:
        sales_df["Publisher"] = sales_df["Publisher"].map(_normalize_publisher_value)
    sales_df["Week"] = pd.to_datetime(sales_df["Week"])
    latest_week = sales_df["Week"].max()

    qty_by_week = (
        sales_df.pivot_table(index="ISBN", columns="Week", values="qty", aggfunc="sum", fill_value=0)
        .sort_index(axis=1, ascending=False)
    )
    qty_by_week.columns = pd.to_datetime(qty_by_week.columns)

    weekly_columns = [week for week in qty_by_week.columns if week >= WEEKLY_HISTORY_START]
    weekly_df = qty_by_week.loc[:, weekly_columns].copy() if weekly_columns else qty_by_week.iloc[:, 0:0].copy()
    weekly_df.columns = [week.strftime("%m-%d-%Y") for week in weekly_columns]

    yearly_frames: list[pd.Series] = []
    for year in range(YEARLY_HISTORY_END, YEARLY_HISTORY_START - 1, -1):
        mask = sales_df["Week"].dt.year == year
        yearly_series = _series_by_isbn(sales_df, mask)
        yearly_series.name = f"12-31-{year}"
        yearly_frames.append(yearly_series)
    yearly_df = pd.concat(yearly_frames, axis=1).fillna(0).astype(int) if yearly_frames else pd.DataFrame(index=qty_by_week.index)

    latest_iso = latest_week.isocalendar()
    iso_parts = sales_df["Week"].dt.isocalendar()
    tytd = _series_by_isbn(sales_df, iso_parts.year == latest_iso.year)
    lytd = _series_by_isbn(
        sales_df,
        (iso_parts.year == latest_iso.year - 1) & (iso_parts.week <= latest_iso.week),
    )
    ly_fy = _series_by_isbn(sales_df, sales_df["Week"].dt.year == latest_week.year - 1)
    ltd = sales_df.groupby("ISBN")["qty"].sum()
    w52 = _series_by_isbn(sales_df, sales_df["Week"].between(latest_week - pd.Timedelta(weeks=51), latest_week))
    last6 = _series_by_isbn(sales_df, sales_df["Week"].between(latest_week - pd.Timedelta(weeks=5), latest_week))

    metrics = pd.concat(
        [
            w52.rename("W52"),
            (last6 / 6.0).round(0).rename("6Wk Avg"),
            tytd.rename("TYTD"),
            lytd.rename("LYTD"),
            (tytd.subtract(lytd, fill_value=0)).rename("YTD Var"),
            ly_fy.rename("LY_FY"),
            ltd.rename("LTD"),
        ],
        axis=1,
    ).fillna(0)

    inventory_df = inventory_df.copy()
    inventory_df["Week"] = pd.to_datetime(inventory_df["Week"])
    inventory_week = latest_week if (inventory_df["Week"] == latest_week).any() else inventory_df["Week"].max()
    latest_inventory = (
        inventory_df[inventory_df["Week"] == inventory_week]
        .drop_duplicates(subset=["ISBN"], keep="last")
        .set_index("ISBN")[["OH_Stores", "OO_Stores", "OH_DC", "OO_DC"]]
        if not inventory_df.empty
        else pd.DataFrame(columns=["OH_Stores", "OO_Stores", "OH_DC", "OO_DC"])
    )
    if not latest_inventory.empty:
        latest_inventory["Total_OH"] = latest_inventory["OH_Stores"] + latest_inventory["OH_DC"]
        latest_inventory["Total_OO"] = latest_inventory["OO_Stores"] + latest_inventory["OO_DC"]
        latest_inventory["Avg_Wk_OH"] = (
            (latest_inventory["OH_Stores"] + latest_inventory["OH_DC"]) / 2.0
        ).round(2)

    metadata = _build_base_metadata(sales_df).set_index("ISBN")
    inventory_isbns = latest_inventory.index.astype(str).tolist() if not latest_inventory.empty else []
    missing_metadata_isbns = [isbn for isbn in inventory_isbns if isbn not in metadata.index.astype(str)]
    if missing_metadata_isbns:
        item_metadata = _fetch_item_metadata_for_isbns(missing_metadata_isbns)
        if not item_metadata.empty:
            item_metadata = item_metadata.set_index("ISBN")
            metadata = pd.concat([metadata, item_metadata[~item_metadata.index.isin(metadata.index)]], axis=0)
    report = pd.concat([metadata, latest_inventory, metrics, weekly_df, yearly_df], axis=1)
    report = report.reset_index()
    report.rename(columns={"index": "ISBN"}, inplace=True)

    ordered_columns = [
        "Pub",
        "PT",
        "CAT",
        "pgrp",
        "ISBN",
        "Title",
        "Price",
        "PubDate",
        "SubjectCode",
        "DeptCode",
        "OH_Stores",
        "OO_Stores",
        "OH_DC",
        "OO_DC",
        "Total_OH",
        "Total_OO",
        "Avg_Wk_OH",
        "W52",
        "6Wk Avg",
        "TYTD",
        "LYTD",
        "YTD Var",
        "LY_FY",
        "LTD",
    ]
    dynamic_columns = [column for column in report.columns if column not in ordered_columns]
    report = report[ordered_columns + dynamic_columns]

    text_columns = ["Pub", "PT", "CAT", "pgrp", "Title", "PubDate"]
    for column in text_columns:
        if column in report.columns:
            report[column] = report[column].fillna("")

    code_columns = ["SubjectCode", "DeptCode"]
    for column in code_columns:
        if column in report.columns:
            report[column] = report[column].fillna("").astype(str)

    int_columns = [
        "OH_Stores",
        "OO_Stores",
        "OH_DC",
        "OO_DC",
        "Total_OH",
        "Total_OO",
        "W52",
        "6Wk Avg",
        "TYTD",
        "LYTD",
        "YTD Var",
        "LY_FY",
        "LTD",
    ] + list(weekly_df.columns) + list(yearly_df.columns)
    for column in int_columns:
        if column in report.columns:
            report[column] = pd.to_numeric(report[column], errors="coerce").fillna(0).astype(int)

    report["Price"] = pd.to_numeric(report["Price"], errors="coerce")
    if "Avg_Wk_OH" in report.columns:
        report["Avg_Wk_OH"] = pd.to_numeric(report["Avg_Wk_OH"], errors="coerce").fillna(0).round(2)
    latest_week_label = latest_week.strftime("%m-%d-%Y")
    sort_columns = [latest_week_label]
    ascending = [False]
    for column in ["Pub", "pgrp", "Title", "ISBN"]:
        if column in report.columns:
            sort_columns.append(column)
            ascending.append(True)
    report = report.sort_values(sort_columns, ascending=ascending).reset_index(drop=True)
    return report, latest_week


def _build_save_options(report_df: pd.DataFrame, latest_week: pd.Timestamp) -> dict[str, object]:
    date_cols = [column for column in report_df.columns if isinstance(column, str) and len(column) == 10 and column.count("-") == 2]
    summary_cols = [
        "OH_Stores",
        "OO_Stores",
        "OH_DC",
        "OO_DC",
        "Total_OH",
        "Total_OO",
        "W52",
        "TYTD",
        "LYTD",
        "YTD Var",
        "LY_FY",
        "LTD",
    ]
    format_cols = date_cols + summary_cols
    decimal_cols = ["Price", "Avg_Wk_OH"]
    totals = build_column_totals(report_df, format_cols)
    top_row_groups = [
        {"label": "STORE", "start_col": 10, "end_col": 11},
        {"label": "DC", "start_col": 12, "end_col": 13},
        {"label": "Total", "start_col": 14, "end_col": 15},
    ]
    header_overrides = {
        10: "O/H",
        11: "O/O",
        12: "O/H",
        13: "O/O",
        14: "O/H",
        15: "O/O",
        16: "Avg OH",
        17: "52 WK",
    }
    header_fill_overrides = {
        10: "#E6B8B7",
        11: "#E6B8B7",
        12: "#E6B8B7",
        13: "#E6B8B7",
        14: "#E6B8B7",
        15: "#E6B8B7",
    }
    accounting_format = {
        "num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
    }
    column_format_overrides = {
        col_idx: {"format": accounting_format}
        for col_idx in range(10, min(24, len(report_df.columns)))
    }
    if 23 in column_format_overrides:
        column_format_overrides[23]["width"] = 10
    history_top_labels: dict[int, str] = {}
    for col_idx, column in enumerate(report_df.columns):
        if isinstance(column, str) and column.startswith("12-31-"):
            history_top_labels[col_idx] = f"FY {column[-4:]}"
    title_block = {
        "start_row": 1,
        "end_row": 2,
        "start_col": 5,
        "end_col": 5,
        "title": "Rolling B&N POS",
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
        "pre_date_column_count": 24,
        "summary_label_col_idx": 10,
        "top_row_groups": top_row_groups,
        "header_overrides": header_overrides,
        "header_fill_overrides": header_fill_overrides,
        "column_width_overrides": None,
        "format_blank_summary_cells": False,
        "title_block": title_block,
        "header_row_override": 5,
        "show_weeknum_label": True,
        "history_top_labels": history_top_labels,
        "weeknum_label_fill": "#C4BD97",
        "column_format_overrides": column_format_overrides,
    }


def save_reports_by_pub(report_df: pd.DataFrame, latest_week: pd.Timestamp) -> int:
    filename = _format_bn_output_filename(latest_week)
    saved_count = 0
    for publisher, folder in bn_dp_folders.items():
        df_pub = report_df[report_df["Pub"] == publisher].copy()
        if df_pub.empty:
            print(f"No data for {publisher}")
            continue
        df_pub = _prune_zero_history_columns(df_pub)
        folder.mkdir(parents=True, exist_ok=True)
        output_file = folder / filename
        print(f"{'':#<40}")
        print(f"{publisher.center(40)}")
        print(f"{'':#<40}")
        print(f"Saving Barnes & Noble rolling report for {publisher} to folder: {folder}")
        save_to_excel(df_pub, output_file, **_build_save_options(df_pub, latest_week))
        print(f"Saved Barnes & Noble rolling report for {publisher} to {output_file}")
        print()
        saved_count += 1
    return saved_count


def build_customer_sales_report(
    raw_folder: str | Path | None = None,
    refresh_sales: bool = False,
    refresh_inventory: bool = False,
    full_refresh: bool = False,
    save_dp: bool = False,
    local_only: bool = False,
    manual_missing_weeks_workbook: str | Path | None = None,
) -> RollingBuildResult:
    if manual_missing_weeks_workbook:
        refresh_manual_missing_weeks_cache(manual_missing_weeks_workbook)

    sales_df = (
        refresh_sales_cache(full_refresh=full_refresh)
        if refresh_sales or full_refresh or not sales_cache_file.exists()
        else _load_parquet_or_empty(sales_cache_file)
    )
    sales_df = _apply_manual_missing_weeks(sales_df)
    inventory_df = (
        refresh_inventory_cache(raw_folder=raw_folder, full_refresh=full_refresh)
        if refresh_inventory or full_refresh or not inventory_cache_file.exists()
        else _load_parquet_or_empty(inventory_cache_file)
    )

    report_df, latest_week = build_report_dataframe(sales_df, inventory_df)
    output_dir = local_review_dir if local_only else bn_rolling_folder
    output_file = output_dir / _format_bn_output_filename(latest_week)
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
    raw_folder: str | Path | None = None,
    full_refresh: bool = False,
) -> CacheRefreshResult:
    latest_sql_week = get_latest_sql_week()
    sales_df = refresh_sales_cache(full_refresh=full_refresh)
    inventory_df = refresh_inventory_cache(raw_folder=raw_folder, full_refresh=full_refresh)
    sales_cache_week = pd.to_datetime(sales_df["Week"]).max() if not sales_df.empty else None
    inventory_cache_week = pd.to_datetime(inventory_df["Week"]).max() if not inventory_df.empty else None
    return CacheRefreshResult(
        latest_sql_week=latest_sql_week,
        sales_cache_week=sales_cache_week,
        inventory_cache_week=inventory_cache_week,
        sales_rows=len(sales_df),
        inventory_rows=len(inventory_df),
    )


def print_result_summary(result: RollingBuildResult) -> None:
    print()
    print(f"Latest week: {result.latest_week:%Y-%m-%d}")
    print(f"Sales cache rows: {result.sales_rows:,}")
    print(f"Inventory cache rows: {result.inventory_rows:,}")
    print(f"Report shape: {result.report_shape}")
    print(f"Saved file: {result.output_file}")
    if result.dp_files_saved:
        print(f"DP files saved: {result.dp_files_saved}")


def print_cache_refresh_summary(result: CacheRefreshResult) -> None:
    print()
    print(
        "Latest SQL week: "
        f"{result.latest_sql_week.strftime('%Y-%m-%d') if result.latest_sql_week is not None else 'None'}"
    )
    print(
        "Sales cache latest week: "
        f"{result.sales_cache_week.strftime('%Y-%m-%d') if result.sales_cache_week is not None else 'None'}"
    )
    print(
        "Inventory cache latest week: "
        f"{result.inventory_cache_week.strftime('%Y-%m-%d') if result.inventory_cache_week is not None else 'None'}"
    )
    print(f"Sales cache rows: {result.sales_rows:,}")
    print(f"Inventory cache rows: {result.inventory_rows:,}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build the Barnes & Noble customer sales rolling report.")
    parser.add_argument("--raw-folder", help="Optional raw folder to use when refreshing the inventory snapshot cache.")
    parser.add_argument("--refresh-sales-cache", action="store_true", help="Refresh the sales cache from SQL before building.")
    parser.add_argument("--refresh-inventory-cache", action="store_true", help="Refresh the inventory snapshot cache before building.")
    parser.add_argument("--full-refresh", action="store_true", help="Rebuild both caches from scratch before building.")
    parser.add_argument("--save-dp", action="store_true", help="Also save publisher-filtered versions to the B&N DP folders.")
    parser.add_argument("--local-only", action="store_true", help="Save the rolling report to the local bn_rolling_reports/review_output folder instead of the network target.")
    parser.add_argument("--manual-missing-weeks-workbook", help="Optional manual workbook used to refresh the supplemental parquet for SQL-missing weekly history.")
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    result = build_customer_sales_report(
        raw_folder=args.raw_folder,
        refresh_sales=args.refresh_sales_cache,
        refresh_inventory=args.refresh_inventory_cache,
        full_refresh=args.full_refresh,
        save_dp=args.save_dp,
        local_only=args.local_only,
        manual_missing_weeks_workbook=args.manual_missing_weeks_workbook,
    )
    print_result_summary(result)


if __name__ == "__main__":
    main()
