from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.append(str(Path(__file__).resolve().parents[1]))

from amazon_rolling_reports.functions import build_column_totals, save_to_excel
from bn_rolling_reports.isbn_utils import normalize_isbn_series
from shared.bookscan_calendar import bookscan_parts, bookscan_week
from shared.db.connection import get_connection
from shared.db.query_runner import fetch_data_from_db
from shared.pg_grouping import apply_pg_grouping

try:
    from .rolling_paths import (
        edelweiss_rolling_folder,
        local_review_dir,
        manual_missing_weeks_file,
        metadata_cache_file,
        sales_cache_file,
        sample_workbook,
    )
    from .rolling_queries import (
        DISTINCT_WEEKS_QUERY,
        LATEST_INVENTORY_QUERY,
        LATEST_WEEK_QUERY,
        MISSING_WEEKS_QUERY,
        SOURCE_METADATA_QUERY,
        build_distinct_weeks_since_query,
        build_sales_query,
    )
except ImportError:
    from rolling_paths import (
        edelweiss_rolling_folder,
        local_review_dir,
        manual_missing_weeks_file,
        metadata_cache_file,
        sales_cache_file,
        sample_workbook,
    )
    from rolling_queries import (
        DISTINCT_WEEKS_QUERY,
        LATEST_INVENTORY_QUERY,
        LATEST_WEEK_QUERY,
        MISSING_WEEKS_QUERY,
        SOURCE_METADATA_QUERY,
        build_distinct_weeks_since_query,
        build_sales_query,
    )


REFRESH_LOOKBACK_WEEKS = 0


@dataclass
class RollingBuildResult:
    output_file: Path
    latest_week: pd.Timestamp
    sales_rows: int
    metadata_rows: int
    report_shape: tuple[int, int]


@dataclass
class CacheRefreshResult:
    refresh_mode: str
    latest_sql_week: pd.Timestamp | None
    sales_cache_week: pd.Timestamp | None
    expected_next_week: pd.Timestamp | None
    missing_week_count: int
    missing_weeks: list[pd.Timestamp]
    sales_rows: int
    metadata_rows: int


@dataclass
class WeekCheckResult:
    min_week: pd.Timestamp | None
    max_week: pd.Timestamp | None
    missing_weeks: list[pd.Timestamp]
    latest_weeks: list[pd.Timestamp]


@dataclass
class DeltaWeekStatus:
    latest_sql_week: pd.Timestamp | None
    latest_cache_week: pd.Timestamp | None
    expected_next_week: pd.Timestamp | None
    missing_weeks: list[pd.Timestamp]


def _ensure_cache_dir() -> None:
    sales_cache_file.parent.mkdir(parents=True, exist_ok=True)


def _load_parquet_or_empty(cache_file: Path) -> pd.DataFrame:
    if cache_file.exists():
        return pd.read_parquet(cache_file)
    return pd.DataFrame()


def _save_parquet(df: pd.DataFrame, cache_file: Path) -> None:
    _ensure_cache_dir()
    df.to_parquet(cache_file, index=False)


def normalize_edelweiss_isbn_series(series: pd.Series) -> pd.Series:
    return normalize_isbn_series(series).astype("object")


def _format_output_filename(week_ending: pd.Timestamp) -> str:
    week = bookscan_week(week_ending)
    return f"Week {week.week:02d} - {week.year} Rolling Edelweiss ({week_ending:%m%d%y}).xlsx"


def get_latest_sql_week() -> pd.Timestamp | None:
    result = fetch_data_from_db(get_connection(), LATEST_WEEK_QUERY)
    if result.empty or "latest_week" not in result.columns or pd.isna(result.at[0, "latest_week"]):
        return None
    return pd.Timestamp(result.at[0, "latest_week"])


def get_latest_cache_week() -> pd.Timestamp | None:
    cached = _load_parquet_or_empty(sales_cache_file)
    if cached.empty or "Week" not in cached.columns:
        return None
    cached["Week"] = pd.to_datetime(cached["Week"], errors="coerce")
    latest = cached["Week"].dropna().max()
    if pd.isna(latest):
        return None
    return pd.Timestamp(latest)


def check_source_weeks() -> WeekCheckResult:
    week_result = fetch_data_from_db(get_connection(), DISTINCT_WEEKS_QUERY)
    if week_result.empty:
        return WeekCheckResult(None, None, [], [])
    week_result["Week"] = pd.to_datetime(week_result["Week"], errors="coerce")
    weeks = sorted(week_result["Week"].dropna().tolist())
    missing_result = fetch_data_from_db(get_connection(), MISSING_WEEKS_QUERY)
    missing: list[pd.Timestamp] = []
    if not missing_result.empty and "missing_week" in missing_result.columns:
        missing_result["missing_week"] = pd.to_datetime(missing_result["missing_week"], errors="coerce")
        missing = sorted(missing_result["missing_week"].dropna().tolist())
    return WeekCheckResult(weeks[0], weeks[-1], missing, weeks[-12:])


def _expected_next_week(week: pd.Timestamp | None) -> pd.Timestamp | None:
    if week is None:
        return None
    return pd.Timestamp(week) + pd.Timedelta(days=7)


def get_delta_week_status(
    latest_cache_week: pd.Timestamp | None = None,
    latest_sql_week: pd.Timestamp | None = None,
) -> DeltaWeekStatus:
    cache_week = pd.Timestamp(latest_cache_week) if latest_cache_week is not None else get_latest_cache_week()
    sql_week = pd.Timestamp(latest_sql_week) if latest_sql_week is not None else get_latest_sql_week()
    expected_next_week = _expected_next_week(cache_week)
    if expected_next_week is None or sql_week is None or expected_next_week > sql_week:
        return DeltaWeekStatus(sql_week, cache_week, expected_next_week, [])

    result = fetch_data_from_db(
        get_connection(),
        build_distinct_weeks_since_query(expected_next_week.strftime("%Y-%m-%d")),
    )
    if result.empty:
        expected = pd.date_range(start=expected_next_week, end=sql_week, freq="7D")
        return DeltaWeekStatus(sql_week, cache_week, expected_next_week, list(expected))

    result["Week"] = pd.to_datetime(result["Week"], errors="coerce")
    weeks = pd.DatetimeIndex(sorted(result["Week"].dropna().tolist()))
    expected = pd.date_range(start=expected_next_week, end=sql_week, freq="7D")
    missing = list(expected.difference(weeks))
    return DeltaWeekStatus(sql_week, cache_week, expected_next_week, missing)


def _sample_week_columns(df: pd.DataFrame) -> list[object]:
    return [column for column in df.columns if hasattr(column, "strftime")]


def refresh_manual_missing_weeks_cache(workbook_path: str | Path = sample_workbook) -> pd.DataFrame:
    workbook = Path(workbook_path)
    if not workbook.exists():
        raise FileNotFoundError(f"Edelweiss sample workbook not found: {workbook}")
    manual_df = pd.read_excel(workbook, sheet_name=0, header=5)
    week_columns = _sample_week_columns(manual_df)
    if not week_columns:
        raise ValueError("The Edelweiss workbook did not contain any weekly date columns.")

    required = ["Pub", "PT", "CAT", "PGR", "ISBN 13", "TITLE", "PRICE", "SHIP", "On Hand", "On Order"]
    missing = [column for column in required if column not in manual_df.columns]
    if missing:
        raise ValueError(f"The Edelweiss workbook is missing required columns: {missing}")

    supplemental = manual_df.loc[:, required + week_columns].copy()
    supplemental = supplemental.rename(
        columns={
            "PGR": "PGRP",
            "ISBN 13": "ISBN",
            "SHIP": "PubDate",
        }
    )
    supplemental["ISBN"] = normalize_edelweiss_isbn_series(supplemental["ISBN"].astype("string"))
    supplemental = supplemental[supplemental["ISBN"].notna()].copy()
    for column in ["Pub", "PT", "CAT", "PGRP", "TITLE"]:
        supplemental[column] = supplemental[column].where(supplemental[column].isna(), supplemental[column].astype(str))
    supplemental["PRICE"] = pd.to_numeric(supplemental["PRICE"], errors="coerce")
    supplemental["PubDate"] = pd.to_datetime(supplemental["PubDate"], errors="coerce")
    supplemental["On Hand"] = pd.to_numeric(supplemental["On Hand"], errors="coerce").fillna(0)
    supplemental["On Order"] = pd.to_numeric(supplemental["On Order"], errors="coerce").fillna(0)

    long_df = supplemental.melt(
        id_vars=["Pub", "PT", "CAT", "PGRP", "ISBN", "TITLE", "PRICE", "PubDate", "On Hand", "On Order"],
        value_vars=week_columns,
        var_name="Week",
        value_name="qty",
    )
    long_df["Week"] = pd.to_datetime(long_df["Week"], errors="coerce")
    long_df["qty"] = pd.to_numeric(long_df["qty"], errors="coerce").fillna(0)
    long_df = long_df[(long_df["Week"].notna()) & (long_df["qty"] != 0)].copy()
    long_df["qty"] = long_df["qty"].astype(int)
    long_df = (
        long_df.groupby(
            ["Week", "ISBN", "Pub", "PT", "CAT", "PGRP", "TITLE", "PRICE", "PubDate", "On Hand", "On Order"],
            as_index=False,
            dropna=False,
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
    supplemental["Week"] = pd.to_datetime(supplemental["Week"], errors="coerce")
    supplemental["ISBN"] = normalize_edelweiss_isbn_series(supplemental["ISBN"].astype("string"))
    supplemental["qty"] = pd.to_numeric(supplemental["qty"], errors="coerce").fillna(0).astype(int)
    supplemental["PubDate"] = pd.to_datetime(supplemental["PubDate"], errors="coerce")
    return supplemental[supplemental["ISBN"].notna()].copy()


def _apply_manual_missing_weeks(sales_df: pd.DataFrame) -> pd.DataFrame:
    supplemental = _load_manual_missing_weeks()
    if supplemental.empty:
        return sales_df
    existing_weeks = set(pd.to_datetime(sales_df["Week"], errors="coerce").dropna())
    sql_min = pd.to_datetime(sales_df["Week"], errors="coerce").min()
    sql_max = pd.to_datetime(sales_df["Week"], errors="coerce").max()
    expected = pd.date_range(sql_min, sql_max, freq="W-SAT")
    missing_weeks = set(expected.difference(pd.DatetimeIndex(sorted(existing_weeks))))
    supplemental = supplemental[supplemental["Week"].isin(missing_weeks)].copy()
    if supplemental.empty:
        return sales_df
    supplemental = supplemental[["Week", "ISBN", "qty"]].copy()
    combined = pd.concat([sales_df.copy(), supplemental], ignore_index=True, sort=False)
    combined["Week"] = pd.to_datetime(combined["Week"], errors="coerce")
    combined["ISBN"] = normalize_edelweiss_isbn_series(combined["ISBN"].astype("string"))
    combined["qty"] = pd.to_numeric(combined["qty"], errors="coerce").fillna(0).astype(int)
    combined = combined[combined["ISBN"].notna() & combined["Week"].notna()].copy()
    combined = combined.groupby(["Week", "ISBN"], as_index=False)["qty"].sum()
    return combined.sort_values(["Week", "ISBN"]).reset_index(drop=True)


def refresh_sales_cache(
    full_refresh: bool = False,
    refresh_lookback_weeks: int = REFRESH_LOOKBACK_WEEKS,
) -> pd.DataFrame:
    refresh_lookback_weeks = max(int(refresh_lookback_weeks), 0)
    cached = _load_parquet_or_empty(sales_cache_file)
    start_date = "2019-01-01"
    if not full_refresh and not cached.empty and "Week" in cached.columns:
        cached["Week"] = pd.to_datetime(cached["Week"], errors="coerce")
        latest_cached_week = cached["Week"].dropna().max()
        if pd.notna(latest_cached_week):
            start_week = latest_cached_week + pd.Timedelta(days=7)
            if refresh_lookback_weeks > 0:
                start_week = latest_cached_week - pd.Timedelta(weeks=refresh_lookback_weeks)
            start_date = start_week.strftime("%Y-%m-%d")

    fetched = fetch_data_from_db(get_connection(), build_sales_query(start_date))
    if fetched.empty:
        if cached.empty:
            raise ValueError("The Edelweiss sales query returned no rows.")
        return cached
    fetched["Week"] = pd.to_datetime(fetched["Week"], errors="coerce")
    fetched["ISBN"] = normalize_edelweiss_isbn_series(fetched["RawISBN"].astype("string"))
    fetched["qty"] = pd.to_numeric(fetched["qty"], errors="coerce").fillna(0).astype(int)
    fetched = fetched[fetched["ISBN"].notna() & fetched["Week"].notna()].copy()
    fetched = fetched.groupby(["Week", "ISBN"], as_index=False)["qty"].sum()

    if full_refresh or cached.empty:
        combined = fetched
    else:
        fetched_min_week = fetched["Week"].min()
        fetched_max_week = fetched["Week"].max()
        cached = cached[~cached["Week"].between(fetched_min_week, fetched_max_week)].copy()
        combined = pd.concat([cached, fetched], ignore_index=True, sort=False)
    combined["Week"] = pd.to_datetime(combined["Week"], errors="coerce")
    combined = combined.groupby(["Week", "ISBN"], as_index=False)["qty"].sum()
    combined = combined.sort_values(["Week", "ISBN"]).reset_index(drop=True)
    _save_parquet(combined, sales_cache_file)
    return combined


def _last_non_null(series: pd.Series):
    non_null = series.dropna()
    if non_null.empty:
        return None
    return non_null.iloc[-1]


def _build_manual_metadata() -> pd.DataFrame:
    supplemental = _load_manual_missing_weeks()
    if supplemental.empty:
        return pd.DataFrame(columns=["ISBN", "Pub", "PT", "CAT", "PGRP", "TITLE", "PRICE", "PubDate", "On Hand", "On Order"])
    return (
        supplemental.sort_values(["Week", "ISBN"])
        .groupby("ISBN", as_index=False)
        .agg(
            {
                "Pub": _last_non_null,
                "PT": _last_non_null,
                "CAT": _last_non_null,
                "PGRP": _last_non_null,
                "TITLE": _last_non_null,
                "PRICE": _last_non_null,
                "PubDate": _last_non_null,
                "On Hand": _last_non_null,
                "On Order": _last_non_null,
            }
        )
    )


def refresh_metadata_cache(force: bool = False) -> pd.DataFrame:
    cached = pd.DataFrame() if force else _load_parquet_or_empty(metadata_cache_file)
    if force or cached.empty:
        cached = fetch_data_from_db(get_connection(), SOURCE_METADATA_QUERY)
    manual = _build_manual_metadata()
    frames = [frame for frame in [cached, manual] if not frame.empty]
    if not frames:
        return pd.DataFrame(columns=["ISBN", "Pub", "PT", "CAT", "PGRP", "TITLE", "PRICE", "PubDate", "On Hand", "On Order"])

    combined = pd.concat(frames, ignore_index=True, sort=False)
    combined["ISBN"] = normalize_edelweiss_isbn_series(combined["ISBN"].astype("string"))
    combined = combined[combined["ISBN"].notna()].copy()
    for column in ["Pub", "PT", "CAT", "PGRP", "TITLE"]:
        if column in combined.columns:
            combined[column] = combined[column].where(combined[column].isna(), combined[column].astype(str))
    combined["PRICE"] = pd.to_numeric(combined["PRICE"], errors="coerce")
    combined["PubDate"] = pd.to_datetime(combined["PubDate"], errors="coerce")
    combined = combined.drop_duplicates(subset=["ISBN"], keep="last").sort_values("ISBN").reset_index(drop=True)
    _save_parquet(combined, metadata_cache_file)
    return combined


def fetch_latest_inventory() -> pd.DataFrame:
    inventory = fetch_data_from_db(get_connection(), LATEST_INVENTORY_QUERY)
    if inventory.empty:
        return pd.DataFrame(columns=["ISBN", "On Hand", "On Order"])
    inventory["ISBN"] = normalize_edelweiss_isbn_series(inventory["RawISBN"].astype("string"))
    inventory = inventory[inventory["ISBN"].notna()].copy()
    for column in ["On Hand", "On Order"]:
        inventory[column] = pd.to_numeric(inventory[column], errors="coerce").fillna(0)
    return inventory.groupby("ISBN", as_index=False)[["On Hand", "On Order"]].sum()


def _series_by_isbn(sales_df: pd.DataFrame, mask: pd.Series) -> pd.Series:
    if not mask.any():
        return pd.Series(dtype="float64")
    return sales_df.loc[mask].groupby("ISBN")["qty"].sum()


def build_report_dataframe(
    sales_df: pd.DataFrame,
    metadata_df: pd.DataFrame,
    inventory_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.Timestamp]:
    if sales_df.empty:
        raise ValueError("Edelweiss sales cache is empty.")

    sales_df = _apply_manual_missing_weeks(sales_df)
    sales_df["Week"] = pd.to_datetime(sales_df["Week"], errors="coerce")
    sales_df["ISBN"] = normalize_edelweiss_isbn_series(sales_df["ISBN"].astype("string"))
    sales_df["qty"] = pd.to_numeric(sales_df["qty"], errors="coerce").fillna(0).astype(int)
    sales_df = sales_df[sales_df["ISBN"].notna() & sales_df["Week"].notna()].copy()
    latest_week = sales_df["Week"].max()
    base_index = pd.Index(sorted(set(sales_df["ISBN"].astype(str)) | set(metadata_df.get("ISBN", pd.Series(dtype=str)).dropna().astype(str))))

    qty_by_week = (
        sales_df.pivot_table(index="ISBN", columns="Week", values="qty", aggfunc="sum", fill_value=0)
        .sort_index(axis=1, ascending=False)
    )
    qty_by_week.columns = pd.to_datetime(qty_by_week.columns)
    weekly_dates = sorted(pd.date_range(sales_df["Week"].min(), latest_week, freq="W-SAT"), reverse=True)
    weekly_df = qty_by_week.reindex(index=base_index, columns=weekly_dates, fill_value=0)
    weekly_df.columns = [week.strftime("%m-%d-%Y") for week in weekly_dates]

    latest_bookscan = bookscan_week(latest_week)
    bookscan_dates = bookscan_parts(sales_df["Week"])
    tytd = _series_by_isbn(sales_df, bookscan_dates["BookScanYear"] == latest_bookscan.year)
    lytd = _series_by_isbn(
        sales_df,
        (bookscan_dates["BookScanYear"] == latest_bookscan.year - 1)
        & (bookscan_dates["BookScanWeek"] <= latest_bookscan.week),
    )
    ly_fy = _series_by_isbn(sales_df, bookscan_dates["BookScanYear"] == latest_bookscan.year - 1)
    ltd = sales_df.groupby("ISBN")["qty"].sum()
    w52 = _series_by_isbn(sales_df, sales_df["Week"].between(latest_week - pd.Timedelta(weeks=51), latest_week))
    last6 = _series_by_isbn(sales_df, sales_df["Week"].between(latest_week - pd.Timedelta(weeks=5), latest_week))

    metrics = pd.concat(
        [
            w52.rename("52 WK"),
            (last6 / 6.0).round(2).rename("6Wk Avg"),
            tytd.rename("TYTD"),
            lytd.rename("LYTD"),
            tytd.subtract(lytd, fill_value=0).rename("YTD Var"),
            ly_fy.rename("LY_FY"),
            ltd.rename("LTD"),
        ],
        axis=1,
    ).fillna(0).reindex(base_index, fill_value=0)

    metadata = metadata_df.copy()
    if metadata.empty:
        metadata = pd.DataFrame(index=base_index)
    else:
        metadata["ISBN"] = normalize_edelweiss_isbn_series(metadata["ISBN"].astype("string"))
        metadata = metadata[metadata["ISBN"].notna()].drop_duplicates(subset=["ISBN"], keep="last").set_index("ISBN")
        metadata = metadata.reindex(base_index)
    for column in ["Pub", "PT", "CAT", "PGRP", "TITLE"]:
        if column not in metadata.columns:
            metadata[column] = ""
    for column in ["PRICE", "PubDate"]:
        if column not in metadata.columns:
            metadata[column] = pd.NA
    metadata["PubDate"] = pd.to_datetime(metadata["PubDate"], errors="coerce").dt.strftime("%m-%d-%Y").fillna("")

    inventory = inventory_df.copy()
    if inventory.empty:
        inventory = pd.DataFrame(index=base_index)
    else:
        inventory["ISBN"] = normalize_edelweiss_isbn_series(inventory["ISBN"].astype("string"))
        inventory = inventory[inventory["ISBN"].notna()].drop_duplicates(subset=["ISBN"], keep="last").set_index("ISBN")
        inventory = inventory.reindex(base_index)
    for column in ["On Hand", "On Order"]:
        if column not in inventory.columns:
            inventory[column] = 0

    report = pd.concat([metadata[["Pub", "PT", "CAT", "PGRP", "TITLE", "PRICE", "PubDate"]], inventory[["On Hand", "On Order"]], metrics, weekly_df], axis=1)
    report = report.reset_index().rename(columns={"index": "ISBN"})
    report = apply_pg_grouping(report, publisher_col="Pub", publishing_group_col="PGRP", product_type_col="PT")

    ordered_columns = [
        "Pub",
        "PT",
        "CAT",
        "PGRP",
        "PG_Grouping",
        "ISBN",
        "TITLE",
        "PRICE",
        "PubDate",
        "On Hand",
        "On Order",
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

    for column in ["Pub", "PT", "CAT", "PGRP", "PG_Grouping", "TITLE", "PubDate"]:
        report[column] = report[column].fillna("")
    report["PRICE"] = pd.to_numeric(report["PRICE"], errors="coerce")
    for column in ["On Hand", "On Order", "52 WK", "TYTD", "LYTD", "YTD Var", "LY_FY", "LTD"]:
        report[column] = pd.to_numeric(report[column], errors="coerce").fillna(0).astype(int)
    report["6Wk Avg"] = pd.to_numeric(report["6Wk Avg"], errors="coerce").fillna(0).round(2)

    latest_week_label = latest_week.strftime("%m-%d-%Y")
    report = report.sort_values([latest_week_label, "Pub", "PGRP", "TITLE", "ISBN"], ascending=[False, True, True, True, True])
    return report.reset_index(drop=True), latest_week


def _history_columns(df: pd.DataFrame) -> list[str]:
    return [column for column in df.columns if isinstance(column, str) and len(column) == 10 and column.count("-") == 2]


def _build_save_options(report_df: pd.DataFrame, latest_week: pd.Timestamp) -> dict[str, object]:
    date_cols = _history_columns(report_df)
    summary_cols = ["On Hand", "On Order", "52 WK", "TYTD", "LYTD", "YTD Var", "LY_FY", "LTD"]
    decimal_cols = ["PRICE", "6Wk Avg"]
    format_cols = date_cols + summary_cols
    totals = build_column_totals(report_df, format_cols + ["6Wk Avg"])
    accounting_format = {"num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'}
    decimal_accounting_format = {"num_format": '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'}
    column_format_overrides = {
        col_idx: {"format": accounting_format}
        for col_idx in range(9, min(18, len(report_df.columns)))
        if col_idx != 12
    }
    if 12 < len(report_df.columns):
        column_format_overrides[12] = {"format": decimal_accounting_format, "width": 10}
    if 6 < len(report_df.columns):
        column_format_overrides[6] = {"format": {"num_format": "General"}, "width": 42}
    title_block = {
        "start_row": 1,
        "end_row": 2,
        "start_col": 6,
        "end_col": 6,
        "title": "Edelweiss Rolling POS",
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
        "pre_date_column_count": 18,
        "summary_label_col_idx": 8,
        "header_fill_overrides": {9: "#E6B8B7", 10: "#E6B8B7"},
        "format_blank_summary_cells": False,
        "title_block": title_block,
        "header_row_override": 5,
        "show_weeknum_label": True,
        "weeknum_label_fill": "#C4BD97",
        "column_format_overrides": column_format_overrides,
    }


def build_customer_sales_report(
    refresh_sales: bool = False,
    refresh_manual_cache: bool = False,
    full_refresh: bool = False,
    refresh_lookback_weeks: int = REFRESH_LOOKBACK_WEEKS,
    local_only: bool = False,
    manual_missing_weeks_workbook: str | Path | None = None,
) -> RollingBuildResult:
    workbook = Path(manual_missing_weeks_workbook) if manual_missing_weeks_workbook else sample_workbook
    if refresh_manual_cache or full_refresh or not manual_missing_weeks_file.exists():
        refresh_manual_missing_weeks_cache(workbook)

    sales_df = (
        refresh_sales_cache(full_refresh=full_refresh, refresh_lookback_weeks=refresh_lookback_weeks)
        if refresh_sales or full_refresh or not sales_cache_file.exists()
        else _load_parquet_or_empty(sales_cache_file)
    )
    metadata_df = refresh_metadata_cache(force=full_refresh)
    inventory_df = fetch_latest_inventory()
    report_df, latest_week = build_report_dataframe(sales_df, metadata_df, inventory_df)
    output_dir = local_review_dir if local_only else edelweiss_rolling_folder
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / _format_output_filename(latest_week)
    save_to_excel(report_df, output_file, **_build_save_options(report_df, latest_week))
    return RollingBuildResult(
        output_file=output_file,
        latest_week=latest_week,
        sales_rows=len(sales_df),
        metadata_rows=len(metadata_df),
        report_shape=report_df.shape,
    )


def refresh_caches_only(
    refresh_manual_cache: bool = False,
    full_refresh: bool = False,
    refresh_lookback_weeks: int = REFRESH_LOOKBACK_WEEKS,
    manual_missing_weeks_workbook: str | Path | None = None,
) -> CacheRefreshResult:
    workbook = Path(manual_missing_weeks_workbook) if manual_missing_weeks_workbook else sample_workbook
    if refresh_manual_cache or full_refresh or not manual_missing_weeks_file.exists():
        refresh_manual_missing_weeks_cache(workbook)

    prior_cache_week = None if full_refresh else get_latest_cache_week()
    latest_sql_week = get_latest_sql_week()
    delta_status = get_delta_week_status(prior_cache_week, latest_sql_week)
    sales_df = refresh_sales_cache(full_refresh=full_refresh, refresh_lookback_weeks=refresh_lookback_weeks)
    metadata_df = refresh_metadata_cache(force=full_refresh)
    sales_cache_week = pd.to_datetime(sales_df["Week"]).max() if not sales_df.empty else None
    return CacheRefreshResult(
        refresh_mode="full" if full_refresh else "incremental",
        latest_sql_week=latest_sql_week,
        sales_cache_week=sales_cache_week,
        expected_next_week=delta_status.expected_next_week,
        missing_week_count=len(delta_status.missing_weeks),
        missing_weeks=delta_status.missing_weeks,
        sales_rows=len(sales_df),
        metadata_rows=len(metadata_df),
    )


def print_week_check(result: WeekCheckResult) -> None:
    print()
    print(
        "Source range: "
        f"{result.min_week.strftime('%Y-%m-%d') if result.min_week is not None else 'None'}"
        " -> "
        f"{result.max_week.strftime('%Y-%m-%d') if result.max_week is not None else 'None'}"
    )
    print(f"Missing SQL weeks: {len(result.missing_weeks)}")
    if result.missing_weeks:
        print("Missing SQL week list: " + ", ".join(week.strftime("%Y-%m-%d") for week in result.missing_weeks))
    if result.latest_weeks:
        print("Latest weeks present: " + ", ".join(week.strftime("%Y-%m-%d") for week in result.latest_weeks))


def print_result_summary(result: RollingBuildResult) -> None:
    print()
    print(f"Latest week: {result.latest_week:%Y-%m-%d}")
    print(f"Sales cache rows: {result.sales_rows:,}")
    print(f"Metadata cache rows: {result.metadata_rows:,}")
    print(f"Report shape: {result.report_shape}")
    print(f"Output file: {result.output_file}")


def print_cache_refresh_summary(result: CacheRefreshResult) -> None:
    print()
    print("Refresh mode: " + ("Full historical rebuild" if result.refresh_mode == "full" else "Latest changes only"))
    print(f"Latest SQL week: {result.latest_sql_week.strftime('%Y-%m-%d') if result.latest_sql_week is not None else 'None'}")
    print(f"Sales cache latest week: {result.sales_cache_week.strftime('%Y-%m-%d') if result.sales_cache_week is not None else 'None'}")
    print(f"Expected next week: {result.expected_next_week.strftime('%Y-%m-%d') if result.expected_next_week is not None else 'None'}")
    print(f"Missing SQL weeks detected: {result.missing_week_count}")
    if result.missing_weeks:
        print("Missing SQL week list: " + ", ".join(week.strftime("%Y-%m-%d") for week in result.missing_weeks))
    print(f"Sales cache rows: {result.sales_rows:,}")
    print(f"Metadata cache rows: {result.metadata_rows:,}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build the Edelweiss rolling report.")
    parser.add_argument("--refresh-sales-cache", action="store_true", help="Refresh the sales cache from SQL before building.")
    parser.add_argument("--refresh-manual-cache", action="store_true", help="Refresh the missing-week supplement cache from the sample workbook.")
    parser.add_argument("--full-refresh", action="store_true", help="Rebuild all Edelweiss caches from scratch before building.")
    parser.add_argument("--refresh-lookback-weeks", type=int, default=REFRESH_LOOKBACK_WEEKS, help="Optional overlap weeks to re-pull from SQL; default 0 means only weeks after the cache max.")
    parser.add_argument("--local-only", action="store_true", help="Save the report to local review_output instead of the network target.")
    parser.add_argument("--manual-missing-weeks-workbook", help="Optional Edelweiss workbook used to seed the missing-week supplement cache.")
    parser.add_argument("--check-weeks", action="store_true", help="Check the Edelweiss SQL source for missing weeks and exit.")
    parser.add_argument("--refresh-caches-only", action="store_true", help="Refresh caches and exit without building.")
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    if args.check_weeks:
        print_week_check(check_source_weeks())
        return
    if args.refresh_caches_only:
        print_cache_refresh_summary(
            refresh_caches_only(
                refresh_manual_cache=args.refresh_manual_cache,
                full_refresh=args.full_refresh,
                refresh_lookback_weeks=args.refresh_lookback_weeks,
                manual_missing_weeks_workbook=args.manual_missing_weeks_workbook,
            )
        )
        return
    print_result_summary(
        build_customer_sales_report(
            refresh_sales=True,
            refresh_manual_cache=args.refresh_manual_cache,
            full_refresh=args.full_refresh,
            refresh_lookback_weeks=args.refresh_lookback_weeks,
            local_only=args.local_only,
            manual_missing_weeks_workbook=args.manual_missing_weeks_workbook,
        )
    )


if __name__ == "__main__":
    main()
