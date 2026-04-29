from __future__ import annotations

import argparse
import glob
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell

try:
    from paths import process_paths
    from shared.db import fetch_data_from_db, get_connection
except ImportError:
    import sys

    sys.path.append(str(Path(__file__).resolve().parents[1]))
    from paths import process_paths
    from shared.db import fetch_data_from_db, get_connection


LOCAL_DIR = Path(__file__).resolve().parent
CACHE_DIR = process_paths.UK_ROLLING_CACHE_DIR
SALES_CACHE = CACHE_DIR / "uk_period_sellthru.parquet"
METADATA_CACHE = CACHE_DIR / "uk_title_metadata.parquet"
INVENTORY_CACHE = CACHE_DIR / "uk_inventory_snapshot.parquet"

REPORT_START_PERIOD = 202301
HISTORICAL_START_YEAR = 2017

METADATA_SQL = """
WITH OSD AS (
    SELECT
        tt.ean13 ISBN,
        tt.active_datevalue OSD
    FROM tmm.cb_Import_Title_Tasks tt
    WHERE
        tt.date_desc = 'On Sale Date'
        AND tt.active_datevalue IS NOT NULL
        AND tt.printingnumber = 1
        AND tt.activeind = '1'
),
uk_price AS (
    SELECT
        ukp.ean13 AS ISBN,
        ukp.finalprice UkRetail
    FROM tmm.cb_import_Title_Prices AS ukp
    WHERE currencytype_desc = 'BRPD' AND activeind = '1'
)
SELECT
    i.item_title AS ISBN,
    i.PUBLISHER_CODE AS Pub,
    i.PRODUCT_TYPE AS PT,
    i.FORMAT AS FT,
    i.PUBLISHING_GROUP AS PGRP,
    i.SHORT_TITLE AS Title,
    i.PRICE_AMOUNT AS Price,
    uk_price.UkRetail,
    CAST(i.AMORTIZATION_DATE AS date) AS PubDate,
    CAST(OSD.OSD AS date) AS OSD
FROM ebs.Item i
    LEFT JOIN OSD ON OSD.ISBN = i.item_title
    LEFT JOIN uk_price ON uk_price.ISBN = i.item_title
WHERE
    i.PUBLISHER_CODE NOT IN (
        'Benefit', 'AFO LLC', 'Glam Media', 'PQ Blackwell', 'PRINCETON', 'AMMO Books',
        'San Francisco Art Institute', 'FareArts', 'Sager', 'In Active', 'Driscolls',
        'Impossible Foods', 'Moleskine'
    )
    AND i.SHORT_TITLE IS NOT NULL
    AND i.ITEM_TITLE IS NOT NULL
    AND i.PRODUCT_TYPE IN ('BK', 'FT', 'CP', 'RP')
"""

SELLTHRU_SQL_TEMPLATE = """
SELECT
    eca.Period,
    eca.ISBN13 AS ISBN,
    SUM(eca.DelQty) UKSellThru
FROM [CBQ2].[cb].[EndCustomerData_Abrams] eca
WHERE eca.TranTypeDesc = 'Gross Sales'
{period_filter}
GROUP BY
    eca.Period,
    eca.ISBN13
"""

MAX_PERIOD_SQL = """
SELECT MAX(eca.Period) AS LatestPeriod
FROM [CBQ2].[cb].[EndCustomerData_Abrams] eca
WHERE eca.TranTypeDesc = 'Gross Sales'
"""


@dataclass(frozen=True)
class SourceFiles:
    folder: Path
    sales_files: list[Path]
    reserve_files: list[Path]
    midas_file: Path | None


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
    text = str(value).replace(",", "").replace("$", "").replace("£", "").strip()
    if text in {"", "-"}:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def period_year(period: int) -> int:
    return int(period) // 100


def period_label(period: int) -> str:
    return pd.Timestamp(year=period_year(period), month=int(period) % 100, day=1).strftime("%B %Y")


def prior_year_period(period: int) -> int:
    return int(period) - 100


def latest_output_path(latest_period: int, output_folder: Path | None = None) -> Path:
    folder = output_folder or process_paths.UK_ROLLING_OUTPUT_FOLDER
    return folder / f"{latest_period} - Rolling UK - Title Sales.xlsx"


def resolve_source_files(source_folder: Path | None = None) -> SourceFiles:
    folder = source_folder or process_paths.UK_ROLLING_SOURCE_FOLDER
    sales_files = sorted(Path(path) for path in glob.glob(str(folder / "OPPSSALX*.txt")))
    reserve_files = sorted(Path(path) for path in glob.glob(str(folder / "SMPSTKRES*.txt")))
    midas_files = sorted(folder.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
    return SourceFiles(folder, sales_files, reserve_files, midas_files[0] if midas_files else None)


def format_modified(path: Path) -> str:
    return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %I:%M %p")


def print_source_files(sources: SourceFiles) -> None:
    print("\nAbrams & Chronicle UK source files:")
    print(f"  Source folder: {sources.folder}")
    print("  Sales files:")
    if sources.sales_files:
        for path in sources.sales_files:
            print(f"    {path.name} (modified {format_modified(path)})")
    else:
        print("    (none found)")
    print("  Reserve files:")
    if sources.reserve_files:
        for path in sources.reserve_files:
            print(f"    {path.name} (modified {format_modified(path)})")
    else:
        print("    (none found)")
    print("  Midas file:")
    if sources.midas_file:
        print(f"    {sources.midas_file.name} (modified {format_modified(sources.midas_file)})")
    else:
        print("    (none found)")


def load_uk_available_snapshot(source_folder: Path | None = None) -> pd.DataFrame:
    sources = resolve_source_files(source_folder)
    if not sources.sales_files:
        raise FileNotFoundError(f"No OPPSSALX*.txt sales files found in {sources.folder}")
    if not sources.reserve_files:
        raise FileNotFoundError(f"No SMPSTKRES*.txt reserve files found in {sources.folder}")
    if sources.midas_file is None:
        raise FileNotFoundError(f"No Midas .xlsx stock file found in {sources.folder}")

    reserve_raw = pd.concat(
        (
            pd.read_csv(
                path,
                usecols=["ISBN", "RESERVED QTY"],
                encoding="unicode_escape",
                dtype={"ISBN": object},
            )
            for path in sources.reserve_files
        ),
        ignore_index=True,
    )
    reserve = pd.DataFrame(
        {
            "ISBN": reserve_raw["ISBN"].map(normalize_isbn),
            "Reserve": reserve_raw["RESERVED QTY"].map(parse_number),
        }
    )
    reserve = reserve[reserve["ISBN"].ne("")]
    reserve = reserve.groupby("ISBN", as_index=False)["Reserve"].sum()

    midas_raw = pd.read_excel(sources.midas_file, skiprows=2, engine="openpyxl", dtype=object)
    required_midas = ["Product", "Warehouse Stock", "Consignment Stock", "All Due Quantity"]
    missing = [column for column in required_midas if column not in midas_raw.columns]
    if missing:
        raise ValueError(f"Midas file is missing columns: {', '.join(missing)}")
    midas = pd.DataFrame(
        {
            "ISBN": midas_raw["Product"].map(normalize_isbn),
            "Available": midas_raw["Warehouse Stock"].map(parse_number),
            "Consignment": midas_raw["Consignment Stock"].map(parse_number),
            "BackOrder": midas_raw["All Due Quantity"].map(parse_number),
        }
    )
    midas = midas[midas["ISBN"].ne("")]
    midas = midas.groupby("ISBN", as_index=False)[["Available", "Consignment", "BackOrder"]].sum()

    combined = midas.merge(reserve, on="ISBN", how="outer")
    for column in ["Available", "Reserve", "BackOrder", "Consignment"]:
        combined[column] = pd.to_numeric(combined[column], errors="coerce").fillna(0)
    return combined[["ISBN", "Available", "Reserve", "BackOrder", "Consignment"]]


def ensure_cache_dir() -> None:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)


def fetch_metadata() -> pd.DataFrame:
    engine = get_connection()
    metadata = fetch_data_from_db(engine, METADATA_SQL)
    metadata["ISBN"] = metadata["ISBN"].map(normalize_isbn)
    metadata = metadata[metadata["ISBN"].ne("")]
    metadata["Price"] = pd.to_numeric(metadata["Price"], errors="coerce").fillna(0)
    metadata["UkRetail"] = pd.to_numeric(metadata["UkRetail"], errors="coerce").fillna(0)
    metadata["PubDate"] = pd.to_datetime(metadata["PubDate"], errors="coerce")
    metadata["OSD"] = pd.to_datetime(metadata["OSD"], errors="coerce")
    metadata = metadata.drop_duplicates(subset=["ISBN"], keep="first")
    return metadata[["Pub", "PT", "FT", "PGRP", "ISBN", "Title", "Price", "UkRetail", "PubDate", "OSD"]]


def fetch_latest_sql_period() -> int | None:
    engine = get_connection()
    df = fetch_data_from_db(engine, MAX_PERIOD_SQL)
    if df.empty or pd.isna(df.iloc[0, 0]):
        return None
    return int(df.iloc[0, 0])


def fetch_sellthru_periods(after_period: int | None = None) -> pd.DataFrame:
    period_filter = ""
    if after_period is not None:
        period_filter = f"    AND eca.Period > {int(after_period)}"
    query = SELLTHRU_SQL_TEMPLATE.format(period_filter=period_filter)
    engine = get_connection()
    sales = fetch_data_from_db(engine, query)
    if sales.empty:
        return pd.DataFrame(columns=["Period", "ISBN", "UKSellThru"])
    sales["Period"] = pd.to_numeric(sales["Period"], errors="coerce").astype("Int64")
    sales["ISBN"] = sales["ISBN"].map(normalize_isbn)
    sales["UKSellThru"] = pd.to_numeric(sales["UKSellThru"], errors="coerce").fillna(0)
    sales = sales[sales["Period"].notna() & sales["ISBN"].ne("")]
    sales["Period"] = sales["Period"].astype(int)
    return sales.groupby(["Period", "ISBN"], as_index=False)["UKSellThru"].sum()


def refresh_metadata_cache() -> pd.DataFrame:
    ensure_cache_dir()
    print("Refreshing UK title metadata from SQL...")
    metadata = fetch_metadata()
    metadata.to_parquet(METADATA_CACHE, index=False)
    print(f"Cached {len(metadata):,} UK title metadata rows.")
    return metadata


def refresh_inventory_cache() -> pd.DataFrame:
    ensure_cache_dir()
    sources = resolve_source_files()
    print_source_files(sources)
    print("\nRefreshing UK availability/reserve/backorder/consignment from source files...")
    inventory = load_uk_available_snapshot()
    inventory.to_parquet(INVENTORY_CACHE, index=False)
    print(f"Cached {len(inventory):,} UK inventory rows.")
    return inventory


def refresh_sellthru_cache(force: bool = False) -> pd.DataFrame:
    ensure_cache_dir()
    if force or not SALES_CACHE.exists():
        cached = pd.DataFrame(columns=["Period", "ISBN", "UKSellThru"])
        after_period = None
        print("Building UK period sellthru cache from SQL...")
    else:
        cached = pd.read_parquet(SALES_CACHE)
        after_period = int(cached["Period"].max()) if not cached.empty else None
        print(f"Updating UK period sellthru cache after period {after_period}...")

    new_sales = fetch_sellthru_periods(after_period)
    if new_sales.empty and not cached.empty:
        print("No new UK sellthru periods found.")
        return cached

    combined = pd.concat([cached, new_sales], ignore_index=True)
    combined["Period"] = pd.to_numeric(combined["Period"], errors="coerce").astype("Int64")
    combined["ISBN"] = combined["ISBN"].map(normalize_isbn)
    combined["UKSellThru"] = pd.to_numeric(combined["UKSellThru"], errors="coerce").fillna(0)
    combined = combined[combined["Period"].notna() & combined["ISBN"].ne("")]
    combined["Period"] = combined["Period"].astype(int)
    combined = (
        combined.groupby(["Period", "ISBN"], as_index=False)["UKSellThru"]
        .sum()
        .sort_values(["ISBN", "Period"])
    )
    combined.to_parquet(SALES_CACHE, index=False)
    print(f"Cached {len(combined):,} UK period sellthru rows through {combined['Period'].max()}.")
    return combined


def refresh_all(force_sales: bool = False) -> None:
    refresh_inventory_cache()
    refresh_metadata_cache()
    refresh_sellthru_cache(force=force_sales)


def load_or_refresh_caches() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if not INVENTORY_CACHE.exists():
        refresh_inventory_cache()
    if not METADATA_CACHE.exists():
        refresh_metadata_cache()
    if not SALES_CACHE.exists():
        refresh_sellthru_cache(force=True)
    return (
        pd.read_parquet(METADATA_CACHE),
        pd.read_parquet(INVENTORY_CACHE),
        pd.read_parquet(SALES_CACHE),
    )


def choose_periods(periods: list[int], latest_period: int) -> tuple[list[int], list[int], list[int], list[int]]:
    periods = sorted(periods)
    current_year = period_year(latest_period)
    prior_year = current_year - 1
    ytd_periods = [period for period in periods if period_year(period) == current_year and period <= latest_period]
    lytd_periods = [prior_year_period(period) for period in ytd_periods if prior_year_period(period) in periods]
    lyfy_periods = [period for period in periods if period_year(period) == prior_year]
    rolling_12 = periods[-12:]
    return rolling_12, ytd_periods, lytd_periods, lyfy_periods


def build_summary(metadata: pd.DataFrame, inventory: pd.DataFrame, sales: pd.DataFrame) -> tuple[pd.DataFrame, int, list[int], list[int]]:
    metadata = metadata.copy()
    inventory = inventory.copy()
    sales = sales.copy()

    metadata["ISBN"] = metadata["ISBN"].map(normalize_isbn)
    inventory["ISBN"] = inventory["ISBN"].map(normalize_isbn)
    sales["ISBN"] = sales["ISBN"].map(normalize_isbn)
    sales["Period"] = pd.to_numeric(sales["Period"], errors="coerce").astype("Int64")
    sales["UKSellThru"] = pd.to_numeric(sales["UKSellThru"], errors="coerce").fillna(0)
    sales = sales[sales["Period"].notna() & sales["ISBN"].ne("")]
    sales["Period"] = sales["Period"].astype(int)

    metadata_isbns = set(metadata["ISBN"])
    sales = sales[sales["ISBN"].isin(metadata_isbns)].copy()
    inventory = inventory[inventory["ISBN"].isin(metadata_isbns)].copy()
    if sales.empty:
        raise ValueError("UK sellthru cache has no rows after SQL metadata filtering.")

    latest_period = int(sales["Period"].max())
    all_periods = sorted(sales["Period"].drop_duplicates().astype(int).tolist())
    monthly_periods = [period for period in all_periods if period >= REPORT_START_PERIOD and period <= latest_period]
    historical_years = list(range(HISTORICAL_START_YEAR, period_year(REPORT_START_PERIOD)))

    pivot = sales.pivot_table(index="ISBN", columns="Period", values="UKSellThru", aggfunc="sum", fill_value=0)
    rolling_12, ytd_periods, lytd_periods, lyfy_periods = choose_periods(all_periods, latest_period)

    keys = sorted(set(sales["ISBN"]) | set(inventory["ISBN"]))
    report = pd.DataFrame({"ISBN": keys})
    report = report.merge(metadata, on="ISBN", how="inner")
    report = report.merge(inventory, on="ISBN", how="left")
    for column in ["Available", "Reserve", "BackOrder", "Consignment"]:
        report[column] = pd.to_numeric(report[column], errors="coerce").fillna(0)

    def sum_periods(isbn: str, selected_periods: list[int]) -> float:
        if isbn not in pivot.index or not selected_periods:
            return 0.0
        available = [period for period in selected_periods if period in pivot.columns]
        return float(pivot.loc[isbn, available].sum()) if available else 0.0

    report["Last12Mons"] = [sum_periods(isbn, rolling_12) for isbn in report["ISBN"]]
    report["YTD"] = [sum_periods(isbn, ytd_periods) for isbn in report["ISBN"]]
    report["LYTD"] = [sum_periods(isbn, lytd_periods) for isbn in report["ISBN"]]
    report["LYFY"] = [sum_periods(isbn, lyfy_periods) for isbn in report["ISBN"]]
    report["LTD"] = [float(pivot.loc[isbn].sum()) if isbn in pivot.index else 0.0 for isbn in report["ISBN"]]
    report["6MonAvg"] = [sum_periods(isbn, all_periods[-6:]) / 6 for isbn in report["ISBN"]]

    extra_columns = {}
    for period in sorted(monthly_periods, reverse=True):
        values = pivot[period] if period in pivot.columns else pd.Series(dtype=float)
        extra_columns[str(period)] = [values.get(isbn, 0) for isbn in report["ISBN"]]
    sales_with_year = sales.copy()
    sales_with_year["Year"] = sales_with_year["Period"].map(period_year)
    yearly = sales_with_year[sales_with_year["Year"].lt(period_year(REPORT_START_PERIOD))]
    yearly_pivot = (
        yearly.pivot_table(index="ISBN", columns="Year", values="UKSellThru", aggfunc="sum", fill_value=0)
        if not yearly.empty
        else pd.DataFrame()
    )
    for year in sorted(historical_years, reverse=True):
        values = yearly_pivot[year] if not yearly_pivot.empty and year in yearly_pivot.columns else pd.Series(dtype=float)
        extra_columns[str(year)] = [values.get(isbn, 0) for isbn in report["ISBN"]]
    if extra_columns:
        report = pd.concat([report, pd.DataFrame(extra_columns)], axis=1)

    report = report.sort_values(["YTD", "Last12Mons", "Title"], ascending=[False, False, True])
    return report, latest_period, monthly_periods, historical_years


def build_report(output_folder: Path | None = None) -> Path:
    metadata, inventory, sales = load_or_refresh_caches()
    report, latest_period, monthly_periods, historical_years = build_summary(metadata, inventory, sales)
    output_path = latest_output_path(latest_period, output_folder)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    base_columns = [
        "Pub",
        "PT",
        "FT",
        "PGRP",
        "ISBN",
        "Title",
        "Price",
        "UkRetail",
        "OSD",
        "Available",
        "Reserve",
        "BackOrder",
        "Consignment",
        "Last12Mons",
        "YTD",
        "LYTD",
        "LYFY",
        "LTD",
        "6MonAvg",
    ]
    rename_columns = {"Price": "US Price", "UkRetail": "UK Price"}
    ordered_columns = base_columns + [str(period) for period in sorted(monthly_periods, reverse=True)]
    ordered_columns.extend(str(year) for year in sorted(historical_years, reverse=True))
    output = report[[column for column in ordered_columns if column in report.columns]].rename(columns=rename_columns)

    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Rolling_Monthly")
        writer.sheets["Rolling_Monthly"] = worksheet

        title_format = workbook.add_format(
            {"font_size": 14, "bg_color": "#C4BD97", "border": 1, "align": "center", "valign": "vcenter"}
        )
        group_format = workbook.add_format(
            {"bold": True, "bg_color": "#B8CCE4", "border": 1, "align": "center", "valign": "vcenter"}
        )
        inventory_group_format = workbook.add_format(
            {"bold": True, "bg_color": "#DDD9C4", "border": 1, "align": "center", "valign": "vcenter"}
        )
        header_format = workbook.add_format(
            {"bold": True, "bg_color": "#CFE8F3", "border": 1, "align": "center", "valign": "vcenter"}
        )
        green_period_format = workbook.add_format(
            {"bold": True, "bg_color": "#D8E4BC", "border": 1, "align": "center", "valign": "vcenter"}
        )
        pink_period_format = workbook.add_format(
            {"bold": True, "bg_color": "#F2DCDB", "border": 1, "align": "center", "valign": "vcenter"}
        )
        summary_label_format = workbook.add_format(
            {"bold": True, "bg_color": "#CCC0DA", "border": 1, "align": "left", "valign": "vcenter"}
        )
        summary_number_format = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#E4DFEC",
                "border": 1,
                "align": "right",
                "num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
            }
        )
        summary_decimal_format = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#E4DFEC",
                "border": 1,
                "align": "right",
                "num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
            }
        )
        unit_format = workbook.add_format({"num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'})
        decimal_format = workbook.add_format({"num_format": '#,##0.00'})
        money_format = workbook.add_format({"num_format": '$#,##0.00'})
        date_format = workbook.add_format({"num_format": "m/d/yyyy"})
        center_across_format = workbook.add_format(
            {"bold": True, "bg_color": "#B8CCE4", "border": 1, "align": "center_across", "valign": "vcenter"}
        )
        center_across_blank_format = workbook.add_format(
            {"bold": True, "bg_color": "#B8CCE4", "border": 1, "align": "center_across", "valign": "vcenter"}
        )

        worksheet.write(1, 5, "Rolling UK Title Sales", title_format)
        worksheet.write(2, 5, f"Sellthrough as of {period_label(latest_period)}", title_format)
        worksheet.write(0, 8, "Grand Total", summary_label_format)
        worksheet.write(1, 8, "Subtotal", summary_label_format)
        worksheet.write(3, 6, "List price", center_across_format)
        worksheet.write_blank(3, 7, None, center_across_blank_format)
        worksheet.merge_range(3, 9, 3, 12, "Inventory", inventory_group_format)

        for col_idx, column in enumerate(output.columns):
            if col_idx >= len(base_columns):
                fmt = green_period_format if col_idx % 2 == 0 else pink_period_format
                worksheet.write(4, col_idx, column, fmt)
            else:
                worksheet.write(4, col_idx, column, header_format)

        for row_idx, row in enumerate(output.itertuples(index=False), start=5):
            for col_idx, value in enumerate(row):
                column = output.columns[col_idx]
                if pd.isna(value):
                    worksheet.write_blank(row_idx, col_idx, None)
                elif isinstance(value, pd.Timestamp):
                    worksheet.write_datetime(row_idx, col_idx, value.to_pydatetime(), date_format)
                elif column in {"US Price", "UK Price"}:
                    worksheet.write_number(row_idx, col_idx, float(value), money_format)
                elif column == "6MonAvg":
                    worksheet.write_number(row_idx, col_idx, float(value), unit_format)
                elif isinstance(value, (int, float)):
                    fmt = unit_format if col_idx >= 9 else None
                    worksheet.write_number(row_idx, col_idx, value, fmt)
                else:
                    worksheet.write(row_idx, col_idx, value)

        last_data_row = len(output) + 4
        col_index = {column: idx for idx, column in enumerate(output.columns)}
        for row_idx in (0, 1):
            for col_idx in range(9, len(output.columns)):
                worksheet.write_blank(row_idx, col_idx, None, summary_number_format)

        sum_columns = {
            "Available",
            "Reserve",
            "BackOrder",
            "Consignment",
            "Last12Mons",
            "YTD",
            "LYTD",
            "LYFY",
            "LTD",
        }
        sum_columns.update(str(period) for period in monthly_periods)
        sum_columns.update(str(year) for year in historical_years)
        for column in sum_columns:
            if column not in col_index:
                continue
            start_cell = xl_rowcol_to_cell(5, col_index[column])
            end_cell = xl_rowcol_to_cell(last_data_row, col_index[column])
            worksheet.write_formula(0, col_index[column], f"=SUM({start_cell}:{end_cell})", summary_number_format)
            worksheet.write_formula(1, col_index[column], f"=SUBTOTAL(9,{start_cell}:{end_cell})", summary_number_format)

        if "6MonAvg" in col_index:
            start_cell = xl_rowcol_to_cell(5, col_index["6MonAvg"])
            end_cell = xl_rowcol_to_cell(last_data_row, col_index["6MonAvg"])
            worksheet.write_formula(0, col_index["6MonAvg"], f"=SUM({start_cell}:{end_cell})", summary_decimal_format)
            worksheet.write_formula(1, col_index["6MonAvg"], f"=SUBTOTAL(9,{start_cell}:{end_cell})", summary_decimal_format)

        worksheet.freeze_panes(5, 5)
        worksheet.autofilter(4, 0, last_data_row, len(output.columns) - 1)
        worksheet.set_row(0, 20)
        worksheet.set_row(1, 20)
        worksheet.set_row(4, 28)
        widths = {
            "A": 14,
            "B": 8,
            "C": 8,
            "D": 12,
            "E": 15,
            "F": 38,
            "G": 10,
            "H": 10,
            "I": 12,
            "J": 11,
            "K": 10,
            "L": 12,
            "M": 12,
            "N": 12,
            "O": 10,
            "P": 10,
            "Q": 10,
            "R": 10,
            "S": 10,
        }
        for col, width in widths.items():
            worksheet.set_column(f"{col}:{col}", width)
        if "LTD" in output.columns:
            ltd_col = output.columns.get_loc("LTD")
            worksheet.set_column(ltd_col, ltd_col, 11)
        if len(output.columns) > len(base_columns):
            worksheet.set_column(len(base_columns), len(output.columns) - 1, 10)

    print(f"Saved Abrams & Chronicle UK rolling report: {output_path}")
    return output_path


def show_status() -> None:
    sources = resolve_source_files()
    print_source_files(sources)
    print("\nShared UK cache:")
    for label, path in {
        "Metadata": METADATA_CACHE,
        "Period sellthru": SALES_CACHE,
        "Inventory": INVENTORY_CACHE,
    }.items():
        if path.exists():
            modified = datetime.fromtimestamp(path.stat().st_mtime)
            print(f"  {label}: {path} (modified {modified:%Y-%m-%d %I:%M %p})")
        else:
            print(f"  {label}: {path} (not created yet)")
    if SALES_CACHE.exists():
        sales = pd.read_parquet(SALES_CACHE)
        if not sales.empty:
            print(f"\n  Last cached period: {int(sales['Period'].max())}")
    try:
        latest_sql = fetch_latest_sql_period()
    except Exception as exc:
        print(f"  Latest SQL period: unavailable ({exc})")
    else:
        print(f"  Latest SQL period: {latest_sql if latest_sql is not None else 'none'}")


def run_menu() -> None:
    while True:
        print("\nAbrams & Chronicle UK Rolling Reports")
        print("    1. Run full process")
        print("    2. Update sellthru cache only")
        print("    3. Rebuild sellthru cache from scratch")
        print("    4. Refresh availability/reserve/backorder/consignment")
        print("    5. Build rolling workbook from current cache")
        print("    6. Show current source/cache status")
        print("    7. Back to main menu")
        choice = input("\nChoose an option: ").strip().lower()

        try:
            if choice == "1":
                refresh_all()
                build_report()
            elif choice == "2":
                refresh_sellthru_cache()
            elif choice == "3":
                refresh_sellthru_cache(force=True)
            elif choice == "4":
                refresh_inventory_cache()
            elif choice == "5":
                build_report()
            elif choice == "6":
                show_status()
            elif choice in {"7", "b", "back", "return", "menu"}:
                return
            else:
                print("Invalid choice. Please select a valid option.")
                continue
        except Exception as exc:
            print(f"Abrams & Chronicle UK process failed: {exc}")
        input("\nPress Enter to return to the UK rolling menu...")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Abrams & Chronicle UK rolling report process")
    parser.add_argument(
        "command",
        nargs="?",
        default="menu",
        choices=["menu", "refresh", "update-sales", "rebuild-sales", "refresh-inventory", "build", "status"],
    )
    parser.add_argument("--local-output", action="store_true", help="Write report to the local output folder for testing.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output_folder = LOCAL_DIR / "output" if args.local_output else None
    if args.command == "menu":
        run_menu()
    elif args.command == "refresh":
        refresh_all()
        build_report(output_folder=output_folder)
    elif args.command == "update-sales":
        refresh_sellthru_cache()
    elif args.command == "rebuild-sales":
        refresh_sellthru_cache(force=True)
    elif args.command == "refresh-inventory":
        refresh_inventory_cache()
    elif args.command == "build":
        build_report(output_folder=output_folder)
    elif args.command == "status":
        show_status()


if __name__ == "__main__":
    main()
