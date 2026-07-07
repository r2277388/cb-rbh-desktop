from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT_DIR = Path(__file__).resolve().parent
for path in (str(SCRIPT_DIR), str(REPO_ROOT)):
    if path in sys.path:
        sys.path.remove(path)
sys.path.insert(0, str(REPO_ROOT))
sys.path.insert(0, str(SCRIPT_DIR))

from file_creation import create_monthly_rolling_report, create_rolling_report, monthly_through_label  # noqa: E402
from functions import build_column_totals, save_to_excel  # noqa: E402
from paths import (  # noqa: E402
    amazon_po_pickle_file,
    amazon_rolling_folder,
    customer_orders_pickle_file,
    monthly_sales_parquet_file,
    units_shipped_pickle_file,
)


MONTHLY_SUMMARY_COLUMNS = ["12M", "TYTD", "LYTD", "YTD Var", "LY_FY", "Total"]


def get_month_columns(df: pd.DataFrame) -> list[str]:
    return [
        column for column in df.columns
        if isinstance(column, str) and len(column) == 6 and column.isdigit()
    ]


def latest_period(df: pd.DataFrame) -> str:
    month_columns = get_month_columns(df)
    if not month_columns:
        raise ValueError("Could not determine latest monthly period; no yyyymm columns found.")
    return max(month_columns)


def build_title_block(report_type: str, through_label: str) -> dict[str, object]:
    title_text = (
        "Rolling Amazon Customer Orders"
        if report_type == "Customer Orders"
        else "Rolling Amazon POS"
    )
    return {
        "start_row": 1,
        "end_row": 2,
        "start_col": 7,
        "end_col": 7,
        "title": title_text,
        "subtitle": through_label,
        "merge_cells": False,
        "align": "center",
    }


def save_monthly_report(df: pd.DataFrame, report_type: str, output_folder: str | Path) -> Path:
    month_columns = get_month_columns(df)
    output_path = Path(output_folder) / f"Monthly Rolling Amazon - {report_type}.xlsx"
    summary = build_column_totals(df, month_columns + MONTHLY_SUMMARY_COLUMNS)
    format_cols = month_columns + MONTHLY_SUMMARY_COLUMNS

    save_to_excel(
        df,
        output_path,
        summary=summary,
        format_cols=format_cols,
        decimal_cols=["Price"],
        rolling_main_layout=True,
        pre_date_column_count=16,
        summary_label_col_idx=9,
        integer_accounting_no_symbol=True,
        title_block=build_title_block(report_type, monthly_through_label(df)),
        weeknum_label_text="MonthNum",
    )
    return output_path


def build_reports(refresh_cache: bool = True) -> tuple[Path, Path, str]:
    if refresh_cache:
        print("Refreshing Amazon monthly sales parquet...")
        subprocess.run([sys.executable, str(SCRIPT_DIR / "monthly_sales.py")], check=True)

    if not Path(monthly_sales_parquet_file).exists():
        raise FileNotFoundError(f"Monthly sales parquet not found: {monthly_sales_parquet_file}")

    print("Building monthly Customer Orders report...")
    customer_weekly = create_rolling_report(customer_orders_pickle_file, amazon_po_pickle_file)
    customer_monthly = create_monthly_rolling_report(
        customer_weekly,
        monthly_sales_parquet_file,
        "Ordered Units",
    )
    customer_output = save_monthly_report(customer_monthly, "Customer Orders", amazon_rolling_folder)
    through_label = monthly_through_label(customer_monthly).removeprefix("Through ")
    print(f"Saved monthly Customer Orders report: {customer_output}")

    print("Building monthly Units Shipped report...")
    units_weekly = create_rolling_report(units_shipped_pickle_file, amazon_po_pickle_file)
    units_monthly = create_monthly_rolling_report(
        units_weekly,
        monthly_sales_parquet_file,
        "Shipped Units",
    )
    units_output = save_monthly_report(units_monthly, "Units Shipped", amazon_rolling_folder)
    print(f"Saved monthly Units Shipped report: {units_output}")
    return customer_output, units_output, through_label


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build standalone Amazon monthly rolling reports.")
    parser.add_argument(
        "--skip-refresh",
        action="store_true",
        help="Use the existing monthly sales parquet instead of rebuilding it first.",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    start_time = time.time()
    customer_output, units_output, through_label = build_reports(refresh_cache=not args.skip_refresh)
    elapsed = time.time() - start_time
    print(f"Monthly Amazon rolling reports complete through {through_label}.")
    print(f"Customer Orders: {customer_output}")
    print(f"Units Shipped:   {units_output}")
    print(f"Runtime: {int(elapsed // 60)} minutes, {int(elapsed % 60)} seconds.")


if __name__ == "__main__":
    main()
