import argparse
import os
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.append(str(REPO_ROOT))

from shared.bookscan_calendar import bookscan_week
from file_creation import create_rolling_report
from functions import build_column_totals, save_to_excel
from load_amazon_open_po import save_latest_amazon_po_as_pickle
from load_lastdate import (
    display_pickle_last_modified,
    lastdate_display,
    lastdate_formats,
)
from load_rolling_reports import run_query_and_save, sql_co, sql_us
from paths import (
    amazon_po_folder,
    amazon_po_pickle_file,
    amazon_rolling_folder,
    customer_orders_pickle_file,
    dp_folders,
    units_shipped_pickle_file,
)

#############


def prune_zero_history_columns(df):
    history_cols = [
        col for col in df.columns
        if isinstance(col, str) and len(col) == 10 and col.count("-") == 2
    ]

    keep_history_cols = [col for col in history_cols if pd.to_numeric(df[col], errors="coerce").fillna(0).sum() != 0]
    keep_cols = [col for col in df.columns if col not in history_cols or col in keep_history_cols]
    return df[keep_cols], keep_history_cols


def get_history_columns(df):
    return [
        col for col in df.columns
        if isinstance(col, str) and len(col) == 10 and col.count("-") == 2
    ]


def get_report_period_from_df(df):
    history_cols = get_history_columns(df)
    if not history_cols:
        raise ValueError("Could not determine report week because no history date columns were found.")

    parsed_dates = pd.to_datetime(history_cols, format="%m-%d-%Y", errors="coerce")
    valid_dates = [value for value in parsed_dates if not pd.isna(value)]
    if not valid_dates:
        raise ValueError("Could not parse any history date columns to determine the report week.")

    latest_date = max(valid_dates)
    date_formatted = latest_date.strftime("%m_%d_%Y")
    bookscan = bookscan_week(latest_date)
    return date_formatted, bookscan.week, str(bookscan.year)


def build_title_block(report_type: str, date_formatted: str) -> dict[str, object]:
    title_text = (
        "Rolling Amazon Customer Orders"
        if report_type == "Customer Orders"
        else "Rolling Amazon POS"
    )
    subtitle_date = datetime.strptime(date_formatted, "%m_%d_%Y").strftime("%B %d, %Y")
    return {
        "start_row": 1,
        "end_row": 2,
        "start_col": 6,
        "end_col": 6,
        "title": title_text,
        "subtitle": f"Week Ending: {subtitle_date}",
        "merge_cells": False,
        "align": "center",
    }


def save_reports_by_pub(
    df,
    report_type,
    week_number,
    full_year,
    date_formatted,
    dp_folders,
    summary=None,
    format_cols=None,
    decimal_cols=None,
):
    for pub, folder in dp_folders.items():
        df_pub = df[df["Pub"] == pub]
        if not df_pub.empty:
            df_pub, history_cols = prune_zero_history_columns(df_pub)
            summary_cols = ["LTD", "LY_FY", "TYTD", "LYTD", "YTD Var", "W52", "OH", "PO_Qty"]
            decimal_cols = ["Price", "OH_Avg"]
            pub_summary = build_column_totals(df_pub, history_cols + summary_cols)
            pub_format_cols = history_cols + [col for col in summary_cols if col in df_pub.columns]

            filename = f"Week {week_number:02d}-{full_year} Rolling Amazon ({date_formatted}) - {report_type}.xlsx"
            filepath = os.path.join(folder, filename)
            os.makedirs(folder, exist_ok=True)
            print(f"{'':#<40}")
            print(f"{pub.center(40)}")
            print(f"{'':#<40}")
            print(f"Saving {report_type} for {pub} to folder: {folder}")
            save_to_excel(
                df_pub,
                filepath,
                summary=pub_summary,
                format_cols=pub_format_cols,
                decimal_cols=decimal_cols,
                rolling_main_layout=True,
                pre_date_column_count=19,
                summary_label_col_idx=8,
                integer_accounting_no_symbol=True,
                title_block=build_title_block(report_type, date_formatted),
            )
            print(f"Saved {report_type} for {pub} to {filepath}")
            print()
        else:
            print(f"No data for {pub} in {report_type}")


#############


def prompt_update(filename, update_func, *args):
    display_pickle_last_modified(filename)
    choice = (
        input(f"Do you want to update {filename}? (y/n or type 'exit' to quit): ")
        .strip()
        .lower()
    )
    if choice in ["exit", "quit", "^x"]:
        print("Exiting as requested by user.")
        exit(0)
    elif choice in ["y", "yes"]:
        update_func(*args)
        print(f"{filename} updated.\n")
    else:
        print(f"Using existing {filename}.\n")


def update_all_caches(pickle_po_file):
    print("Updating all Amazon rolling report caches...")
    print("PO file status:")
    lastdate_display()
    save_latest_amazon_po_as_pickle(
        amazon_po_folder,
        pickle_po_file,
    )
    print(f"{pickle_po_file} updated.\n")

    print("Customer Orders file status:")
    run_query_and_save(sql_co, customer_orders_pickle_file, "Customer Orders")
    print(f"{customer_orders_pickle_file} updated.\n")

    print("Units Shipped file status:")
    run_query_and_save(sql_us, units_shipped_pickle_file, "Units Shipped")
    print(f"{units_shipped_pickle_file} updated.\n")


def build_parser():
    parser = argparse.ArgumentParser(description="Build Amazon rolling reports.")
    parser.add_argument(
        "--main-only",
        action="store_true",
        help="Save only the two main rolling report workbooks and skip publisher versions.",
    )
    parser.add_argument(
        "--report-only",
        action="store_true",
        help="Reuse the current local pickles and rebuild the reports without any SQL checks or refresh prompts.",
    )
    parser.add_argument(
        "--refresh-all-caches",
        action="store_true",
        help="Refresh PO, Customer Orders, and Units Shipped caches without prompting for each one.",
    )
    return parser


def main():
    args = build_parser().parse_args()
    start_time = time.time()
    pickle_po_file = amazon_po_pickle_file
    if args.report_only:
        print("Report-only mode: using existing local pickles without SQL refresh.")
        required_files = [
            pickle_po_file,
            customer_orders_pickle_file,
            units_shipped_pickle_file,
        ]
        missing_files = [filename for filename in required_files if not os.path.exists(filename)]
        if missing_files:
            raise FileNotFoundError(
                "Report-only mode requires existing local pickle files. Missing: "
                + ", ".join(str(filename) for filename in missing_files)
            )
    elif args.refresh_all_caches:
        update_all_caches(pickle_po_file)
    else:
        print("PO file status:")
        lastdate_display()
        prompt_update(
            pickle_po_file,
            save_latest_amazon_po_as_pickle,
            amazon_po_folder,
            pickle_po_file,
        )

        # --- Customer Orders ---
        pickle_file1 = customer_orders_pickle_file
        print("Customer Orders file status:")
        prompt_update(
            pickle_file1, run_query_and_save, sql_co, pickle_file1, "Customer Orders"
        )

        # --- Units Shipped ---
        pickle_file2 = units_shipped_pickle_file
        print("Units Shipped file status:")
        prompt_update(
            pickle_file2, run_query_and_save, sql_us, pickle_file2, "Units Shipped"
        )

    ###### CUSTOMER ORDERS ############################################

    # Create and save Customer Orders report
    pickle_file1 = customer_orders_pickle_file
    name1 = "Customer Orders"
    df_customer = create_rolling_report(pickle_file1, pickle_po_file)
    date_formatted, week_number, full_year = get_report_period_from_df(df_customer)
    date_cols = get_history_columns(df_customer)
    sort_col = date_cols[0] if date_cols else df_customer.columns[-1]
    df_customer = df_customer.sort_values(by=sort_col, ascending=False)

    summary_cols = ["LTD", "LY_FY", "TYTD", "LYTD", "YTD Var", "W52", "OH", "PO_Qty"]
    decimal_cols = ["Price", "OH_Avg", "6Wk Avg"]

    totals_co = build_column_totals(df_customer, date_cols + summary_cols)
    format_cols = date_cols + ["LTD", "LY_FY", "TYTD", "LYTD", "YTD Var", "W52", "OH", "PO_Qty"]

    # Saving to the main folder
    print(rf"Saving {name1} to the main Rolling Reports folder...")

    if "ISBN" in df_customer.columns:
        df_customer["ISBN"] = pd.to_numeric(df_customer["ISBN"], errors="coerce")

    save_to_excel(
        df_customer,
        os.path.join(
            amazon_rolling_folder,
            f"Week {week_number:02d}-{full_year} Rolling Amazon ({date_formatted}) - {name1}.xlsx",
        ),
        summary=totals_co,
        format_cols=format_cols,
        decimal_cols=decimal_cols,
        rolling_main_layout=True,
        pre_date_column_count=19,
        summary_label_col_idx=8,
        integer_accounting_no_symbol=True,
        title_block=build_title_block(name1, date_formatted),
    )
    # Saving to the dp folders
    if args.main_only:
        print("Skipping publisher versions for Customer Orders (--main-only).")
    else:
        print("Saving to the dp folders...")
        save_reports_by_pub(
            df_customer,
            "Customer Orders",
            week_number,
            full_year,
            date_formatted,
            dp_folders,
            summary=totals_co,
            format_cols=format_cols,
            decimal_cols=decimal_cols,
        )

    # UNITS SHIPPED #############################################
    print("#############################################")
    print("Now creating the Units Shipped report...")
    print("#############################################")
    # Create and save Units Shipped report
    pickle_file2 = units_shipped_pickle_file
    name2 = "Units Shipped"
    df_units = create_rolling_report(pickle_file2, pickle_po_file)
    units_date_formatted, units_week_number, units_full_year = get_report_period_from_df(df_units)
    date_cols = get_history_columns(df_units)
    sort_col = date_cols[0] if date_cols else df_units.columns[-1]
    df_units = df_units.sort_values(by=sort_col, ascending=False)

    summary_cols = ["LTD", "LY_FY", "TYTD", "LYTD", "YTD Var", "W52", "OH", "PO_Qty"]
    decimal_cols = ["Price", "OH_Avg", "6Wk Avg"]

    totals_us = build_column_totals(df_units, date_cols + summary_cols)
    format_cols = date_cols + ["LTD", "LY_FY", "TYTD", "LYTD", "YTD Var", "W52", "OH", "PO_Qty"]
    # Saving to the main folder
    print(rf"Saving {name2} to the main Rolling Reports folder...")

    if "ISBN" in df_units.columns:
        df_units["ISBN"] = pd.to_numeric(df_units["ISBN"], errors="coerce")

    save_to_excel(
        df_units,
        os.path.join(
            amazon_rolling_folder,
            f"Week {units_week_number:02d}-{units_full_year} Rolling Amazon ({units_date_formatted}) - {name2}.xlsx",
        ),
        summary=totals_us,
        format_cols=format_cols,
        decimal_cols=decimal_cols,
        rolling_main_layout=True,
        pre_date_column_count=19,
        summary_label_col_idx=8,
        integer_accounting_no_symbol=True,
        title_block=build_title_block(name2, units_date_formatted),
    )
    # Saving to the dp folders
    if args.main_only:
        print("Skipping publisher versions for Units Shipped (--main-only).")
    else:
        print("Saving to the dp folders...")
        save_reports_by_pub(
            df_units,
            "Units Shipped",
            units_week_number,
            units_full_year,
            units_date_formatted,
            dp_folders,
            summary=totals_us,
            format_cols=format_cols,
            decimal_cols=decimal_cols,
        )

    #################################################################

    end_time = time.time()  # End timer
    elapsed = end_time - start_time
    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)
    print(f"All done! Total runtime: {minutes} minutes, {seconds} seconds.")


if __name__ == "__main__":
    main()
