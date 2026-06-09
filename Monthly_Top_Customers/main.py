from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
import re
import sys

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

try:
    from .queries import (
        national_specialty_title_sales_sql,
        rep_based_title_sales_sql,
        rep_code_lookup_sql,
        title_sales_sql,
        x_gap_title_sales_sql,
    )
except ImportError:
    from queries import (
        national_specialty_title_sales_sql,
        rep_based_title_sales_sql,
        rep_code_lookup_sql,
        title_sales_sql,
        x_gap_title_sales_sql,
    )

from shared.db import fetch_data_from_db, get_connection


BASE_OUTPUT_DIR = Path(r"G:\READTHIS\FINANCE\SALES\Top Customers")
LOCAL_CACHE_DIR = Path(__file__).resolve().parent / "cache"
CUSTOMERS = [
    "Amazon.com",
    "Barnes & Noble",
    "Canada",
    "Ingram",
    "Readerlink",
    "UK/Europe",
]
TRADE_REP_CODES = ["1007", "1024", "1032", "1043", "1117", "2002", "2006", "2030", "2300"]
SPECIALTY_REP_CODES = [
    "1016",
    "1023",
    "1025",
    "1034",
    "1046",
    "1064",
    "1088",
    "1089",
    "1103",
    "2025",
    "3313",
    "3403",
]
GENERIC_REPORTS = [
    {
        "name": "Trade Reps",
        "cache_name": "trade_reps",
        "query": lambda period: rep_based_title_sales_sql(period, TRADE_REP_CODES),
        "rep_codes": TRADE_REP_CODES,
        "total_label": "Trade Reps",
    },
    {
        "name": "Specialty Reps",
        "cache_name": "specialty_reps",
        "query": lambda period: rep_based_title_sales_sql(period, SPECIALTY_REP_CODES),
        "rep_codes": SPECIALTY_REP_CODES,
        "total_label": "Specialty Reps",
    },
    {
        "name": "National Specialty Reps",
        "cache_name": "national_specialty_reps",
        "query": national_specialty_title_sales_sql,
        "total_label": "Total Specialty",
    },
    {
        "name": "Barnes & Noble X Gap",
        "cache_name": "barnes_noble_x_gap",
        "query": lambda period: x_gap_title_sales_sql(period, "Barnes & Noble", "20", include_subject=True),
        "total_label": "Barnes & Noble",
    },
    {
        "name": "Amazon X Gap",
        "cache_name": "amazon_x_gap",
        "query": lambda period: x_gap_title_sales_sql(period, "Amazon", "6"),
        "total_label": "Amazon",
    },
]
FILE_ACCOUNT_NAMES = {
    "Amazon.com": "Amazon",
    "UK/Europe": "UK Europe",
}
DISPLAY_HEADERS = [
    "pub",
    "pt",
    "cat",
    "pgr",
    "isbn",
    "title",
    "price",
    "pub date",
    "sea",
    "$",
    "retail",
    "units",
    "disc",
    "$",
    "units",
    "$",
    "retail",
    "units",
    "disc",
    "$",
    "units",
    "$",
    "retail",
    "units",
    "disc",
    "$",
    "units",
    "$",
    "retail",
    "units",
    "disc",
    "$",
    "units",
    "$",
    "retail",
    "units",
    "disc",
    "$",
    "units",
    "$",
    "retail",
    "units",
    "disc",
    "$",
    "units",
]
GROUPS = [
    (9, 14, "TY_Month"),
    (15, 20, "LY_Month"),
    (21, 26, "TY_YTD"),
    (27, 32, "LY_YTD"),
    (33, 38, "LY_FY"),
    (39, 44, "LLY_FY"),
]
SUMMARY_COLS = [9, 10, 11, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23, 25, 26, 27, 28, 29, 31, 32, 33, 34, 35, 37, 38, 39, 40, 41, 43, 44]
DISCOUNT_COLS = [12, 18, 24, 30, 36, 42]
DATE_COLS = [7]
HEADER_BLUE = "#1F497D"
WHITE = "#FFFFFF"
GROUP_BORDER_COLOR = "#BFBFBF"
CB_GROUP_FILL = "#60497A"
DEFAULT_FONT_SIZE = 11
NATIONAL_SPECIALTY_LABELS = {
    "Anthro": "Anthropologie",
    "CC": "Calendar Club",
    "Cost Plus": "Cost Plus World Market",
    "C&B": "Crate & Barrel",
    "DB": "Dick Blick",
    "FedEx": "Fed Ex Office",
    "Fran": "Francesca's",
    "Fuego": "Fuego",
    "Hobby": "Hobby Lobby",
    "Pot": "Potpourri Group",
    "PBK": "Pottery Barn Kids",
    "Spencer": "Spencer Gift",
    "Sub Box": "Subscription Boxes",
    "Container": "The Container Store",
    "UG": "Uncommon Goods",
    "Urban": "Urban Outfitters",
    "West Bway": "West Broadway",
    "Indie Spec": "Indie Specialty",
    "Faire": "Faire",
    "Target NR": "Target",
}


def default_period() -> str:
    today = datetime.today()
    year = today.year
    month = today.month - 1
    if month == 0:
        month = 12
        year -= 1
    return f"{year}{month:02d}"


def validate_period(period: str) -> str:
    normalized = period.strip()
    if not re.fullmatch(r"\d{6}", normalized):
        raise ValueError("Period must be in YYYYMM format, for example 202604.")
    datetime.strptime(normalized, "%Y%m")
    return normalized


def prompt_for_period() -> str:
    fallback = default_period()
    while True:
        entered = input(f"What period are we creating this report for? YYYYMM [{fallback}]: ").strip()
        if not entered:
            return fallback
        try:
            return validate_period(entered)
        except ValueError as exc:
            print(exc)


def month_label(period: str, year_delta: int = 0) -> str:
    dt = datetime.strptime(period, "%Y%m")
    return dt.replace(year=dt.year + year_delta).strftime("%B %Y")


def ytd_label(period: str, year_delta: int = 0) -> str:
    year = int(period[:4]) + year_delta
    return f"YTD {year}"


def fy_label(period: str, year_delta: int = 0) -> str:
    year = int(period[:4]) + year_delta
    return f"FY {year}"


def period_folder_name(period: str) -> str:
    dt = datetime.strptime(period, "%Y%m")
    return dt.strftime("%m%Y")


def cache_path(period: str) -> Path:
    return LOCAL_CACHE_DIR / f"title_sales_{period}.parquet"


def report_cache_path(period: str, cache_name: str) -> Path:
    return LOCAL_CACHE_DIR / f"{cache_name}_{period}.parquet"


def load_or_fetch_data(period: str, refresh: bool = False) -> pd.DataFrame:
    LOCAL_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    path = cache_path(period)
    if path.exists() and not refresh:
        print(f"Using cached SQL results: {path}")
        return pd.read_parquet(path)

    print("Running SQL query. This can take a couple of minutes...")
    engine = get_connection()
    df = fetch_data_from_db(engine, title_sales_sql(period))
    df.to_parquet(path, index=False)
    df.to_csv(path.with_suffix(".csv"), index=False)
    print(f"Saved SQL results cache: {path}")
    return df


def load_or_fetch_report_data(period: str, cache_name: str, query: str, refresh: bool = False) -> pd.DataFrame:
    LOCAL_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    path = report_cache_path(period, cache_name)
    if path.exists() and not refresh:
        print(f"Using cached SQL results: {path}")
        return pd.read_parquet(path)

    print(f"Running SQL query for {cache_name}. This can take a couple of minutes...")
    engine = get_connection()
    df = fetch_data_from_db(engine, query)
    df.to_parquet(path, index=False)
    df.to_csv(path.with_suffix(".csv"), index=False)
    print(f"Saved SQL results cache: {path}")
    return df


def fetch_rep_code_lookup(rep_codes: list[str], refresh: bool = False) -> pd.DataFrame:
    cache_name = "rep_codes_" + "_".join(rep_codes)
    path = report_cache_path("lookup", cache_name)
    if path.exists() and not refresh:
        return pd.read_parquet(path)

    engine = get_connection()
    df = fetch_data_from_db(engine, rep_code_lookup_sql(rep_codes))
    LOCAL_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    df.to_parquet(path, index=False)
    return df


def prepare_output_dir(period: str, assume_yes: bool = False) -> Path:
    year_dir = BASE_OUTPUT_DIR / period[:4]
    if not year_dir.exists():
        if not assume_yes:
            answer = input(f"Create new year folder {year_dir}? [y/N]: ").strip().lower()
            if answer not in {"y", "yes"}:
                raise RuntimeError(f"Year folder does not exist: {year_dir}")
        year_dir.mkdir(parents=True, exist_ok=True)

    month_dir = year_dir / period_folder_name(period)
    month_dir.mkdir(parents=True, exist_ok=True)
    return month_dir


def output_file_name(period: str, customer: str) -> str:
    account_name = FILE_ACCOUNT_NAMES.get(customer, customer)
    account_name = re.sub(r'[<>:"/\\|?*]', " ", account_name)
    account_name = re.sub(r"\s+", " ", account_name).strip()
    return f"{period_folder_name(period)} Title Sales - {account_name}.xlsx"


def generic_output_file_name(period: str, report_name: str) -> str:
    safe_name = re.sub(r'[<>:"/\\|?*]', " ", report_name)
    safe_name = re.sub(r"\s+", " ", safe_name).strip()
    return f"{period_folder_name(period)} Title Sales - {safe_name}.xlsx"


def normalize_data_types(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    if "pub date" in result.columns:
        result["pub date"] = pd.to_datetime(result["pub date"], errors="coerce")
    for column in result.columns[10:]:
        result[column] = pd.to_numeric(result[column], errors="coerce")
    return result


def write_customer_workbook(df: pd.DataFrame, period: str, customer: str, output_path: Path) -> None:
    report_df = df[df["type"].eq(customer)].copy()
    report_df = normalize_data_types(report_df)
    report_df = report_df.drop(columns=["type"])
    report_df.columns = DISPLAY_HEADERS
    header_row = 4
    data_start_row = 5
    first_data_excel_row = data_start_row + 1
    last_excel_row = len(report_df) + data_start_row

    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        report_df.to_excel(writer, sheet_name="Title Sales", startrow=header_row, index=False)
        workbook = writer.book
        worksheet = writer.sheets["Title Sales"]
        sql_worksheet = workbook.add_worksheet("SQL")
        writer.sheets["SQL"] = sql_worksheet

        title_fmt = workbook.add_format(
            {
                "bg_color": HEADER_BLUE,
                "font_color": WHITE,
                "align": "center_across",
                "valign": "vcenter",
                "font_size": DEFAULT_FONT_SIZE,
            }
        )
        total_label_fmt = workbook.add_format({"bg_color": "#B7DEE8", "bold": True, "align": "center"})
        subtotal_label_fmt = workbook.add_format({"bg_color": "#DAEEF3", "bold": True, "align": "center"})
        total_num_fmt = workbook.add_format({"bg_color": "#B7DEE8", "num_format": '#,##0;(#,##0);"-"', "bold": True})
        subtotal_num_fmt = workbook.add_format({"bg_color": "#DAEEF3", "num_format": '#,##0;(#,##0);"-"', "bold": True})
        total_pct_fmt = workbook.add_format({"bg_color": "#B7DEE8", "num_format": '0.0%;(0.0%);"-"', "bold": True})
        subtotal_pct_fmt = workbook.add_format({"bg_color": "#DAEEF3", "num_format": '0.0%;(0.0%);"-"', "bold": True})
        group_fmt = workbook.add_format(
            {
                "align": "center_across",
                "valign": "vcenter",
                "bg_color": HEADER_BLUE,
                "font_color": WHITE,
                "border": 1,
                "border_color": GROUP_BORDER_COLOR,
                "font_size": DEFAULT_FONT_SIZE,
            }
        )
        sub_group_fmt = workbook.add_format(
            {
                "align": "center_across",
                "valign": "vcenter",
                "bg_color": HEADER_BLUE,
                "font_color": WHITE,
                "border": 1,
                "border_color": GROUP_BORDER_COLOR,
            }
        )
        alt_group_fmt = workbook.add_format(
            {
                "align": "center_across",
                "valign": "vcenter",
                "bg_color": CB_GROUP_FILL,
                "font_color": WHITE,
                "border": 1,
                "border_color": GROUP_BORDER_COLOR,
                "font_size": DEFAULT_FONT_SIZE,
            }
        )
        alt_sub_group_fmt = workbook.add_format(
            {
                "align": "center_across",
                "valign": "vcenter",
                "bg_color": CB_GROUP_FILL,
                "font_color": WHITE,
                "border": 1,
                "border_color": GROUP_BORDER_COLOR,
            }
        )
        header_fmt = workbook.add_format(
            {
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "bg_color": HEADER_BLUE,
                "font_color": WHITE,
                "border": 1,
                "border_color": GROUP_BORDER_COLOR,
                "font_size": DEFAULT_FONT_SIZE,
            }
        )
        alt_header_fmt = workbook.add_format(
            {
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "bg_color": CB_GROUP_FILL,
                "font_color": WHITE,
                "border": 1,
                "border_color": GROUP_BORDER_COLOR,
                "font_size": DEFAULT_FONT_SIZE,
            }
        )
        text_fmt = workbook.add_format({"font_size": DEFAULT_FONT_SIZE})
        integer_fmt = workbook.add_format({"num_format": '#,##0;(#,##0);"-"', "font_size": DEFAULT_FONT_SIZE})
        money_fmt = workbook.add_format({"num_format": '#,##0;(#,##0);"-"', "font_size": DEFAULT_FONT_SIZE})
        price_fmt = workbook.add_format({"num_format": "0.00", "font_size": DEFAULT_FONT_SIZE})
        pct_fmt = workbook.add_format({"num_format": '0.0%;(0.0%);"-"', "font_size": DEFAULT_FONT_SIZE})
        date_fmt = workbook.add_format({"num_format": "m/d/yyyy", "font_size": DEFAULT_FONT_SIZE})
        boundary_text_fmt = workbook.add_format({"left": 1, "left_color": GROUP_BORDER_COLOR, "font_size": DEFAULT_FONT_SIZE})
        boundary_integer_fmt = workbook.add_format(
            {
                "num_format": '#,##0;(#,##0);"-"',
                "left": 1,
                "left_color": GROUP_BORDER_COLOR,
                "font_size": DEFAULT_FONT_SIZE,
            }
        )
        boundary_pct_fmt = workbook.add_format(
            {
                "num_format": '0.0%;(0.0%);"-"',
                "left": 1,
                "left_color": GROUP_BORDER_COLOR,
                "font_size": DEFAULT_FONT_SIZE,
            }
        )
        sql_fmt = workbook.add_format({"font_name": "Consolas", "font_size": 10, "text_wrap": False})

        for row, value in [(0, customer), (1, month_label(period))]:
            worksheet.write(row, 0, value, title_fmt)
            for col in range(1, 7):
                worksheet.write_blank(row, col, None, title_fmt)

        worksheet.write(0, 8, "Total", total_label_fmt)
        worksheet.write(1, 8, "Subtotal", subtotal_label_fmt)

        group_labels = [
            month_label(period),
            month_label(period, -1),
            ytd_label(period),
            ytd_label(period, -1),
            fy_label(period, -1),
            fy_label(period, -2),
        ]
        for group_idx, ((start_col, end_col, _), label) in enumerate(zip(GROUPS, group_labels)):
            period_group_fmt = group_fmt if group_idx % 2 == 0 else alt_group_fmt
            period_sub_group_fmt = sub_group_fmt if group_idx % 2 == 0 else alt_sub_group_fmt

            worksheet.write(2, start_col, label, period_group_fmt)
            for col in range(start_col + 1, end_col + 1):
                worksheet.write_blank(2, col, None, period_group_fmt)

            worksheet.write(3, start_col, "Sales", period_sub_group_fmt)
            for col in range(start_col + 1, start_col + 4):
                worksheet.write_blank(3, col, None, period_sub_group_fmt)
            worksheet.write(3, start_col + 4, "Returns", period_sub_group_fmt)
            worksheet.write_blank(3, start_col + 5, None, period_sub_group_fmt)

        for col_idx, column_name in enumerate(report_df.columns):
            column_group_idx = next(
                (group_idx for group_idx, (start_col, end_col, _) in enumerate(GROUPS) if start_col <= col_idx <= end_col),
                None,
            )
            column_header_fmt = alt_header_fmt if column_group_idx is not None and column_group_idx % 2 else header_fmt
            worksheet.write(header_row, col_idx, column_name, column_header_fmt)

        for col_idx in SUMMARY_COLS:
            col_letter = xl_col(col_idx)
            if len(report_df) == 0:
                worksheet.write_blank(0, col_idx, None, total_num_fmt)
                worksheet.write_blank(1, col_idx, None, subtotal_num_fmt)
                continue
            worksheet.write_formula(
                0,
                col_idx,
                f"=SUM({col_letter}{first_data_excel_row}:{col_letter}{last_excel_row})",
                total_num_fmt,
            )
            worksheet.write_formula(
                1,
                col_idx,
                f"=SUBTOTAL(9,{col_letter}{first_data_excel_row}:{col_letter}{last_excel_row})",
                subtotal_num_fmt,
            )

        for col_idx in DISCOUNT_COLS:
            sales_col = xl_col(col_idx - 3)
            retail_col = xl_col(col_idx - 2)
            letter = xl_col(col_idx)
            worksheet.write_formula(0, col_idx, f'=IF({retail_col}1=0,"-",1-({sales_col}1/{retail_col}1))', total_pct_fmt)
            worksheet.write_formula(1, col_idx, f'=IF({retail_col}2=0,"-",1-({sales_col}2/{retail_col}2))', subtotal_pct_fmt)

        width_formats = {idx: text_fmt for idx in range(len(report_df.columns))}
        for col_idx in SUMMARY_COLS:
            width_formats[col_idx] = integer_fmt
        for col_idx in DISCOUNT_COLS:
            width_formats[col_idx] = pct_fmt
        for col_idx in DATE_COLS:
            width_formats[col_idx] = date_fmt
        if "price" in report_df.columns:
            width_formats[report_df.columns.get_loc("price")] = price_fmt

        boundary_start_cols = {start_col for start_col, _, _ in GROUPS[1:]}
        for col_idx, width in enumerate(auto_column_widths(report_df)):
            col_format = width_formats.get(col_idx, text_fmt)
            if col_idx in boundary_start_cols:
                if col_idx in DISCOUNT_COLS:
                    col_format = boundary_pct_fmt
                elif col_idx in SUMMARY_COLS:
                    col_format = boundary_integer_fmt
                else:
                    col_format = boundary_text_fmt
            worksheet.set_column(col_idx, col_idx, width, col_format)

        worksheet.freeze_panes(data_start_row, 0)
        worksheet.autofilter(header_row, 0, max(last_excel_row - 1, header_row), len(report_df.columns) - 1)
        worksheet.hide_gridlines(2)
        worksheet.set_zoom(90)

        sql_worksheet.set_column("A:A", 160, sql_fmt)
        for row_idx, line in enumerate(title_sales_sql(period).strip().splitlines()):
            sql_worksheet.write_string(row_idx, 0, line.rstrip(), sql_fmt)


def write_generic_workbook(
    df: pd.DataFrame,
    period: str,
    report_name: str,
    output_path: Path,
    sql_query: str,
    rep_lookup: pd.DataFrame | None = None,
) -> None:
    report_df = normalize_generic_data_types(df)
    is_rep_code_report = rep_lookup is not None and not rep_lookup.empty
    header_row = 4 if is_rep_code_report else 3
    data_start_row = 4
    if is_rep_code_report:
        data_start_row = 5
    first_data_excel_row = data_start_row + 1
    last_excel_row = len(report_df) + data_start_row
    metadata_col_count = metadata_column_count(report_df)
    summary_label_col = max(metadata_col_count - 1, 0)
    summary_cols = numeric_summary_columns(report_df)
    pct_cols = percent_columns(report_df)

    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        report_df.to_excel(writer, sheet_name="Title Sales", startrow=header_row, index=False)
        workbook = writer.book
        worksheet = writer.sheets["Title Sales"]
        sql_worksheet = workbook.add_worksheet("SQL")
        writer.sheets["SQL"] = sql_worksheet

        title_fmt = workbook.add_format(
            {
                "bg_color": HEADER_BLUE,
                "font_color": WHITE,
                "align": "center_across",
                "valign": "vcenter",
                "font_size": DEFAULT_FONT_SIZE,
            }
        )
        group_fmt = workbook.add_format(
            {
                "bg_color": HEADER_BLUE,
                "font_color": WHITE,
                "align": "center_across",
                "valign": "vcenter",
                "border": 1,
                "border_color": GROUP_BORDER_COLOR,
                "font_size": DEFAULT_FONT_SIZE,
            }
        )
        cb_group_fmt = workbook.add_format(
            {
                "bg_color": CB_GROUP_FILL,
                "font_color": WHITE,
                "align": "center_across",
                "valign": "vcenter",
                "border": 1,
                "border_color": GROUP_BORDER_COLOR,
                "font_size": DEFAULT_FONT_SIZE,
            }
        )
        purple_header_fmt = workbook.add_format(
            {
                "bold": True,
                "bg_color": CB_GROUP_FILL,
                "font_color": WHITE,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": GROUP_BORDER_COLOR,
                "font_size": DEFAULT_FONT_SIZE,
            }
        )
        header_fmt = workbook.add_format(
            {
                "bold": True,
                "bg_color": HEADER_BLUE,
                "font_color": WHITE,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": GROUP_BORDER_COLOR,
                "font_size": DEFAULT_FONT_SIZE,
            }
        )
        total_label_fmt = workbook.add_format({"bg_color": "#B7DEE8", "bold": True, "align": "center", "font_size": DEFAULT_FONT_SIZE})
        subtotal_label_fmt = workbook.add_format({"bg_color": "#DAEEF3", "bold": True, "align": "center", "font_size": DEFAULT_FONT_SIZE})
        total_num_fmt = workbook.add_format({"bg_color": "#B7DEE8", "num_format": '#,##0;(#,##0);"-"', "bold": True, "font_size": DEFAULT_FONT_SIZE})
        subtotal_num_fmt = workbook.add_format({"bg_color": "#DAEEF3", "num_format": '#,##0;(#,##0);"-"', "bold": True, "font_size": DEFAULT_FONT_SIZE})
        total_pct_fmt = workbook.add_format({"bg_color": "#B7DEE8", "num_format": '0.0%;(0.0%);"-"', "bold": True, "font_size": DEFAULT_FONT_SIZE})
        subtotal_pct_fmt = workbook.add_format({"bg_color": "#DAEEF3", "num_format": '0.0%;(0.0%);"-"', "bold": True, "font_size": DEFAULT_FONT_SIZE})
        text_fmt = workbook.add_format({"font_size": DEFAULT_FONT_SIZE})
        integer_fmt = workbook.add_format({"num_format": '#,##0;(#,##0);"-"', "font_size": DEFAULT_FONT_SIZE})
        price_fmt = workbook.add_format({"num_format": "0.00", "font_size": DEFAULT_FONT_SIZE})
        pct_fmt = workbook.add_format({"num_format": '0.0%;(0.0%);"-"', "font_size": DEFAULT_FONT_SIZE})
        date_fmt = workbook.add_format({"num_format": "m/d/yyyy", "font_size": DEFAULT_FONT_SIZE})
        boundary_text_fmt = workbook.add_format({"left": 1, "left_color": GROUP_BORDER_COLOR, "font_size": DEFAULT_FONT_SIZE})
        boundary_integer_fmt = workbook.add_format(
            {"num_format": '#,##0;(#,##0);"-"', "left": 1, "left_color": GROUP_BORDER_COLOR, "font_size": DEFAULT_FONT_SIZE}
        )
        boundary_pct_fmt = workbook.add_format(
            {"num_format": '0.0%;(0.0%);"-"', "left": 1, "left_color": GROUP_BORDER_COLOR, "font_size": DEFAULT_FONT_SIZE}
        )
        pink_pct_fmt = workbook.add_format({"bg_color": "#FFC7CE", "font_size": DEFAULT_FONT_SIZE})
        sql_fmt = workbook.add_format({"font_name": "Consolas", "font_size": 10, "text_wrap": False})

        for row, value in [(0, report_name), (1, ytd_label(period))]:
            worksheet.write(row, 0, value, title_fmt)
            for col in range(1, min(metadata_col_count, 7)):
                worksheet.write_blank(row, col, None, title_fmt)

        worksheet.write(0, summary_label_col, "Total", total_label_fmt)
        worksheet.write(1, summary_label_col, "Subtotal", subtotal_label_fmt)
        is_x_gap_report = "X Gap" in report_name
        is_national_specialty_report = report_name == "National Specialty Reps"
        rep_name_by_code = rep_lookup_map(rep_lookup) if is_rep_code_report else {}
        write_generic_group_headers(
            worksheet,
            report_df,
            metadata_col_count,
            group_fmt,
            cb_group_fmt if (is_x_gap_report or is_rep_code_report or is_national_specialty_report) else None,
            x_gap=is_x_gap_report,
            alternating=bool(is_rep_code_report or is_national_specialty_report),
            label_map=NATIONAL_SPECIALTY_LABELS if is_national_specialty_report else None,
        )
        if is_rep_code_report:
            blank_total_group_header_for_rep_reports(
                worksheet,
                report_df,
                metadata_col_count,
                group_fmt,
            )
            write_rep_name_row(
                worksheet,
                report_df,
                metadata_col_count,
                rep_name_by_code,
                group_fmt,
                cb_group_fmt,
            )

        header_group_labels = (
            x_gap_group_labels(report_df, metadata_col_count)
            if is_x_gap_report
            else [group_label_for_column(column) for column in report_df.columns]
        )
        rep_alternating_formats = (
            alternating_group_formats(header_group_labels, group_fmt, cb_group_fmt)
            if (is_rep_code_report or is_national_specialty_report)
            else {}
        )
        for col_idx, column_name in enumerate(report_df.columns):
            if (is_rep_code_report or is_national_specialty_report) and col_idx >= metadata_col_count:
                header_format = (
                    purple_header_fmt
                    if rep_alternating_formats.get(header_group_labels[col_idx]) is cb_group_fmt
                    else header_fmt
                )
            else:
                header_format = (
                    cb_group_fmt
                    if is_x_gap_report and header_group_labels[col_idx].startswith("CB")
                    else header_fmt
                )
            worksheet.write(
                header_row,
                col_idx,
                display_header_name(column_name, report_name=report_name if is_x_gap_report else None),
                header_format,
            )

        for col_idx in summary_cols:
            col_letter = xl_col(col_idx)
            total_fmt = total_pct_fmt if col_idx in pct_cols else total_num_fmt
            subtotal_fmt = subtotal_pct_fmt if col_idx in pct_cols else subtotal_num_fmt
            if len(report_df) == 0:
                worksheet.write_blank(0, col_idx, None, total_fmt)
                worksheet.write_blank(1, col_idx, None, subtotal_fmt)
                continue
            if col_idx in pct_cols:
                numerator_col = xl_col(col_idx - 2)
                denominator_idx = percent_denominator_column_index(report_df, col_idx)
                denominator_col = xl_col(denominator_idx)
                worksheet.write_formula(
                    0,
                    col_idx,
                    f'=IFERROR({numerator_col}1/{denominator_col}1,"-")',
                    total_fmt,
                )
                worksheet.write_formula(
                    1,
                    col_idx,
                    f'=IFERROR({numerator_col}2/{denominator_col}2,"-")',
                    subtotal_fmt,
                )
            else:
                worksheet.write_formula(
                    0,
                    col_idx,
                    f"=SUM({col_letter}{first_data_excel_row}:{col_letter}{last_excel_row})",
                    total_fmt,
                )
                worksheet.write_formula(
                    1,
                    col_idx,
                    f"=SUBTOTAL(9,{col_letter}{first_data_excel_row}:{col_letter}{last_excel_row})",
                    subtotal_fmt,
                )

        width_formats = {idx: text_fmt for idx in range(len(report_df.columns))}
        for col_idx in summary_cols:
            width_formats[col_idx] = pct_fmt if col_idx in pct_cols else integer_fmt
        for col_idx, column in enumerate(report_df.columns):
            if normalized_col_name(column) in {"pub date", "ship"}:
                width_formats[col_idx] = date_fmt
            elif normalized_col_name(column) == "price":
                width_formats[col_idx] = price_fmt

        boundary_start_cols = generic_boundary_start_columns(
            report_df,
            metadata_col_count,
            x_gap=is_x_gap_report,
        )
        for col_idx, width in enumerate(auto_column_widths(report_df)):
            col_format = width_formats.get(col_idx, text_fmt)
            if col_idx in boundary_start_cols:
                if col_idx in pct_cols:
                    col_format = boundary_pct_fmt
                elif col_idx in summary_cols:
                    col_format = boundary_integer_fmt
                else:
                    col_format = boundary_text_fmt
            worksheet.set_column(col_idx, col_idx, width, col_format)

        worksheet.freeze_panes(data_start_row, 0)
        worksheet.autofilter(header_row, 0, max(last_excel_row - 1, header_row), len(report_df.columns) - 1)
        worksheet.hide_gridlines(2)
        worksheet.set_zoom(90)
        if is_x_gap_report:
            for col_idx in pct_cols:
                col_letter = xl_col(col_idx)
                first_cell = f"{col_letter}{first_data_excel_row}"
                worksheet.conditional_format(
                    f"{first_cell}:{col_letter}{last_excel_row}",
                    {
                        "type": "formula",
                        "criteria": f"=AND(ISNUMBER({first_cell}),{first_cell}<=10%)",
                        "format": pink_pct_fmt,
                    },
                )

        sql_worksheet.set_column("A:A", 160, sql_fmt)
        for row_idx, line in enumerate(sql_query.strip().splitlines()):
            sql_worksheet.write_string(row_idx, 0, line.rstrip(), sql_fmt)

        if rep_lookup is not None and not rep_lookup.empty:
            rep_lookup.to_excel(writer, sheet_name="Rep Codes", index=False)
            rep_ws = writer.sheets["Rep Codes"]
            rep_ws.set_column("A:A", 14)
            rep_ws.set_column("B:B", 40)


def normalize_generic_data_types(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result.columns = [str(column).strip() for column in result.columns]
    rename_map = {
        "PUBLISHER_CODE": "pub",
        "PRODUCT_TYPE": "pt",
        "REPORTING_CATEGORY": "cat",
        "PUBLISHING_GROUP": "pgr",
        "ITEM_TITLE": "isbn",
        "SHORT_TITLE": "title",
        "PRICE_AMOUNT": "price",
        "AMORTIZATION_DATE": "pub date",
        "SEASON": "sea",
    }
    result = result.rename(columns={key: value for key, value in rename_map.items() if key in result.columns})
    for column in result.columns:
        name = normalized_col_name(column)
        if name in {"pub date", "ship"}:
            result[column] = pd.to_datetime(result[column], errors="coerce")
        elif column in numeric_like_columns(result):
            result[column] = pd.to_numeric(result[column], errors="coerce")
    return result


def metadata_column_count(df: pd.DataFrame) -> int:
    for idx, column in enumerate(df.columns):
        name = normalized_col_name(column)
        if "$" in name or "units" in name or name.endswith("%") or name == "$":
            return idx
    return min(9, len(df.columns))


def normalized_col_name(column: object) -> str:
    return str(column).strip().lower()


def numeric_like_columns(df: pd.DataFrame) -> set[str]:
    numeric_cols = set()
    for column in df.columns:
        name = normalized_col_name(column)
        if (
            "$" in name
            or "units" in name
            or "%" in name
            or name in {"price", "$", "retail"}
        ):
            numeric_cols.add(column)
    return numeric_cols


def numeric_summary_columns(df: pd.DataFrame) -> list[int]:
    numeric_cols = numeric_like_columns(df)
    return [idx for idx, column in enumerate(df.columns) if column in numeric_cols and normalized_col_name(column) != "price"]


def percent_columns(df: pd.DataFrame) -> set[int]:
    return {idx for idx, column in enumerate(df.columns) if "%" in normalized_col_name(column)}


def first_dollar_column_index(df: pd.DataFrame) -> int:
    for idx, column in enumerate(df.columns):
        name = normalized_col_name(column)
        if "$" in name and normalized_col_name(column) != "price":
            return idx
    return metadata_column_count(df)


def percent_denominator_column_index(df: pd.DataFrame, percent_col_idx: int) -> int:
    if percent_col_idx + 1 < len(df.columns) and normalized_col_name(df.columns[percent_col_idx + 1]).startswith("cb "):
        return percent_col_idx + 1
    return first_dollar_column_index(df)


def display_header_name(column: object, report_name: str | None = None) -> str:
    name = str(column)
    replacements = {
        "pub date": "ship",
        "Total Specialty $": "$",
        "Total Units": "units",
    }
    if name in replacements:
        return replacements[name]
    if name.endswith(" $"):
        return "$"
    if name.endswith(" Units") or name.endswith(" units"):
        return "units"
    if "%" in name:
        if report_name == "Amazon X Gap":
            return "Amaz%"
        if report_name == "Barnes & Noble X Gap":
            return "B&N%"
        return "% tot"
    return name


def write_generic_group_headers(
    worksheet,
    df: pd.DataFrame,
    metadata_col_count: int,
    group_fmt,
    cb_group_fmt=None,
    x_gap: bool = False,
    alternating: bool = False,
    label_map: dict[str, str] | None = None,
) -> None:
    if len(df.columns) <= metadata_col_count:
        return
    labels = (
        x_gap_group_labels(df, metadata_col_count)
        if x_gap
        else [group_label_for_column(column) for column in df.columns]
    )
    first_label = labels[metadata_col_count]
    alternating_formats = (
        alternating_group_formats(labels, group_fmt, cb_group_fmt)
        if alternating and cb_group_fmt is not None
        else {}
    )
    first_format = alternating_formats.get(first_label)
    if first_format is None:
        first_format = cb_group_fmt if cb_group_fmt is not None and first_label.startswith("CB") else group_fmt
    worksheet.write(2, metadata_col_count, display_group_label(first_label, label_map), first_format)
    for col_idx in range(metadata_col_count + 1, len(df.columns)):
        label = labels[col_idx]
        previous = labels[col_idx - 1]
        cell_format = alternating_formats.get(label)
        if cell_format is None:
            cell_format = cb_group_fmt if cb_group_fmt is not None and label.startswith("CB") else group_fmt
        if label != previous:
            worksheet.write(2, col_idx, display_group_label(label, label_map), cell_format)
        else:
            worksheet.write_blank(2, col_idx, None, cell_format)


def display_group_label(label: str, label_map: dict[str, str] | None = None) -> str:
    if label_map is None:
        return label
    return label_map.get(label, label)


def write_rep_name_row(
    worksheet,
    df: pd.DataFrame,
    metadata_col_count: int,
    rep_name_by_code: dict[str, str],
    group_fmt,
    purple_group_fmt,
) -> None:
    labels = [group_label_for_column(column) for column in df.columns]
    alternating_formats = alternating_group_formats(labels, group_fmt, purple_group_fmt)
    for col_idx in range(metadata_col_count, len(df.columns)):
        label = labels[col_idx]
        cell_format = alternating_formats.get(label, group_fmt)
        previous = labels[col_idx - 1] if col_idx > metadata_col_count else None
        value = rep_name_by_code.get(label, label if label == "TOTAL" else "")
        if label != previous:
            worksheet.write(3, col_idx, value, cell_format)
        else:
            worksheet.write_blank(3, col_idx, None, cell_format)


def blank_total_group_header_for_rep_reports(worksheet, df: pd.DataFrame, metadata_col_count: int, group_fmt) -> None:
    labels = [group_label_for_column(column) for column in df.columns]
    for col_idx in range(metadata_col_count, len(df.columns)):
        if labels[col_idx] != "TOTAL":
            continue
        worksheet.write_blank(2, col_idx, None, group_fmt)


def rep_lookup_map(rep_lookup: pd.DataFrame | None) -> dict[str, str]:
    if rep_lookup is None or rep_lookup.empty:
        return {}
    return {
        str(row["rep_number"]).strip(): str(row["rep_name"]).strip()
        for _, row in rep_lookup.iterrows()
        if pd.notna(row.get("rep_number")) and pd.notna(row.get("rep_name"))
    }


def alternating_group_formats(labels: list[str], first_fmt, second_fmt) -> dict[str, object]:
    formats = {}
    ordered_labels = []
    for label in labels:
        if not label or label in ordered_labels:
            continue
        ordered_labels.append(label)
    for idx, label in enumerate(ordered_labels):
        formats[label] = first_fmt if idx % 2 == 0 else second_fmt
    return formats


def group_label_for_column(column: object) -> str:
    name = str(column).strip()
    if name in {"$", "units", "Total Specialty $", "Total Units"}:
        return "TOTAL"
    if name.startswith("CB "):
        return "CB"
    if " YTD" in name:
        return name.split(" YTD", 1)[0] + " YTD"
    if " FY" in name:
        return name.split(" FY", 1)[0] + " FY"
    if name.endswith(" $"):
        return name[:-2]
    if name.endswith(" Units"):
        return name[:-6]
    if name.endswith(" units"):
        return name[:-6]
    if "%" in name:
        return name[:-2]
    return ""


def x_gap_group_labels(df: pd.DataFrame, metadata_col_count: int) -> list[str]:
    labels = ["" for _ in df.columns]
    value_cols = list(range(metadata_col_count, len(df.columns)))
    group_sizes = [3, 2, 3, 2]
    cursor = 0
    for group_index, group_size in enumerate(group_sizes):
        if cursor >= len(value_cols):
            break
        group_cols = value_cols[cursor : cursor + group_size]
        first_name = str(df.columns[group_cols[0]]).strip()
        if first_name.startswith("CB "):
            label = "CB YTD" if group_index == 1 else "CB"
        elif " FY" in first_name:
            label = first_name.split(" FY", 1)[0] + " FY"
        elif " YTD" in first_name:
            label = first_name.split(" YTD", 1)[0] + " YTD"
        else:
            label = group_label_for_column(first_name)
        for col_idx in group_cols:
            labels[col_idx] = label
        cursor += group_size
    return labels


def generic_boundary_start_columns(df: pd.DataFrame, metadata_col_count: int, x_gap: bool = False) -> set[int]:
    starts = set()
    previous = None
    labels = (
        x_gap_group_labels(df, metadata_col_count)
        if x_gap
        else [group_label_for_column(column) for column in df.columns]
    )
    for idx in range(metadata_col_count, len(df.columns)):
        label = labels[idx]
        if previous is not None and label != previous:
            starts.add(idx)
        previous = label
    return starts


def auto_column_widths(df: pd.DataFrame) -> list[float]:
    widths = []
    max_widths = {
        "title": 60,
        "isbn": 18,
        "pub date": 12,
        "sea": 14,
    }
    min_widths = {
        "$": 10,
        "retail": 12,
        "units": 10,
        "disc": 8,
        "price": 9,
    }
    for col_idx, column in enumerate(df.columns):
        column_key = normalized_col_name(column)
        series = df.iloc[:, col_idx]
        values = [column]
        if not df.empty:
            sample = series.dropna()
            if column_key in {"pub date", "ship"}:
                values.extend(pd.to_datetime(sample, errors="coerce").dt.strftime("%m/%d/%Y").dropna().tolist())
            elif column_key in {"price", "$", "retail"} or "$" in column_key:
                values.extend([f"{value:,.2f}" for value in pd.to_numeric(sample, errors="coerce").dropna()])
            elif column_key in {"disc", "% tot"} or column_key.endswith("%"):
                values.extend([f"{value:.1%}" for value in pd.to_numeric(sample, errors="coerce").dropna()])
            else:
                values.extend(sample.astype(str).tolist())
        width = max(len(value) for value in values) + 2
        width = max(width, min_widths.get(column_key, 0))
        width = min(width, max_widths.get(column_key, 15))
        widths.append(width)
    return widths


def xl_col(zero_based_index: int) -> str:
    letters = ""
    value = zero_based_index + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def build_reports(period: str, refresh: bool = False, assume_yes: bool = False) -> list[Path]:
    df = load_or_fetch_data(period, refresh=refresh)
    missing = sorted(set(CUSTOMERS) - set(df["type"].dropna().unique()))
    if missing:
        print(f"Warning: no rows found for: {', '.join(missing)}")

    output_dir = prepare_output_dir(period, assume_yes=assume_yes)
    outputs = []
    for customer in CUSTOMERS:
        output_path = output_dir / output_file_name(period, customer)
        write_customer_workbook(df, period, customer, output_path)
        outputs.append(output_path)
        print(f"Saved {output_path}")

    for report in GENERIC_REPORTS:
        sql_query = report["query"](period)
        report_df = load_or_fetch_report_data(period, report["cache_name"], sql_query, refresh=refresh)
        rep_lookup = None
        if "rep_codes" in report:
            rep_lookup = fetch_rep_code_lookup(report["rep_codes"], refresh=refresh)
        output_path = output_dir / generic_output_file_name(period, report["name"])
        write_generic_workbook(report_df, period, report["name"], output_path, sql_query, rep_lookup=rep_lookup)
        outputs.append(output_path)
        print(f"Saved {output_path}")
    return outputs


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build monthly top customer title sales workbooks.")
    parser.add_argument("--period", help="Report period in YYYYMM format.")
    parser.add_argument("--refresh", action="store_true", help="Rerun SQL even when cached results exist.")
    parser.add_argument("--yes", action="store_true", help="Create missing year folder without prompting.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    period = validate_period(args.period) if args.period else prompt_for_period()
    outputs = build_reports(period, refresh=args.refresh, assume_yes=args.yes)
    print()
    print(f"Created {len(outputs)} workbook(s).")


if __name__ == "__main__":
    main()
