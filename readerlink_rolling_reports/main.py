from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

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

POS_HISTORY_CACHE = CACHE_DIR / "readerlink_pos_history.parquet"
LATEST_INVENTORY_CACHE = CACHE_DIR / "readerlink_latest_inventory.parquet"
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
    "SAMS",
    "SAMSCLUB.COM",
    "TARGET",
    "WALMART",
    "WHOLE FOODS",
    "WM.COM",
]
CHAIN_SHEET_ORDER = ["Summary - All Accounts"] + TRACKED_CHAINS + ["All Other Accounts"]
CHAIN_ALIASES = {
    "TARGET STORES": "TARGET",
    "TARGET": "TARGET",
    "SAMS CLUB": "SAMS",
    "SAM'S CLUB": "SAMS",
    "SAMS": "SAMS",
    "BJS": "BJS",
    "COSTCO": "COSTCO",
    "FRED MEYER": "FRED MEYER",
    "HUDSON": "HUDSON",
    "KROGER": "KROGER",
    "MEIJER": "MEIJER",
    "PARADIES": "PARADIES",
    "AAFES": "AAFES",
    "WALMART": "WALMART",
    "WHOLE FOODS": "WHOLE FOODS",
    "WM.COM": "WM.COM",
    "SAMSCLUB.COM": "SAMSCLUB.COM",
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
    date_value = pd.Timestamp(date_value)
    jan1 = pd.Timestamp(year=date_value.year, month=1, day=1)
    days_to_saturday = (5 - jan1.weekday()) % 7
    first_saturday = jan1 + pd.Timedelta(days=days_to_saturday)
    if date_value < first_saturday:
        return readerlink_week(pd.Timestamp(year=date_value.year - 1, month=12, day=31))
    return int(((date_value - first_saturday).days // 7) + 1)


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
        usecols=["ISBN", "Available To Sell", "Frozen"],
        dtype={"ISBN": str},
    )
    df["ISBN"] = df["ISBN"].map(normalize_isbn)
    df = df[df["ISBN"] != ""]
    for column in ["Available To Sell", "Frozen"]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)
    df = df.groupby("ISBN", as_index=False)[["Available To Sell", "Frozen"]].sum()
    return df.rename(columns={"Available To Sell": "HBG_Avail", "Frozen": "Freezes"})


def load_readerlink_inventory() -> pd.DataFrame:
    df = pd.read_parquet(LATEST_INVENTORY_CACHE)
    df = df.rename(
        columns={
            "isbn": "ISBN",
            "rds_dc_inventory_oh": "OH",
            "rds_open_po_quantity": "OO",
        }
    )
    df["ISBN"] = df["ISBN"].map(normalize_isbn)
    return df.groupby("ISBN", as_index=False)[["OH", "OO"]].sum()


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
    readerlink_inventory: pd.DataFrame,
    all_week_cols: list[str],
    report_week_cols: list[str],
    older_year_cols: list[str],
    latest_week: pd.Timestamp,
) -> pd.DataFrame:
    history = build_sheet_history(pos_history, sheet_name, all_week_cols)
    inventory_isbns = readerlink_inventory[
        pd.to_numeric(readerlink_inventory["OH"], errors="coerce").fillna(0).ne(0)
        | pd.to_numeric(readerlink_inventory["OO"], errors="coerce").fillna(0).ne(0)
    ][["ISBN"]].drop_duplicates()
    history = pd.concat([history, inventory_isbns], ignore_index=True, sort=False)
    history["ISBN"] = history["ISBN"].map(normalize_isbn)
    history = history[history["ISBN"] != ""].drop_duplicates(subset="ISBN", keep="first")
    for column in all_week_cols:
        if column not in history.columns:
            history[column] = 0
    history[all_week_cols] = history[all_week_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

    report = history.merge(metadata, on="ISBN", how="left")
    report = report[report["Pub"].notna() & report["Pub"].astype(str).str.strip().ne("")]
    report = report.merge(hbg_inventory, on="ISBN", how="left")
    report = report.merge(readerlink_inventory, on="ISBN", how="left")

    for column in ["HBG_Avail", "Freezes", "OH", "OO", "Price"]:
        report[column] = pd.to_numeric(report.get(column, 0), errors="coerce").fillna(0)

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
    report["OH_Avg"] = (report["OH"] / report["6Wk Avg"].where(report["6Wk Avg"].ne(0))).fillna(0).round(2)
    report["OH+OO_Avg"] = (
        (report["OH"] + report["OO"]) / report["6Wk Avg"].where(report["6Wk Avg"].ne(0))
    ).fillna(0).round(2)
    if latest_label in report.columns and prior_label in report.columns:
        report["PW %"] = (
            (report[latest_label] - report[prior_label])
            / report[prior_label].where(report[prior_label].ne(0))
        ).fillna(0)
    else:
        report["PW %"] = 0

    active_mask = (
        report["OH"].ne(0)
        | report["OO"].ne(0)
        | report[active_sales_cols].sum(axis=1).ne(0)
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
        "Freezes",
        "OH",
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
    return OUTPUT_DIR / f"Week {readerlink_week(latest_week):02d} - {latest_week.year} Rolling Readerlink ({latest_week.strftime('%m%d%y')}).xlsx"


def write_workbook(sheets: dict[str, pd.DataFrame], output_path: Path, latest_week: pd.Timestamp) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    header_row = 4
    data_start_row = header_row + 1
    pre_week_cols = 26

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

        for sheet_name, df in sheets.items():
            safe_sheet = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_sheet, startrow=header_row, index=False)
            ws = writer.sheets[safe_sheet]
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
                        oh_total = _xl_cell(0, col_positions["OH"])
                        avg_total = _xl_cell(0, col_positions["6Wk Avg"])
                        oh_subtotal = _xl_cell(1, col_positions["OH"])
                        avg_subtotal = _xl_cell(1, col_positions["6Wk Avg"])
                        total_formula = f'=IFERROR({oh_total}/{avg_total},0)'
                        subtotal_formula = f'=IFERROR({oh_subtotal}/{avg_subtotal},0)'
                        total_6wk = pd.to_numeric(df["6Wk Avg"], errors="coerce").fillna(0).sum()
                        value = 0 if total_6wk == 0 else float(pd.to_numeric(df["OH"], errors="coerce").fillna(0).sum() / total_6wk)
                    elif col_name == "OH+OO_Avg":
                        oh_total = _xl_cell(0, col_positions["OH"])
                        oo_total = _xl_cell(0, col_positions["OO"])
                        avg_total = _xl_cell(0, col_positions["6Wk Avg"])
                        oh_subtotal = _xl_cell(1, col_positions["OH"])
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
                                    pd.to_numeric(df["OH"], errors="coerce").fillna(0).sum()
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
            ws.set_column(25, 25, 9, pct_fmt)

    print(f"Saved Readerlink rolling report: {output_path}")


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
    metadata = load_metadata()
    hbg_inventory = load_hbg_inventory()
    readerlink_inventory = load_readerlink_inventory()

    sheets: dict[str, pd.DataFrame] = {}
    for sheet_name in CHAIN_SHEET_ORDER:
        print(f"Building sheet: {sheet_name}")
        sheets[sheet_name] = build_report_sheet(
            sheet_name,
            pos_history,
            metadata,
            hbg_inventory,
            readerlink_inventory,
            week_cols,
            display_week_cols,
            display_year_cols,
            latest_week,
        )

    output_path = args.output or output_file_for(latest_week)
    write_workbook(sheets, output_path, latest_week)


if __name__ == "__main__":
    main()
