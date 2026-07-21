from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from amazon_rolling_reports.functions import fetch_data_from_db, get_connection

OUTPUT_DIR = Path(r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Ingram")
OH_OO_DIR = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Ingram\Ingram_OH_OO")
SALES_DIR = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Ingram\Ingram_Sales")
CACHE_DIR = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Ingram\cache")
SALES_CACHE = CACHE_DIR / "ingram_weekly_sales.parquet"
INVENTORY_DETAIL_FILE = Path(r"G:\OPS\Inventory\Daily\Finance_Only\Inventory Detail.xlsx")
BOOTSTRAP_REPORT = OUTPUT_DIR / "Daily Report - 2026 Flash Ingram (070226).xlsx"

INVENTORY_COLUMNS = [
    "On Hand Total", "On Order Total", "Customer Backorder Total",
    "Current Week Demand Total", "Previous Week Demand Total",
    "Two Weeks Ago Demand Total", "Three Weeks Ago Demand Total",
    "Four Weeks Ago Demand Total",
]


def normalize_isbn(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().replace("-", "").replace(" ", "")
    if text.endswith(".0"):
        text = text[:-2]
    digits = "".join(char for char in text if char.isdigit())
    return digits.zfill(13)[-13:] if digits else ""


def latest_xlsx(folder: Path) -> Path:
    files = [path for path in folder.glob("*.xlsx") if not path.name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"No Excel workbooks found in {folder}")
    return max(files, key=lambda path: path.stat().st_mtime)


def sales_period(path: Path) -> tuple[pd.Timestamp, pd.Timestamp]:
    match = re.search(r"(\d{8})\s*-\s*(\d{8})", path.stem)
    if not match:
        raise ValueError(f"Could not read a date range from sales filename: {path.name}")
    return tuple(pd.Timestamp(datetime.strptime(value, "%m%d%Y")) for value in match.groups())


def period_label(start: pd.Timestamp, end: pd.Timestamp, prior: bool = False) -> str:
    prefix = "Prior Sales" if prior else "Sales"
    return f"{prefix} ({start:%m/%d} - {end:%m/%d})"


def load_metadata() -> pd.DataFrame:
    query = """
    SELECT
        i.PUBLISHER_CODE AS Pub,
        i.PRODUCT_TYPE AS PT,
        i.[FORMAT] AS FT,
        i.PUBLISHING_GROUP AS PGRP,
        i.ISBN AS ISBN,
        i.SHORT_TITLE AS Title,
        i.PRICE_AMOUNT AS Price,
        i.AMORTIZATION_DATE AS PubDate
    FROM ebs.item i;
    """
    df = fetch_data_from_db(get_connection(), query)
    df["ISBN"] = df["ISBN"].map(normalize_isbn)
    return df[df["ISBN"].ne("")].drop_duplicates("ISBN")


def load_top_titles() -> pd.DataFrame:
    query = """
    SELECT i.isbn,
        CASE
            WHEN ISNULL(i.AMORTIZATION_DATE,osd.RELEASE) BETWEEN
                 DATEFROMPARTS(YEAR(GETDATE())-1,MONTH(GETDATE())+1,1) AND GETDATE() THEN 'FL'
            WHEN osd.release > GETDATE() OR i.AVAILABILITY_STATUS = 'NYP' THEN 'NYP'
            WHEN t100.isbn IS NOT NULL THEN 'top1800'
            ELSE 'other'
        END cat
    FROM ebs.Item i
    LEFT JOIN (SELECT DISTINCT isbn FROM [CBQ2].[cb].[Top100]) t100 ON i.ISBN = t100.isbn
    LEFT JOIN (
        SELECT DISTINCT ISBN, RELEASE FROM [CBQ2].[pm].[TitleSched_Current]
        WHERE IMPRESSION = 1 AND ISBN IS NOT NULL AND RELEASE IS NOT NULL
    ) osd ON i.ISBN = osd.ISBN
    GROUP BY i.isbn,
        CASE
            WHEN ISNULL(i.AMORTIZATION_DATE,osd.RELEASE) BETWEEN
                 DATEFROMPARTS(YEAR(GETDATE())-1,MONTH(GETDATE())+1,1) AND GETDATE() THEN 'FL'
            WHEN osd.release > GETDATE() OR i.AVAILABILITY_STATUS = 'NYP' THEN 'NYP'
            WHEN t100.isbn IS NOT NULL THEN 'top1800'
            ELSE 'other'
        END
    HAVING CASE
            WHEN ISNULL(i.AMORTIZATION_DATE,osd.RELEASE) BETWEEN
                 DATEFROMPARTS(YEAR(GETDATE())-1,MONTH(GETDATE())+1,1) AND GETDATE() THEN 'FL'
            WHEN osd.release > GETDATE() OR i.AVAILABILITY_STATUS = 'NYP' THEN 'NYP'
            WHEN t100.isbn IS NOT NULL THEN 'top1800'
            ELSE 'other'
        END != 'other';
    """
    df = fetch_data_from_db(get_connection(), query)
    df["ISBN"] = df["isbn"].map(normalize_isbn)
    return df[["ISBN", "cat"]].drop_duplicates("ISBN")


def load_ingram_inventory(path: Path) -> pd.DataFrame:
    usecols = ["EAN", "Title"] + INVENTORY_COLUMNS
    df = pd.read_excel(path, sheet_name=0, usecols=usecols, dtype={"EAN": str})
    df["ISBN"] = df["EAN"].map(normalize_isbn)
    for column in INVENTORY_COLUMNS:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)
    df = df[df["ISBN"].ne("")]
    titles = (
        df[["ISBN", "Title"]]
        .dropna(subset=["Title"])
        .drop_duplicates("ISBN")
        .rename(columns={"Title": "Ingram Title"})
    )
    totals = df.groupby("ISBN", as_index=False)[INVENTORY_COLUMNS].sum()
    return totals.merge(titles, on="ISBN", how="left")


def load_hachette_inventory() -> pd.DataFrame:
    columns = ["ISBN", "Ctn Qty", "Frozen", "Available To Sell", "Reprint Quantity", "Reprint Due Date"]
    df = pd.read_excel(INVENTORY_DETAIL_FILE, usecols=columns, dtype={"ISBN": str})
    df["ISBN"] = df["ISBN"].map(normalize_isbn)
    for column in ["Ctn Qty", "Frozen", "Available To Sell", "Reprint Quantity"]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)
    df = df[df["ISBN"].ne("")]
    return df.drop_duplicates("ISBN")


def bootstrap_sales_cache() -> pd.DataFrame:
    if not BOOTSTRAP_REPORT.exists():
        return pd.DataFrame(columns=["ISBN", "period_start", "period_end", "gross_qty", "source_file"])
    old = pd.read_excel(BOOTSTRAP_REPORT, sheet_name="Full List", header=2, dtype={"EAN": str})
    rows = []
    for column in old.columns:
        match = re.fullmatch(r"(?:Prior )?Sales \((\d{2}/\d{2}) - (\d{2}/\d{2})\)", str(column))
        if not match:
            continue
        start = pd.Timestamp(datetime.strptime(f"{match.group(1)}/2026", "%m/%d/%Y"))
        end = pd.Timestamp(datetime.strptime(f"{match.group(2)}/2026", "%m/%d/%Y"))
        part = pd.DataFrame({
            "ISBN": old["EAN"].map(normalize_isbn),
            "period_start": start,
            "period_end": end,
            "gross_qty": pd.to_numeric(old[column], errors="coerce").fillna(0),
            "source_file": BOOTSTRAP_REPORT.name,
        })
        rows.append(part[part["ISBN"].ne("")])
    return pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()


def update_sales_cache(sales_path: Path) -> pd.DataFrame:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache = pd.read_parquet(SALES_CACHE) if SALES_CACHE.exists() else bootstrap_sales_cache()
    start, end = sales_period(sales_path)
    sales = pd.read_excel(sales_path, usecols=["EAN", "Gross Qty"], dtype={"EAN": str})
    sales["ISBN"] = sales["EAN"].map(normalize_isbn)
    sales["gross_qty"] = pd.to_numeric(sales["Gross Qty"], errors="coerce").fillna(0)
    sales = sales[sales["ISBN"].ne("")].groupby("ISBN", as_index=False)[["gross_qty"]].sum()
    sales["period_start"], sales["period_end"], sales["source_file"] = start, end, sales_path.name
    if not cache.empty:
        cache = cache[pd.to_datetime(cache["period_end"]).dt.normalize().ne(end.normalize())]
    cache = pd.concat([cache, sales], ignore_index=True)
    cache["period_start"] = pd.to_datetime(cache["period_start"])
    cache["period_end"] = pd.to_datetime(cache["period_end"])
    cache.to_parquet(SALES_CACHE, index=False)
    return cache


def buyer(row: pd.Series) -> str:
    pub = str(row.get("Pub", "")).strip().lower()
    pt = str(row.get("PT", "")).strip().lower()
    publisher_buyers = {
        "post wave": "Tyler",
        "quadrille publishing limited": "Tyler",
        "paperblanks": "Renee",
    }
    if pub in publisher_buyers:
        return publisher_buyers[pub]
    if pub == "chronicle":
        return "Renee" if pt == "ft" else "Tyler" if pt == "bk" else "-"
    if pub == "galison":
        return "Renee"
    if pub in {"sierra club", "hardie grant", "hardie grant publishing", "laurence king", "princeton", "creative company", "tourbillon", "levine querido"}:
        return "Tyler" if pub != "sierra club" else "Renee"
    return "-"


def build_full_list(metadata, sales_cache, hachette, ingram) -> pd.DataFrame:
    periods = sales_cache[["period_start", "period_end"]].drop_duplicates().sort_values("period_end", ascending=False).head(4)
    report = metadata.copy()
    sales_columns = []
    for index, period in periods.reset_index(drop=True).iterrows():
        start, end = period["period_start"], period["period_end"]
        column = period_label(start, end, prior=index > 0)
        values = sales_cache[pd.to_datetime(sales_cache["period_end"]).dt.normalize().eq(end.normalize())][["ISBN", "gross_qty"]]
        report = report.merge(values.rename(columns={"gross_qty": column}), on="ISBN", how="left")
        sales_columns.append(column)
    for column in sales_columns:
        report[column] = pd.to_numeric(report[column], errors="coerce").fillna(0)
    while len(sales_columns) < 4:
        column = f"Prior Sales (Unavailable {len(sales_columns)})"
        report[column] = 0
        sales_columns.append(column)
    report["QTY Variance"] = report[sales_columns[0]] - report[sales_columns[1]]
    report["Avg 4wk Sales"] = report[sales_columns].sum(axis=1) / 4
    report = report.merge(ingram, on="ISBN", how="outer").merge(hachette, on="ISBN", how="left")
    report["Title"] = report["Title"].fillna(report.get("Ingram Title")).fillna("")
    numeric = ["Ctn Qty", "Frozen", "Available To Sell", "Reprint Quantity"] + INVENTORY_COLUMNS
    for column in numeric:
        report[column] = pd.to_numeric(report.get(column, 0), errors="coerce").fillna(0)
    for column in sales_columns:
        report[column] = pd.to_numeric(report[column], errors="coerce").fillna(0)
    report["QTY Variance"] = report[sales_columns[0]] - report[sales_columns[1]]
    report["Avg 4wk Sales"] = report[sales_columns].sum(axis=1) / 4
    activity_columns = sales_columns + numeric
    report = report[report[activity_columns].ne(0).any(axis=1)].copy()
    report["Total Ingram OH & OO"] = report["On Hand Total"] + report["On Order Total"]
    report["Total Demand last 4 weeks"] = report[
        [
            "Previous Week Demand Total",
            "Two Weeks Ago Demand Total",
            "Three Weeks Ago Demand Total",
            "Four Weeks Ago Demand Total",
        ]
    ].sum(axis=1)
    report["Suggested Buy"] = (
        report["Total Demand last 4 weeks"] * 3
        + report["Customer Backorder Total"]
        - report["Total Ingram OH & OO"]
    ).clip(lower=0)
    report["Buyer"] = report.apply(buyer, axis=1)
    rename = {
        "Ctn Qty": "Carton Qty", "Frozen": "Total Freezes",
        "Available To Sell": "Hachette Available", "On Hand Total": "Ingram On Hand",
        "On Order Total": "Ingram On Order", "Customer Backorder Total": "Total Ingram's Customer BO",
    }
    report = report.rename(columns=rename)
    columns = ["Pub", "PT", "FT", "PGRP", "ISBN", "Title", "Price", "PubDate", "QTY Variance"] + sales_columns + [
        "Avg 4wk Sales", "Carton Qty", "Total Freezes", "Hachette Available", "Reprint Quantity",
        "Reprint Due Date", "Ingram On Hand", "Ingram On Order", "Total Ingram OH & OO",
        "Total Ingram's Customer BO", "Suggested Buy", "Buyer",
    ]
    return (
        report[columns]
        .drop_duplicates(subset="ISBN", keep="first")
        .sort_values(sales_columns[0], ascending=False)
    )


def build_inventory_tab(ingram: pd.DataFrame, metadata: pd.DataFrame) -> pd.DataFrame:
    ingram = ingram[
        ingram[INVENTORY_COLUMNS]
        .apply(pd.to_numeric, errors="coerce")
        .fillna(0)
        .ne(0)
        .any(axis=1)
    ].copy()
    result = ingram.merge(
        metadata[["ISBN", "Title", "Pub"]]
        .drop_duplicates("ISBN")
        .rename(columns={"Title": "EBS Title"}),
        on="ISBN",
        how="left",
    )
    result["Title"] = result["EBS Title"].fillna(result["Ingram Title"]).fillna("")
    result["Pub"] = result["Pub"].fillna("")
    result["Total Demand last 4 weeks"] = result[
        ["Previous Week Demand Total", "Two Weeks Ago Demand Total", "Three Weeks Ago Demand Total", "Four Weeks Ago Demand Total"]
    ].sum(axis=1)
    return (
        result[["ISBN", "Title", "Pub"] + INVENTORY_COLUMNS + ["Total Demand last 4 weeks"]]
        .drop_duplicates(subset="ISBN", keep="first")
        .sort_values("On Hand Total", ascending=False)
        .reset_index(drop=True)
    )


def write_report(full_list: pd.DataFrame, inventory: pd.DataFrame, output: Path, start: pd.Timestamp, end: pd.Timestamp) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="mm/dd/yyyy") as writer:
        full_header_row = 3
        inventory_header_row = 4
        full_list.to_excel(writer, sheet_name="Full List", startrow=full_header_row, index=False)
        inventory.to_excel(writer, sheet_name="Ingram Inventory", startrow=inventory_header_row, index=False)
        workbook = writer.book
        header = workbook.add_format({
            "bold": True, "bg_color": "#B8CCE4", "border": 1,
            "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        title = workbook.add_format({
            "bold": True, "font_size": 14, "bg_color": "#FCD5B4",
            "align": "center_across", "valign": "vcenter",
        })
        accounting = workbook.add_format({
            "num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
        })
        accounting_total = workbook.add_format({
            "bold": True, "bg_color": "#E4DFEC", "border": 1,
            "num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
        })
        inventory_total = workbook.add_format({
            "bold": True, "bg_color": "#DCE6F1",
            "num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
        })
        inventory_total_label = workbook.add_format({
            "bold": True, "bg_color": "#DCE6F1",
        })
        total_label = workbook.add_format({
            "bold": True, "bg_color": "#CCC0DA", "border": 1,
        })
        price_format = workbook.add_format({"num_format": "$#,##0.00"})
        date_format = workbook.add_format({"num_format": "mm/dd/yyyy"})
        isbn_format = workbook.add_format({"num_format": "0"})

        for name, df, header_row in [
            ("Full List", full_list, full_header_row),
            ("Ingram Inventory", inventory, inventory_header_row),
        ]:
            ws = writer.sheets[name]
            for col, value in enumerate(df.columns):
                ws.write(header_row, col, value, header)
            ws.set_row(header_row, 32)
            ws.freeze_panes(header_row + 1, 0)
            ws.autofilter(header_row, 0, header_row + len(df), len(df.columns) - 1)
            if name == "Ingram Inventory":
                for col in range(len(df.columns)):
                    cell_format = accounting if col >= 3 else (isbn_format if col == 0 else None)
                    ws.set_column(col, col, 16.71, cell_format)
                ws.write(1, 2, "Totals", inventory_total_label)
                ws.write(2, 2, "Subtotals", inventory_total_label)
                first_data_row = inventory_header_row + 1
                last_data_row = inventory_header_row + len(df)
                from xlsxwriter.utility import xl_rowcol_to_cell
                for col in range(3, len(df.columns)):
                    first_cell = xl_rowcol_to_cell(first_data_row, col)
                    last_cell = xl_rowcol_to_cell(last_data_row, col)
                    total_value = float(pd.to_numeric(df.iloc[:, col], errors="coerce").fillna(0).sum())
                    ws.write_formula(
                        1, col, f"=SUM({first_cell}:{last_cell})",
                        inventory_total, total_value,
                    )
                    ws.write_formula(
                        2, col, f"=SUBTOTAL(9,{first_cell}:{last_cell})",
                        inventory_total, total_value,
                    )

        full_ws = writer.sheets["Full List"]
        full_ws.freeze_panes(4, 8)
        full_ws.write(
            0, 0,
            f"Chronicle Ingram Full List {start:%m/%d/%y} to {end:%m/%d/%y}",
            title,
        )
        for column in range(1, 5):
            full_ws.write_blank(0, column, None, title)
        full_ws.set_column(0, len(full_list.columns) - 1, 12)
        full_ws.set_column(full_list.columns.get_loc("ISBN"), full_list.columns.get_loc("ISBN"), 14, isbn_format)
        full_ws.set_column(full_list.columns.get_loc("Title"), full_list.columns.get_loc("Title"), 37.57)
        full_ws.set_column(full_list.columns.get_loc("Price"), full_list.columns.get_loc("Price"), 10, price_format)
        full_ws.set_column(full_list.columns.get_loc("PubDate"), full_list.columns.get_loc("PubDate"), 11, date_format)
        non_numeric = {
            "Pub", "PT", "FT", "PGRP", "ISBN", "Title", "Price", "PubDate",
            "Reprint Due Date", "Buyer",
        }
        numeric_columns = [column for column in full_list.columns if column not in non_numeric]
        for column in numeric_columns:
            col = full_list.columns.get_loc(column)
            width = 14.5 if 9 <= col <= 12 else (9.71 if column == "Avg 4wk Sales" else 13)
            full_ws.set_column(col, col, width, accounting)

        for col in range(9, 13):
            full_ws.set_column(col, col, 14.5, accounting)
        avg_col = full_list.columns.get_loc("Avg 4wk Sales")
        full_ws.set_column(avg_col, avg_col, 9.71, accounting)

        first_numeric_col = full_list.columns.get_loc("QTY Variance")
        label_col = first_numeric_col - 1
        full_ws.write(0, label_col, "Total", total_label)
        full_ws.write(1, label_col, "Subtotal", total_label)
        first_data_row = full_header_row + 1
        last_data_row = full_header_row + len(full_list)
        from xlsxwriter.utility import xl_rowcol_to_cell
        for column in numeric_columns:
            col = full_list.columns.get_loc(column)
            first_cell = xl_rowcol_to_cell(first_data_row, col)
            last_cell = xl_rowcol_to_cell(last_data_row, col)
            total_value = float(pd.to_numeric(full_list[column], errors="coerce").fillna(0).sum())
            full_ws.write_formula(
                0, col, f"=SUM({first_cell}:{last_cell})", accounting_total, total_value
            )
            full_ws.write_formula(
                1, col, f"=SUBTOTAL(9,{first_cell}:{last_cell})", accounting_total, total_value
            )


def confirm_sources(sales: Path, inventory: Path) -> bool:
    print("\nIngram weekly report sources")
    print(f"  Sales:              {sales}")
    print(f"  Ingram OH/OO:       {inventory}")
    print(f"  Hachette inventory: {INVENTORY_DETAIL_FILE}")
    while True:
        answer = input("Use these files? [Y/n]: ").strip().lower()
        if answer in {"", "y", "yes"}: return True
        if answer in {"n", "no"}: return False
        print("Please enter y or n.")


def main() -> int:
    parser = argparse.ArgumentParser(description="Build the weekly Ingram report.")
    parser.add_argument("--yes", action="store_true", help="Use the latest source files without prompting.")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    sales_path, ingram_path = latest_xlsx(SALES_DIR), latest_xlsx(OH_OO_DIR)
    if not args.yes and not confirm_sources(sales_path, ingram_path):
        print("Ingram weekly report cancelled.")
        return 10
    start, end = sales_period(sales_path)
    output = args.output or OUTPUT_DIR / f"Daily Report - {end.year} Flash Ingram ({pd.Timestamp.today():%m%d%y}).xlsx"
    cache = update_sales_cache(sales_path)
    ingram = load_ingram_inventory(ingram_path)
    metadata = load_metadata()
    full = build_full_list(metadata, cache, load_hachette_inventory(), ingram)
    write_report(full, build_inventory_tab(ingram, metadata), output, start, end)
    print(f"Saved Ingram weekly report: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
