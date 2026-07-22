from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from xlsxwriter.utility import xl_col_to_name

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from amazon_rolling_reports.functions import fetch_data_from_db, get_connection
from paths import process_paths
from shared.bookscan_calendar import bookscan_parts, bookscan_week

F = Path(r"F:\ANALYSIS\Finance\DataWarehouse")
AMZ = F / "Atelier AmazonRolling" / "cache"
AMZ_CO, AMZ_US, AMZ_PO = AMZ / "rr_customer_orders.pkl", AMZ / "rr_units_shipped.pkl", AMZ / "latest_amazon_po.pkl"
BN = F / "Atelier BarnesNoble" / "cache"
BN_SALES, BN_INV = BN / "bn_customer_sales.parquet", BN / "bn_inventory_snapshots.parquet"
ED = F / "Atelier Edelweiss" / "cache"
ED_SALES, ED_META = ED / "edelweiss_sales.parquet", ED / "edelweiss_metadata.parquet"
RL = F / "Atelier Readerlink" / "cache" / "readerlink_pos_history.parquet"
RL_STORE_CACHE = F / "Atelier Readerlink" / "cache" / "readerlink_store_inventory_history.parquet"
RL_STORE = F / "Weekly reports" / "2026" / "Readerlink" / "OH_Store_TitlePerformanceReport"
TG = process_paths.TARGET_NOC_CACHE_DIR
TG_SALES, TG_META, TG_INV = TG / "target_noc_weekly_sales.parquet", TG / "target_noc_metadata.parquet", TG / "target_noc_inventory.parquet"
DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")
ACCOUNTING = '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'


def normalize_isbn(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    digits = re.sub(r"\D", "", text)
    return digits.zfill(13)[-13:] if digits else ""


def num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def grouped(df: pd.DataFrame, mask: pd.Series, value: str) -> pd.Series:
    if not mask.any():
        return pd.Series(dtype="float64")
    return df.loc[mask].groupby("ISBN")[value].sum()


def weekly_metrics(
    df: pd.DataFrame,
    *,
    date_col: str,
    value_col: str,
    include_last_year: bool = True,
    as_of: pd.Timestamp | None = None,
):
    work = df[["ISBN", date_col, value_col]].copy()
    work["ISBN"] = work["ISBN"].map(normalize_isbn)
    work[date_col] = pd.to_datetime(work[date_col], errors="coerce").dt.normalize()
    work[value_col] = num(work[value_col])
    work = work[work["ISBN"].ne("") & work[date_col].notna()]
    if work.empty:
        raise ValueError(f"No usable rows for {date_col}/{value_col}")
    latest = work[date_col].max()
    current = bookscan_week(as_of if as_of is not None else latest)
    parts = bookscan_parts(work[date_col])
    frames = [
        grouped(work, work[date_col].eq(latest), value_col).rename("Weekly"),
        grouped(work, parts.BookScanYear.eq(current.year), value_col).rename("YTD"),
        grouped(work, parts.BookScanYear.eq(current.year - 1) & parts.BookScanWeek.le(current.week), value_col).rename("LYTD"),
    ]
    if include_last_year:
        frames.append(grouped(work, parts.BookScanYear.eq(current.year - 1), value_col).rename("Last Year"))
    return pd.concat(frames, axis=1).fillna(0), latest


def amazon_metrics():
    customer, shipped, po = pd.read_pickle(AMZ_CO), pd.read_pickle(AMZ_US), pd.read_pickle(AMZ_PO)
    dates = sorted((pd.Timestamp(datetime.strptime(c, "%m-%d-%Y")), c) for c in shipped.columns if DATE_RE.fullmatch(str(c)))
    if not dates:
        raise ValueError("Amazon cache has no weekly columns")
    latest, latest_col = dates[-1]
    current = bookscan_week(latest)

    def history(frame, year, through=None):
        cols = []
        for column in frame.columns:
            if not DATE_RE.fullmatch(str(column)):
                continue
            part = bookscan_week(pd.Timestamp(datetime.strptime(column, "%m-%d-%Y")))
            if part.year == year and (through is None or part.week <= through):
                cols.append(column)
        return frame[cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if cols else pd.Series(0, index=frame.index)

    out = pd.DataFrame({"ISBN": shipped.ISBN.map(normalize_isbn)})
    co = customer.assign(ISBN=customer.ISBN.map(normalize_isbn)).drop_duplicates("ISBN").set_index("ISBN")
    out["Customer Order"] = out.ISBN.map(num(co[latest_col]) if latest_col in co else pd.Series(dtype=float)).fillna(0)
    out["Units Shipped"] = num(shipped[latest_col])
    out["Units YTD"] = history(shipped, current.year, current.week)
    out["Units LYTD"] = history(shipped, current.year - 1, current.week)
    out["Units Last Year"] = history(shipped, current.year - 1)
    out["On Hand"] = num(shipped.get("OH", pd.Series(0, index=shipped.index)))
    po["ISBN"] = po.ISBN.map(normalize_isbn)
    po["PO_Qty"] = num(po.PO_Qty)
    out["On Order"] = out.ISBN.map(po.groupby("ISBN").PO_Qty.sum()).fillna(0)
    return out.groupby("ISBN").sum(numeric_only=True), latest


def bn_metrics():
    sales = pd.read_parquet(BN_SALES, columns=["Week", "ISBN", "qty"])
    result, latest = weekly_metrics(sales, date_col="Week", value_col="qty")
    inv = pd.read_parquet(BN_INV)
    inv["Week"] = pd.to_datetime(inv.Week)
    inv = inv[inv.Week.eq(inv.Week.max())].copy()
    inv["ISBN"] = inv.ISBN.map(normalize_isbn)
    inv["On Hand"] = num(inv.OH_Stores) + num(inv.OH_DC)
    inv["On Order"] = num(inv.OO_Stores) + num(inv.OO_DC)
    return result.join(inv.groupby("ISBN")[["On Hand", "On Order"]].sum(), how="outer").fillna(0), latest


def edelweiss_metrics():
    sales = pd.read_parquet(ED_SALES, columns=["Week", "ISBN", "qty"])
    result, latest = weekly_metrics(sales, date_col="Week", value_col="qty")
    meta = pd.read_parquet(ED_META, columns=["ISBN", "On Hand", "On Order"])
    meta["ISBN"] = meta.ISBN.map(normalize_isbn)
    return result.join(meta.groupby("ISBN")[["On Hand", "On Order"]].sum(), how="outer").fillna(0), latest


def latest_xlsx(folder: Path) -> Path:
    files = [p for p in folder.glob("*.xlsx") if not p.name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"No workbook in {folder}")
    return max(files, key=lambda p: p.stat().st_mtime)


def target_metrics():
    target = pd.read_parquet(TG_SALES, columns=["ISBN", "Week", "Units"])
    reader = pd.read_parquet(RL, columns=["isbn", "week_end", "master_chain", "cy_pos_units"]).rename(
        columns={"isbn": "ISBN", "week_end": "Week", "cy_pos_units": "Units"})
    chains = reader.master_chain.astype("string").str.strip().str.upper()
    reader = reader[chains.isin({"TARGET", "TARGET STORES"})]
    td = pd.to_datetime(target["Week"]).max()
    rd = pd.to_datetime(reader["Week"]).max()
    target_as_of = max(td, rd)
    tm, _ = weekly_metrics(target, date_col="Week", value_col="Units", as_of=target_as_of)
    rm, _ = weekly_metrics(reader, date_col="Week", value_col="Units", as_of=target_as_of)
    result = tm.add(rm, fill_value=0)

    meta = pd.read_parquet(TG_META, columns=["DPCI", "ISBN"])
    inv = pd.read_parquet(TG_INV)
    inv["Week"] = pd.to_datetime(inv.Week)
    inv = inv[inv.Week.eq(inv.Week.max())].merge(meta, on="DPCI", how="left")
    inv["ISBN"] = inv.ISBN.map(normalize_isbn)
    target_oh = inv.groupby("ISBN")["On Hand"].sum()

    if RL_STORE_CACHE.exists():
        store = pd.read_parquet(RL_STORE_CACHE)
        store["week_end"] = pd.to_datetime(store["week_end"])
        store = store[store["week_end"].eq(store["week_end"].max())].rename(
            columns={"master_chain": "MASTER CHAIN", "isbn": "EAN", "store_oh_units": "TOTAL OH UNITS"}
        )
        store_source = RL_STORE_CACHE
    else:
        store_source = latest_xlsx(RL_STORE)
        store = pd.read_excel(store_source, sheet_name="Export", usecols=["MASTER CHAIN", "EAN", "TOTAL OH UNITS"], dtype={"EAN": str})
    chains = store["MASTER CHAIN"].astype("string").str.strip().str.upper()
    store = store[chains.isin({"TARGET", "TARGET STORES"})].copy()
    store["ISBN"], store["On Hand"] = store.EAN.map(normalize_isbn), num(store["TOTAL OH UNITS"])
    result["On Hand"] = target_oh.add(store.groupby("ISBN")["On Hand"].sum(), fill_value=0)
    # Filled from hachette.HachetteOrders after the shared SQL query is joined.
    result["On Order (NOC)"] = 0
    return result.fillna(0), td, rd, store_source


ITEM_SQL = """
WITH sales_summary AS (
 SELECT sd.ITEM_ID,
  SUM(CASE WHEN sd.TRX_DATE >= DATEFROMPARTS(YEAR(GETDATE()),1,1)
            AND sd.TRX_DATE < DATEADD(day,1,CAST(GETDATE() AS date)) THEN sd.QUANTITY_INVOICED ELSE 0 END) YTD,
  SUM(CASE WHEN sd.TRX_DATE >= DATEFROMPARTS(YEAR(GETDATE())-1,1,1)
            AND sd.TRX_DATE < DATEADD(day,1,DATEADD(year,-1,CAST(GETDATE() AS date))) THEN sd.QUANTITY_INVOICED ELSE 0 END) LYTD,
  SUM(CASE WHEN sd.TRX_DATE >= DATEFROMPARTS(YEAR(GETDATE())-1,1,1)
            AND sd.TRX_DATE < DATEFROMPARTS(YEAR(GETDATE()),1,1) THEN sd.QUANTITY_INVOICED ELSE 0 END) FYLY
 FROM ebs.sales sd
 WHERE sd.TRX_DATE >= DATEFROMPARTS(YEAR(GETDATE())-1,1,1)
  AND cbq2.dbo.fnSaleTypeCode(sd.AR_TRX_TYPE_ID)='N' AND sd.INVOICE_LINE_TYPE='SALE'
 GROUP BY sd.ITEM_ID
), ltd_summary AS (
 SELECT ys.ITEM_ID, SUM(ys.SalesQty) LTD FROM summary.ItemBillToYearlySales ys GROUP BY ys.ITEM_ID
), target_noc_open_orders AS (
 SELECT ho.ISBN, SUM(ho.Quantity) Qty
 FROM hachette.HachetteOrders ho
 WHERE ho.SSRRowID = '140'
 GROUP BY ho.ISBN
)
SELECT i.ITEM_TITLE ISBN, i.PUBLISHER_CODE Pub, i.PRODUCT_TYPE PT, i.FORMAT Cat,
 i.PUBLISHING_GROUP PGRP, i.SHORT_TITLE Title, i.PRICE_AMOUNT Price,
 CAST(i.AMORTIZATION_DATE AS date) PubDate,
 COALESCE(s.YTD,0) YTD, COALESCE(s.LYTD,0) LYTD, COALESCE(s.FYLY,0) FYLY, COALESCE(l.LTD,0) LTD,
 COALESCE(tno.Qty,0) TargetNOCOpenOrder
FROM ebs.Item i
LEFT JOIN sales_summary s ON s.ITEM_ID=i.ITEM_ID
LEFT JOIN ltd_summary l ON l.ITEM_ID=i.ITEM_ID
LEFT JOIN target_noc_open_orders tno ON tno.ISBN=i.ITEM_TITLE
WHERE i.PRODUCT_TYPE IN ('BK','FT','CP','RP') AND i.ITEM_TITLE IS NOT NULL;
"""


def item_and_sell_in():
    print("Loading EBS item metadata and dynamic sell-in totals...", flush=True)
    data = fetch_data_from_db(get_connection(), ITEM_SQL)
    data["ISBN"] = data.ISBN.map(normalize_isbn)
    data = data[data.ISBN.ne("")].drop_duplicates("ISBN")
    for column in ["Price", "YTD", "LYTD", "FYLY", "LTD", "TargetNOCOpenOrder"]:
        data[column] = num(data[column])
    data["PubDate"] = pd.to_datetime(data.PubDate, errors="coerce")
    return data.set_index("ISBN")


def max_date(path: Path, column: str):
    if not path.exists():
        return None
    value = pd.to_datetime(pd.read_parquet(path, columns=[column])[column], errors="coerce").max()
    return None if pd.isna(value) else pd.Timestamp(value)


def amazon_date():
    if not AMZ_US.exists():
        return None
    columns = pd.read_pickle(AMZ_US).columns
    values = [pd.Timestamp(datetime.strptime(c, "%m-%d-%Y")) for c in columns if DATE_RE.fullmatch(str(c))]
    return max(values) if values else None


def source_rows():
    if RL_STORE_CACHE.exists():
        store = RL_STORE_CACHE
        store_through = max_date(RL_STORE_CACHE, "week_end")
    else:
        store = latest_xlsx(RL_STORE) if RL_STORE.exists() else RL_STORE
        store_through = None
    ad = amazon_date()
    return [
        ("Amazon customer orders", AMZ_CO, ad, ""), ("Amazon units shipped", AMZ_US, ad, ""),
        ("Amazon on order", AMZ_PO, None, ""), ("Barnes & Noble sales", BN_SALES, max_date(BN_SALES, "Week"), ""),
        ("Barnes & Noble inventory", BN_INV, max_date(BN_INV, "Week"), ""),
        ("Edelweiss sales", ED_SALES, max_date(ED_SALES, "Week"), ""),
        ("Edelweiss inventory", ED_META, max_date(ED_SALES, "Week"), ""),
        ("Readerlink Target-tab history", RL, max_date(RL, "week_end"), ""),
        ("Readerlink Target store OH", store, store_through, ""),
        ("Target NOC sales", TG_SALES, max_date(TG_SALES, "Week"), ""),
        ("Target NOC inventory", TG_INV, max_date(TG_INV, "Week"), "open orders use live SQL"),
    ]


def show_sources(rows):
    print("\nX-Gap data sources")
    print("=" * 100)
    print(f"{'Source':<34} {'Through':<14} {'File modified':<20} Status")
    for label, path, through, note in rows:
        thru = through.strftime("%m/%d/%Y") if through is not None else "-"
        modified = datetime.fromtimestamp(path.stat().st_mtime).strftime("%m/%d/%Y %I:%M %p") if path.exists() else "-"
        state = "Ready" if path.exists() else "MISSING"
        if note:
            state += "; " + note
        print(f"{label:<34} {thru:<14} {modified:<20} {state}")
    print(f"{'EBS item / sell-in SQL':<34} {pd.Timestamp.today():%m/%d/%Y}     {'live query':<20} Ready")
    print(f"{'Target NOC open orders':<34} {pd.Timestamp.today():%m/%d/%Y}     {'HachetteOrders':<20} Ready (SSRRowID 140)")


def prefixed(frame, prefix):
    return frame.rename(columns={c: f"{prefix} {c}" for c in frame.columns})


def build_dataframe():
    print("Loading Amazon cache...", flush=True)
    amazon, amazon_d = amazon_metrics()
    print("Loading Barnes & Noble cache...", flush=True)
    bn, bn_d = bn_metrics()
    print("Loading Edelweiss cache...", flush=True)
    ed, ed_d = edelweiss_metrics()
    print("Loading Target and Readerlink caches...", flush=True)
    target, target_noc_d, readerlink_d, store_file = target_metrics()

    all_data = prefixed(amazon, "Amazon").join(prefixed(bn, "BN"), how="outer")
    for frame, prefix in [(ed, "Edelweiss"), (target, "Target")]:
        all_data = all_data.join(prefixed(frame, prefix), how="outer")
    all_data = all_data.fillna(0)
    totals = {
        "Weekly": ["Amazon Units Shipped", "BN Weekly", "Edelweiss Weekly", "Target Weekly"],
        "YTD": ["Amazon Units YTD", "BN YTD", "Edelweiss YTD", "Target YTD"],
        "LYTD": ["Amazon Units LYTD", "BN LYTD", "Edelweiss LYTD", "Target LYTD"],
        "Last Year": ["Amazon Units Last Year", "BN Last Year", "Edelweiss Last Year", "Target Last Year"],
    }
    for metric, columns in totals.items():
        for column in columns:
            if column not in all_data:
                all_data[column] = 0
        all_data[f"Total {metric}"] = all_data[columns].sum(axis=1)
    all_data["Total YOY"] = all_data["Total YTD"] - all_data["Total LYTD"]
    all_data = all_data[all_data["Total YTD"].ne(0) | all_data["Total LYTD"].ne(0)]
    all_data = item_and_sell_in().join(all_data, how="inner")
    all_data["Target On Order (NOC)"] = all_data["TargetNOCOpenOrder"]
    for account, weekly, ytd in [("Amazon", "Amazon Units Shipped", "Amazon Units YTD"), ("BN", "BN Weekly", "BN YTD"), ("Edelweiss", "Edelweiss Weekly", "Edelweiss YTD"), ("Target", "Target Weekly", "Target YTD")]:
        all_data[f"{account} % Total Weekly"] = all_data[weekly].div(all_data["Total Weekly"].where(all_data["Total Weekly"].ne(0))).fillna(0)
        all_data[f"{account} % Total YTD"] = all_data[ytd].div(all_data["Total YTD"].where(all_data["Total YTD"].ne(0))).fillna(0)

    sources = [
        "Pub", "PT", "Cat", "PGRP", None, "Title", "Price", "PubDate",
        "Total Weekly", "Total YTD", "Total LYTD", "Total YOY", "Total Last Year",
        "Amazon Customer Order", "Amazon Units Shipped", "Amazon Units YTD", "Amazon Units LYTD", "Amazon Units Last Year",
        "Amazon % Total Weekly", "Amazon % Total YTD", "Amazon On Hand", "Amazon On Order",
        "BN Weekly", "BN YTD", "BN LYTD", "BN Last Year", "BN % Total Weekly", "BN % Total YTD", "BN On Hand", "BN On Order",
        "Edelweiss Weekly", "Edelweiss YTD", "Edelweiss LYTD", "Edelweiss Last Year", "Edelweiss % Total Weekly", "Edelweiss % Total YTD", "Edelweiss On Hand", "Edelweiss On Order",
        "Target Weekly", "Target YTD", "Target LYTD", "Target Last Year", "Target % Total Weekly", "Target % Total YTD", "Target On Hand", "Target On Order (NOC)",
        "YTD", "LYTD", "FYLY", "LTD",
    ]
    output = pd.DataFrame(index=all_data.index)
    for index, source in enumerate(sources):
        output[index] = all_data.index if source is None else all_data[source]
    output = output.sort_values([8, 0, 3, 5, 4], ascending=[False, True, True, True, True], kind="stable").reset_index(drop=True)
    freshness = {
        "Amazon": amazon_d, "B&N": bn_d, "Edelweiss": ed_d,
        "Target": max(target_noc_d, readerlink_d),
        "Sell-In": pd.Timestamp.today().normalize(),
    }
    return output, freshness


HEADERS = [
    "Pub", "PT", "Cat", "PGRP", "ISBN", "Title", "Price", "PubDate",
    "Weekly", "YTD", "LYTD", "YOY", "Last Year",
    "Customer Order", "Units Shipped", "Units YTD", "Units LYTD", "Units Last Year", "% Total Weekly", "% Total YTD", "On Hand", "On Order",
    "Weekly", "YTD", "LYTD", "Last Year", "% Total Weekly", "% Total YTD", "On Hand", "On Order",
    "Weekly", "YTD", "LYTD", "Last Year", "% Total Weekly", "% Total YTD", "On Hand", "On Order",
    "Weekly", "YTD", "LYTD", "Last Year", "% Total Weekly", "% Total YTD", "On Hand", "On Order (NOC)",
    "YTD", "LYTD", "FYLY", "LTD",
]
GROUPS = [
    ("TOTAL SELL-THROUGH", 8, 12, "#EBF1DE"), ("AMAZON", 13, 21, "#FDE9D9"),
    ("B&N", 22, 29, "#EBF1DE"), ("EDELWEISS", 30, 37, "#FDE9D9"),
    ("TARGET", 38, 45, "#EBF1DE"), ("TOTAL SELL-IN", 46, 49, "#DDD9C4"),
]


def report_through_date(freshness):
    return max(freshness[key] for key in ("Amazon", "B&N", "Edelweiss", "Target"))


def save_workbook(report, freshness, output_path: Path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    last_row = len(report) + 6
    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="mm/dd/yyyy") as writer:
        book = writer.book
        sheet = book.add_worksheet("Full List")
        writer.sheets["Full List"] = sheet
        title = book.add_format({"bold": True, "bg_color": "#DCE6F1", "align": "center_across"})
        metadata_block = book.add_format({"bg_color": "#DCE6F1"})
        label = book.add_format({"bold": True, "bg_color": "#EEECE1"})
        number = book.add_format({"num_format": ACCOUNTING})
        percent = book.add_format({"num_format": "0.0%"})
        total_number = book.add_format({"num_format": ACCOUNTING, "bg_color": "#EEECE1"})
        total_percent = book.add_format({"num_format": "0.0%", "bg_color": "#EEECE1"})
        money = book.add_format({"num_format": "$#,##0.00"})
        date_fmt = book.add_format({"num_format": "mm/dd/yyyy"})
        report_through = report_through_date(freshness)

        def center_across(row, start, end, text, fmt):
            sheet.write(row, start, text, fmt)
            for column in range(start + 1, end + 1):
                sheet.write_blank(row, column, None, fmt)

        def grouped_center_across(row, start, end, text, color, *, top=False, bottom=False):
            for column in range(start, end + 1):
                properties = {
                    "bold": True,
                    "bg_color": color,
                    "align": "center_across",
                    "valign": "vcenter",
                }
                if top:
                    properties["top"] = 1
                if bottom:
                    properties["bottom"] = 1
                if column == start:
                    properties["left"] = 1
                if column == end:
                    properties["right"] = 1
                fmt = book.add_format(properties)
                if column == start:
                    sheet.write(row, column, text, fmt)
                else:
                    sheet.write_blank(row, column, None, fmt)

        for row in (3, 4):
            for column in range(8):
                sheet.write_blank(row, column, None, metadata_block)
        center_across(3, 0, 4, f"Chronicle X-Gap through {report_through:%m/%d/%Y}", title)
        freshness_keys = {"AMAZON": "Amazon", "B&N": "B&N", "EDELWEISS": "Edelweiss", "TARGET": "Target", "TOTAL SELL-IN": "Sell-In"}
        for group, start, end, color in GROUPS:
            grouped_center_across(3, start, end, group, color, top=True)
            if group == "TOTAL SELL-THROUGH" or group in freshness_keys:
                value = report_through if group == "TOTAL SELL-THROUGH" else freshness[freshness_keys[group]]
                text = f"Report Run Date {value:%m/%d/%Y}" if group == "TOTAL SELL-IN" else f"Through Week Ending {value:%m/%d/%Y}"
                grouped_center_across(4, start, end, text, color, bottom=True)
        base_header = book.add_format({"bg_color": "#EBF1DE", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
        metric_header = book.add_format({"bg_color": "#DCE6F1", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
        for column, text in enumerate(HEADERS):
            sheet.write(5, column, text, base_header if column < 8 else metric_header)
        for row, values in enumerate(report.itertuples(index=False, name=None), start=6):
            for column, value in enumerate(values):
                fmt = money if column == 6 else date_fmt if column == 7 else percent if column in {18,19,26,27,34,35,42,43} else number if column >= 8 else None
                if pd.isna(value):
                    sheet.write_blank(row, column, None, fmt)
                else:
                    sheet.write(row, column, value, fmt)
        sheet.write(0, 7, "Total", label)
        sheet.write(1, 7, "Subtotal", label)
        ratio_columns = {18: (14, 8), 19: (15, 9), 26: (22, 8), 27: (23, 9), 34: (30, 8), 35: (31, 9), 42: (38, 8), 43: (39, 9)}
        for column in range(8, len(HEADERS)):
            letter = xl_col_to_name(column)
            if column in ratio_columns:
                numerator, denominator = ratio_columns[column]
                n, d = xl_col_to_name(numerator), xl_col_to_name(denominator)
                sheet.write_formula(0, column, f"=IFERROR({n}1/{d}1,0)", total_percent)
                sheet.write_formula(1, column, f"=IFERROR({n}2/{d}2,0)", total_percent)
            else:
                sheet.write_formula(0, column, f"=SUM({letter}7:{letter}{last_row})", total_number)
                sheet.write_formula(1, column, f"=SUBTOTAL(9,{letter}7:{letter}{last_row})", total_number)
        sheet.autofilter(5, 0, max(5, last_row - 1), len(HEADERS) - 1)
        sheet.freeze_panes(6, 8)
        sheet.set_row(5, 32)
        sheet.set_column(0, 4, 13)
        sheet.set_column(5, 5, 38)
        sheet.set_column(6, 7, 12)
        sheet.set_column(8, len(HEADERS) - 1, 12)
        sheet.set_zoom(75)
        sheet.hide_gridlines(2)
    print(f"Saved X-Gap report: {output_path}")


def output_path(freshness):
    latest = report_through_date(freshness)
    week = bookscan_week(latest)
    folder = process_paths.x_gap_output_folder(latest.to_pydatetime())
    return folder / f"Week {week.week} - {week.year} New X-Gap ({week.week_end:%m%d%y}).xlsx"


def run(destination: Path | None = None):
    report, freshness = build_dataframe()
    path = destination or output_path(freshness)
    save_workbook(report, freshness, path)
    print(f"Rows included (non-zero YTD or LYTD sell-through): {len(report):,}")
    print("Target NOC on-order loaded from hachette.HachetteOrders (SSRRowID 140).")
    return path


def main():
    parser = argparse.ArgumentParser(description="Build the combined retailer X-Gap report")
    parser.add_argument("command", nargs="?", choices=("run", "status"), default="run")
    parser.add_argument("--yes", action="store_true")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    rows = source_rows()
    show_sources(rows)
    missing = [label for label, path, _, _ in rows if not path.exists()]
    if missing:
        print("\nCannot build; missing: " + ", ".join(missing))
        return 1
    if args.command == "status":
        return 0
    if not args.yes and input("\nBuild X-Gap using the sources shown above? (y/n): ").strip().lower() not in {"y", "yes"}:
        print("X-Gap build cancelled; no files were changed.")
        return 0
    run(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
