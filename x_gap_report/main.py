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
RL = F / "Atelier Readerlink" / "cache" / "readerlink_weekly_sales.parquet"
RL_STORE = F / "Weekly reports" / "2026" / "Readerlink" / "OH_Store_TitlePerformanceReport"
TG = process_paths.TARGET_NOC_CACHE_DIR
TG_SALES, TG_META, TG_INV = TG / "target_noc_weekly_sales.parquet", TG / "target_noc_metadata.parquet", TG / "target_noc_inventory.parquet"
IG = F / "Atelier Ingram"
IG_SALES, IG_INV = IG / "cache" / "ingram_weekly_sales.parquet", IG / "Ingram_OH_OO"
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


def weekly_metrics(df: pd.DataFrame, *, date_col: str, value_col: str, include_last_year: bool = True):
    work = df[["ISBN", date_col, value_col]].copy()
    work["ISBN"] = work["ISBN"].map(normalize_isbn)
    work[date_col] = pd.to_datetime(work[date_col], errors="coerce").dt.normalize()
    work[value_col] = num(work[value_col])
    work = work[work["ISBN"].ne("") & work[date_col].notna()]
    if work.empty:
        raise ValueError(f"No usable rows for {date_col}/{value_col}")
    latest = work[date_col].max()
    current = bookscan_week(latest)
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
    tm, td = weekly_metrics(target, date_col="Week", value_col="Units", include_last_year=False)
    reader = pd.read_parquet(RL, columns=["isbn", "week_end", "master_chain", "cy_pos_units"]).rename(
        columns={"isbn": "ISBN", "week_end": "Week", "cy_pos_units": "Units"})
    chains = reader.master_chain.astype("string").str.strip().str.upper()
    reader = reader[chains.isin({"TARGET", "TARGET STORES"})]
    rm, rd = weekly_metrics(reader, date_col="Week", value_col="Units", include_last_year=False)
    result = tm.add(rm, fill_value=0)

    meta = pd.read_parquet(TG_META, columns=["DPCI", "ISBN"])
    inv = pd.read_parquet(TG_INV)
    inv["Week"] = pd.to_datetime(inv.Week)
    inv = inv[inv.Week.eq(inv.Week.max())].merge(meta, on="DPCI", how="left")
    inv["ISBN"] = inv.ISBN.map(normalize_isbn)
    target_oh = inv.groupby("ISBN")["On Hand"].sum()

    store_file = latest_xlsx(RL_STORE)
    store = pd.read_excel(store_file, sheet_name="Export", usecols=["MASTER CHAIN", "EAN", "TOTAL OH UNITS"], dtype={"EAN": str})
    chains = store["MASTER CHAIN"].astype("string").str.strip().str.upper()
    store = store[chains.isin({"TARGET", "TARGET STORES"})].copy()
    store["ISBN"], store["On Hand"] = store.EAN.map(normalize_isbn), num(store["TOTAL OH UNITS"])
    result["On Hand"] = target_oh.add(store.groupby("ISBN")["On Hand"].sum(), fill_value=0)
    # Filled from hachette.HachetteOrders after the shared SQL query is joined.
    result["On Order (NOC)"] = 0
    return result.fillna(0), td, rd, store_file


def ingram_metrics():
    sales = pd.read_parquet(IG_SALES, columns=["ISBN", "period_end", "gross_qty"]).rename(columns={"period_end": "Week", "gross_qty": "qty"})
    result, latest = weekly_metrics(sales, date_col="Week", value_col="qty", include_last_year=False)
    inv_file = latest_xlsx(IG_INV)
    inv = pd.read_excel(inv_file, usecols=["EAN", "On Hand Total", "On Order Total"], dtype={"EAN": str})
    inv["ISBN"], inv["On Hand"], inv["On Order"] = inv.EAN.map(normalize_isbn), num(inv["On Hand Total"]), num(inv["On Order Total"])
    return result.join(inv.groupby("ISBN")[["On Hand", "On Order"]].sum(), how="outer").fillna(0), latest, inv_file


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
    store = latest_xlsx(RL_STORE) if RL_STORE.exists() else RL_STORE
    ingram_inv = latest_xlsx(IG_INV) if IG_INV.exists() else IG_INV
    ad = amazon_date()
    return [
        ("Amazon customer orders", AMZ_CO, ad, ""), ("Amazon units shipped", AMZ_US, ad, ""),
        ("Amazon on order", AMZ_PO, None, ""), ("Barnes & Noble sales", BN_SALES, max_date(BN_SALES, "Week"), ""),
        ("Barnes & Noble inventory", BN_INV, max_date(BN_INV, "Week"), ""),
        ("Edelweiss sales", ED_SALES, max_date(ED_SALES, "Week"), ""),
        ("Edelweiss inventory", ED_META, max_date(ED_SALES, "Week"), ""),
        ("Readerlink Target sales", RL, max_date(RL, "week_end"), ""),
        ("Readerlink Target store OH", store, None, ""),
        ("Target NOC sales", TG_SALES, max_date(TG_SALES, "Week"), ""),
        ("Target NOC inventory", TG_INV, max_date(TG_INV, "Week"), "no on-order field in cache"),
        ("Ingram sales", IG_SALES, max_date(IG_SALES, "period_end"), ""),
        ("Ingram inventory", ingram_inv, None, ""),
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
    print("Loading Ingram cache and inventory...", flush=True)
    ingram, ingram_d, ingram_file = ingram_metrics()

    all_data = prefixed(amazon, "Amazon").join(prefixed(bn, "BN"), how="outer")
    for frame, prefix in [(ed, "Edelweiss"), (target, "Target"), (ingram, "Ingram")]:
        all_data = all_data.join(prefixed(frame, prefix), how="outer")
    all_data = all_data.fillna(0)
    totals = {
        "Weekly": ["Amazon Units Shipped", "BN Weekly", "Edelweiss Weekly", "Target Weekly", "Ingram Weekly"],
        "YTD": ["Amazon Units YTD", "BN YTD", "Edelweiss YTD", "Target YTD", "Ingram YTD"],
        "LYTD": ["Amazon Units LYTD", "BN LYTD", "Edelweiss LYTD", "Target LYTD", "Ingram LYTD"],
        "Last Year": ["Amazon Units Last Year", "BN Last Year", "Edelweiss Last Year"],
    }
    for metric, columns in totals.items():
        for column in columns:
            if column not in all_data:
                all_data[column] = 0
        all_data[f"Total {metric}"] = all_data[columns].sum(axis=1)
    all_data = all_data[all_data["Total YTD"].ne(0) | all_data["Total LYTD"].ne(0)]
    all_data = item_and_sell_in().join(all_data, how="inner")
    all_data["Target On Order (NOC)"] = all_data["TargetNOCOpenOrder"]
    for account, weekly, ytd in [("Amazon", "Amazon Units Shipped", "Amazon Units YTD"), ("BN", "BN Weekly", "BN YTD"), ("Edelweiss", "Edelweiss Weekly", "Edelweiss YTD")]:
        all_data[f"{account} % Total Weekly"] = all_data[weekly].div(all_data["Total Weekly"].where(all_data["Total Weekly"].ne(0))).fillna(0)
        all_data[f"{account} % Total YTD"] = all_data[ytd].div(all_data["Total YTD"].where(all_data["Total YTD"].ne(0))).fillna(0)

    sources = [
        "Pub", "PT", "Cat", "PGRP", None, "Title", "Price", "PubDate",
        "Total Weekly", "Total YTD", "Total LYTD", "Total Last Year",
        "Amazon Customer Order", "Amazon Units Shipped", "Amazon Units YTD", "Amazon Units LYTD", "Amazon Units Last Year",
        "Amazon % Total Weekly", "Amazon % Total YTD", "Amazon On Hand", "Amazon On Order",
        "BN Weekly", "BN YTD", "BN LYTD", "BN Last Year", "BN % Total Weekly", "BN % Total YTD", "BN On Hand", "BN On Order",
        "Edelweiss Weekly", "Edelweiss YTD", "Edelweiss LYTD", "Edelweiss Last Year", "Edelweiss % Total Weekly", "Edelweiss % Total YTD", "Edelweiss On Hand", "Edelweiss On Order",
        "Target Weekly", "Target YTD", "Target LYTD", "Target On Hand", "Target On Order (NOC)",
        "Ingram Weekly", "Ingram YTD", "Ingram LYTD", "Ingram On Hand", "Ingram On Order",
        "YTD", "LYTD", "FYLY", "LTD",
    ]
    output = pd.DataFrame(index=all_data.index)
    for index, source in enumerate(sources):
        output[index] = all_data.index if source is None else all_data[source]
    output = output.sort_values([0, 3, 5, 4], kind="stable").reset_index(drop=True)
    freshness = {
        "Amazon": amazon_d, "B&N": bn_d, "Edelweiss": ed_d,
        "Target": f"NOC {target_noc_d:%m/%d/%Y}; RL {readerlink_d:%m/%d/%Y}; OH {datetime.fromtimestamp(store_file.stat().st_mtime):%m/%d/%Y}; OO live",
        "Ingram": f"Sales {ingram_d:%m/%d/%Y}; inventory file {datetime.fromtimestamp(ingram_file.stat().st_mtime):%m/%d/%Y}",
        "Sell-In": pd.Timestamp.today().normalize(),
    }
    return output, freshness


HEADERS = [
    "Pub", "PT", "Cat", "PGRP", "ISBN", "Title", "Price", "PubDate",
    "Weekly", "YTD", "LYTD", "Last Year",
    "Customer Order", "Units Shipped", "Units YTD", "Units LYTD", "Units Last Year", "% Total Weekly", "% Total YTD", "On Hand", "On Order",
    "Weekly", "YTD", "LYTD", "Last Year", "% Total Weekly", "% Total YTD", "On Hand", "On Order",
    "Weekly", "YTD", "LYTD", "Last Year", "% Total Weekly", "% Total YTD", "On Hand", "On Order",
    "Weekly", "YTD", "LYTD", "On Hand", "On Order (NOC)",
    "Weekly", "YTD", "LYTD", "On Hand", "On Order",
    "YTD", "LYTD", "FYLY", "LTD",
]
GROUPS = [
    ("TOTAL SELL-THROUGH", 8, 11, "#B4C6E7"), ("AMAZON", 12, 20, "#F4B183"),
    ("B&N", 21, 28, "#FFE699"), ("EDELWEISS", 29, 36, "#C6E0B4"),
    ("TARGET", 37, 41, "#D9EAD3"), ("INGRAM", 42, 46, "#D9E1F2"),
    ("TOTAL SELL-IN", 47, 50, "#D9D2E9"),
]


def save_workbook(report, freshness, output_path: Path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    last_row = len(report) + 6
    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="mm/dd/yyyy") as writer:
        book = writer.book
        sheet = book.add_worksheet("Full List")
        writer.sheets["Full List"] = sheet
        title = book.add_format({"bold": True, "font_size": 14, "bg_color": "#D9EAD3", "border": 1})
        label = book.add_format({"bold": True, "bg_color": "#DDD9C4", "border": 1})
        number = book.add_format({"num_format": ACCOUNTING})
        percent = book.add_format({"num_format": "0.0%"})
        money = book.add_format({"num_format": "$#,##0.00"})
        date_fmt = book.add_format({"num_format": "mm/dd/yyyy"})
        dates = [v for v in freshness.values() if isinstance(v, pd.Timestamp)]
        week = bookscan_week(max(dates) if dates else pd.Timestamp.today())
        sheet.write(3, 0, f"Chronicle X-Gap {week.week_start:%m/%d/%y} to {week.week_end:%m/%d/%y}", title)
        freshness_keys = {"AMAZON": "Amazon", "B&N": "B&N", "EDELWEISS": "Edelweiss", "TARGET": "Target", "INGRAM": "Ingram", "TOTAL SELL-IN": "Sell-In"}
        header_formats = {}
        for group, start, end, color in GROUPS:
            fmt = book.add_format({"bold": True, "bg_color": color, "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
            sheet.merge_range(3, start, 3, end, group, fmt)
            if group in freshness_keys:
                value = freshness[freshness_keys[group]]
                text = f"Through {value:%m/%d/%Y}" if isinstance(value, pd.Timestamp) else str(value)
                sheet.merge_range(4, start, 4, end, text, fmt)
            for column in range(start, end + 1):
                header_formats[column] = fmt
        base_header = book.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1, "align": "center"})
        for column, text in enumerate(HEADERS):
            sheet.write(5, column, text, header_formats.get(column, base_header))
        for row, values in enumerate(report.itertuples(index=False, name=None), start=6):
            for column, value in enumerate(values):
                fmt = money if column == 6 else date_fmt if column == 7 else percent if column in {17,18,25,26,33,34} else number if column >= 8 else None
                if pd.isna(value):
                    sheet.write_blank(row, column, None, fmt)
                else:
                    sheet.write(row, column, value, fmt)
        sheet.write(0, 7, "Total", label)
        sheet.write(1, 7, "Subtotal", label)
        ratio_columns = {17: (13, 8), 18: (14, 9), 25: (21, 8), 26: (22, 9), 33: (29, 8), 34: (30, 9)}
        for column in range(8, len(HEADERS)):
            letter = xl_col_to_name(column)
            if column in ratio_columns:
                numerator, denominator = ratio_columns[column]
                n, d = xl_col_to_name(numerator), xl_col_to_name(denominator)
                sheet.write_formula(0, column, f"=IFERROR({n}1/{d}1,0)", percent)
                sheet.write_formula(1, column, f"=IFERROR({n}2/{d}2,0)", percent)
            else:
                sheet.write_formula(0, column, f"=SUM({letter}7:{letter}{last_row})", number)
                sheet.write_formula(1, column, f"=SUBTOTAL(9,{letter}7:{letter}{last_row})", number)
        sheet.autofilter(5, 0, max(5, last_row - 1), len(HEADERS) - 1)
        sheet.freeze_panes(6, 7)
        sheet.set_row(5, 32)
        sheet.set_column(0, 4, 13)
        sheet.set_column(5, 5, 38)
        sheet.set_column(6, 7, 12)
        sheet.set_column(8, len(HEADERS) - 1, 12)
        sheet.hide_gridlines(2)
    print(f"Saved X-Gap report: {output_path}")


def output_path(freshness):
    dates = [v for v in freshness.values() if isinstance(v, pd.Timestamp)]
    latest = max(dates) if dates else pd.Timestamp.today().normalize()
    week = bookscan_week(latest)
    folder = process_paths.x_gap_output_folder(latest.to_pydatetime())
    return folder / f"Week {week.week} - {week.year} X-Gap ({week.week_end:%m%d%y}).xlsx"


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
