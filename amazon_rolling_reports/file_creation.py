from pathlib import Path
from functools import lru_cache

import pandas as pd
pd.set_option('future.no_silent_downcasting', True)

from paths import amazon_po_pickle_file, customer_orders_pickle_file, monthly_sales_parquet_file, oracle_ypticod_file
from shared.pg_grouping import apply_pg_grouping


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


def normalize_asin(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    if not text or text.lower() == "nan":
        return ""
    return text.zfill(10)


@lru_cache(maxsize=1)
def load_monthly_asin_map(monthly_sales_parquet: str | Path = monthly_sales_parquet_file) -> pd.Series:
    parquet_path = Path(monthly_sales_parquet)
    if not parquet_path.exists():
        return pd.Series(dtype="string")

    df = pd.read_parquet(parquet_path, columns=["Period", "ISBN", "ASIN"])
    df = df[df["ISBN"].notna() & df["ASIN"].notna()].copy()
    df = df[df["ISBN"].astype(str).str.strip() != "NO_ISBN"]
    df["ISBN"] = df["ISBN"].map(normalize_isbn)
    df["ASIN"] = df["ASIN"].map(normalize_asin)
    df["Period"] = df["Period"].astype(str).str.strip()
    df = df[(df["ISBN"] != "") & (df["ASIN"] != "")]
    df = df.sort_values(["ISBN", "Period"], ascending=[True, False])
    return df.drop_duplicates(subset="ISBN", keep="first").set_index("ISBN")["ASIN"]


@lru_cache(maxsize=1)
def load_ypticod_asin_map(ypticod_file: str | Path = oracle_ypticod_file) -> pd.Series:
    ypticod_path = Path(ypticod_file)
    if not ypticod_path.exists():
        return pd.Series(dtype="string")

    df = pd.read_excel(ypticod_path, usecols=["ISBN", "ISBN10", "Publisher Name"])
    publisher_delete_list = [
        "Princeton Architectural Press",
        "AFO LLC",
        "Benefit",
        "Driscolls",
        "FareArts",
        "Moleskine",
        "No Publisher Name",
        "PQ Blackwell",
        "Sager",
        "San Francisco Art Institute",
        "Glam Media",
    ]
    if "Publisher Name" in df.columns:
        df = df[~df["Publisher Name"].isin(publisher_delete_list)]
    df["ISBN"] = df["ISBN"].map(normalize_isbn)
    df["ASIN"] = df["ISBN10"].map(normalize_asin)
    df = df[(df["ISBN"] != "") & (df["ASIN"] != "")]
    df = df[~df["ISBN"].duplicated(keep=False)]
    df = df[~df["ASIN"].duplicated(keep=False)]
    return df.drop_duplicates(subset="ISBN", keep="first").set_index("ISBN")["ASIN"]


def load_asin_map() -> pd.Series:
    monthly_map = load_monthly_asin_map()
    ypticod_map = load_ypticod_asin_map()
    if monthly_map.empty:
        return ypticod_map
    if ypticod_map.empty:
        return monthly_map
    return monthly_map.combine_first(ypticod_map)


def add_asin_before_isbn(df: pd.DataFrame) -> pd.DataFrame:
    if "ISBN" not in df.columns:
        return df

    output = df.copy()
    isbn_values = output["ISBN"].map(normalize_isbn)
    asin_map = load_asin_map()
    output["ASIN"] = isbn_values.map(asin_map).fillna("")
    missing_asin_count = int(output["ASIN"].astype(str).str.strip().eq("").sum())
    if missing_asin_count:
        print(f"Warning: {missing_asin_count:,} Amazon rolling row(s) still have no ASIN mapping.")

    columns = list(output.columns)
    columns.remove("ASIN")
    isbn_index = columns.index("ISBN")
    columns.insert(isbn_index, "ASIN")
    return output.loc[:, columns]


def prepare_weekly_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if "ISBN" not in df.columns:
        return df

    output = df.copy()
    output["ISBN"] = output["ISBN"].map(normalize_isbn)
    before = len(output)
    output = output[output["ISBN"] != ""]
    output = output.drop_duplicates(subset="ISBN", keep="first")
    dropped = before - len(output)
    if dropped:
        print(f"Removed {dropped:,} duplicate/blank ISBN row(s) from weekly Amazon data before PO merge.")
    return apply_pg_grouping(output)


def prepare_po_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if "ISBN" not in df.columns:
        return pd.DataFrame(columns=["ISBN", "PO_Qty"])

    output = df.copy()
    output["ISBN"] = output["ISBN"].map(normalize_isbn)
    output = output[output["ISBN"] != ""]
    if "PO_Qty" not in output.columns and "Accepted Quantity" in output.columns:
        output = output.rename(columns={"Accepted Quantity": "PO_Qty"})
    output["PO_Qty"] = pd.to_numeric(output.get("PO_Qty", 0), errors="coerce").fillna(0)
    before = len(output)
    output = output.groupby("ISBN", as_index=False)["PO_Qty"].sum()
    combined = before - len(output)
    if combined:
        print(f"Combined {combined:,} duplicate Amazon PO row(s) by ISBN before weekly merge.")
    return output

def create_rolling_report(pickle_file,pickle_po):
    df_co = prepare_weekly_dataframe(pd.read_pickle(pickle_file))
    df_po = prepare_po_dataframe(pd.read_pickle(pickle_po))
    df_combined = pd.merge(df_co, df_po, how='left', left_on='ISBN', right_on='ISBN')
    df_combined['PO_Qty'] = pd.to_numeric(df_combined['PO_Qty'], errors='coerce').fillna(0).astype(int)

    if 'PO_Qty' in df_combined.columns and 'AvgLast6W' in df_combined.columns:
        df_combined['AvgLast6W'] = pd.to_numeric(df_combined['AvgLast6W'], errors='coerce').fillna(0)
        divisor = df_combined['AvgLast6W'].where(df_combined['AvgLast6W'].ne(0))
        df_combined['OH_Avg'] = (df_combined['PO_Qty'] / divisor).round(2)
        df_combined['OH_Avg'] = df_combined['OH_Avg'].fillna(0).infer_objects(copy=False)
    else:
        print("PO_Qty or AvgLast6W column missing!")
        return df_combined

    if 'TYTD' in df_combined.columns and 'LYTD' in df_combined.columns:
        tytd = pd.to_numeric(df_combined['TYTD'], errors='coerce').fillna(0)
        lytd = pd.to_numeric(df_combined['LYTD'], errors='coerce').fillna(0)
        df_combined['YTD Var'] = tytd - lytd

    # Reorder columns: place PO_Qty after PubDate and before OH, and OH_Avg after OH
    cols = list(df_combined.columns)
    if 'PO_Qty' in cols and 'PubDate' in cols and 'OH' in cols:
        cols.remove('PO_Qty')
        pubdate_index = cols.index('PubDate')
        cols.insert(pubdate_index + 1, 'PO_Qty')
    if 'OH_Avg' in cols and 'OH' in cols:
        cols.remove('OH_Avg')
        oh_index = cols.index('OH')
        cols.insert(oh_index + 1, 'OH_Avg')
    if 'YTD Var' in cols and 'LYTD' in cols:
        cols.remove('YTD Var')
        lytd_index = cols.index('LYTD')
        cols.insert(lytd_index + 1, 'YTD Var')
    df_combined = df_combined[cols]
    if 'AvgLast6W' in df_combined.columns:
        df_combined = df_combined.rename(columns={'AvgLast6W': '6Wk Avg'})
    return add_asin_before_isbn(df_combined)


MONTHLY_BASE_COLUMNS = [
    "Pub",
    "pt",
    "ft",
    "pgrp",
    "PG_Grouping",
    "ASIN",
    "ISBN",
    "Title",
    "Price",
    "PubDate",
]
MONTHLY_SUMMARY_COLUMNS = ["12M", "TYTD", "LYTD", "YTD Var", "LY_FY", "Total"]
MIN_MONTHLY_REPORT_PERIOD = "202401"


def _period_year(period: str) -> int:
    return int(str(period)[:4])


def _period_month(period: str) -> int:
    return int(str(period)[4:6])


def _through_label(period: str) -> str:
    date_value = pd.to_datetime(f"{period}01", format="%Y%m%d")
    return f"Through {date_value.strftime('%B %Y')}"


def monthly_through_label(monthly_df: pd.DataFrame) -> str:
    periods = sorted(
        column for column in monthly_df.columns
        if isinstance(column, str) and len(column) == 6 and column.isdigit()
    )
    if not periods:
        return "Through"
    return _through_label(periods[-1])


def create_monthly_rolling_report(
    weekly_df: pd.DataFrame,
    monthly_sales_parquet: str | Path,
    units_column: str,
) -> pd.DataFrame:
    monthly_sales = pd.read_parquet(monthly_sales_parquet)
    monthly_sales = monthly_sales[monthly_sales["ISBN"].notna()].copy()
    monthly_sales = monthly_sales[monthly_sales["ISBN"].astype(str).str.strip() != "NO_ISBN"]
    monthly_sales["ISBN"] = monthly_sales["ISBN"].astype(str).str.strip().str.zfill(13)
    monthly_sales["Period"] = monthly_sales["Period"].astype(str).str.strip()
    monthly_sales = monthly_sales[monthly_sales["Period"] >= MIN_MONTHLY_REPORT_PERIOD]
    monthly_sales[units_column] = pd.to_numeric(monthly_sales[units_column], errors="coerce").fillna(0)

    periods = sorted(monthly_sales["Period"].dropna().unique(), reverse=True)
    if not periods:
        raise ValueError(
            f"No monthly sales periods from {MIN_MONTHLY_REPORT_PERIOD} onward found in {monthly_sales_parquet}"
        )

    monthly_units = (
        monthly_sales.groupby(["ISBN", "Period"], as_index=False)[units_column]
        .sum()
        .pivot(index="ISBN", columns="Period", values=units_column)
        .fillna(0)
    )
    monthly_units = monthly_units.reindex(columns=periods, fill_value=0)

    metadata = weekly_df[[column for column in MONTHLY_BASE_COLUMNS if column in weekly_df.columns]].copy()
    metadata["ISBN"] = metadata["ISBN"].astype(str).str.replace(r"\.0$", "", regex=True).str.zfill(13)
    if "ASIN" not in metadata.columns:
        metadata["ASIN"] = metadata["ISBN"].map(load_monthly_asin_map()).fillna("")
    metadata = metadata.drop_duplicates(subset="ISBN", keep="first").set_index("ISBN")

    monthly_only_isbns = monthly_units.index.difference(metadata.index)
    if len(monthly_only_isbns):
        missing_metadata = pd.DataFrame(
            {column: "" for column in metadata.columns},
            index=monthly_only_isbns,
        )
        missing_metadata.index.name = "ISBN"
        metadata = pd.concat([metadata, missing_metadata], axis=0)

    report = metadata.join(monthly_units, how="right").reset_index()
    current_period = periods[0]
    current_year = _period_year(current_period)
    current_month = _period_month(current_period)
    prior_year = current_year - 1

    last_12_periods = periods[:12]
    tytd_periods = [
        period for period in periods
        if _period_year(period) == current_year and _period_month(period) <= current_month
    ]
    lytd_periods = [
        period for period in periods
        if _period_year(period) == prior_year and _period_month(period) <= current_month
    ]
    ly_fy_periods = [period for period in periods if _period_year(period) == prior_year]

    report["12M"] = report[last_12_periods].sum(axis=1)
    report["TYTD"] = report[tytd_periods].sum(axis=1)
    report["LYTD"] = report[lytd_periods].sum(axis=1)
    report["YTD Var"] = report["TYTD"] - report["LYTD"]
    report["LY_FY"] = report[ly_fy_periods].sum(axis=1)
    report["Total"] = report[periods].sum(axis=1)

    for column in MONTHLY_BASE_COLUMNS:
        if column not in report.columns:
            report[column] = ""

    report = report[MONTHLY_BASE_COLUMNS + MONTHLY_SUMMARY_COLUMNS + periods]
    report = report.sort_values(by=current_period, ascending=False)
    return report

def main():
    pickle_file1 = customer_orders_pickle_file
    pickle_po = amazon_po_pickle_file
    
    df_combined = create_rolling_report(pickle_file1,pickle_po)
    print(df_combined.shape)
    print(df_combined.columns[:20])
    print(df_combined.head())

if __name__ == "__main__":
    main()
