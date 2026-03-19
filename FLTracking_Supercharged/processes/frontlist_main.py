from datetime import datetime
from pathlib import Path
import sys
import tkinter as tk
from tkinter import filedialog

import pandas as pd


sys.path.append(str(Path(__file__).resolve().parents[1]))

from config import OUTPUT_DIR
from isbn_utils import normalize_isbn_series
from processes.amazon_preorders import (
    AMAZON_PREORDERS_PATH,
    load_amazon_preorders,
    load_amazon_preorders_cached,
)
from processes.amazon_sellthrough import load_amazon_sellthrough
from processes.barnes_noble_weekly import (
    build_bn_date_header,
    load_barnes_noble_weekly,
    load_barnes_noble_weekly_cached,
    resolve_barnes_noble_weekly_path,
)
from processes.faire_orders import load_faire_orders
from processes.faire_qty import load_faire_qty
from processes.ingram_daily_report import (
    build_modified_date_header,
    load_ingram_daily_report,
    load_ingram_daily_report_cached,
    resolve_ingram_daily_report_path,
)
from processes.inventory_detail import load_inventory_detail, load_inventory_detail_cached
from processes.inventory_detail import resolve_inventory_detail_path
from processes.amazon_sellthrough import SQL_FILE as AMAZON_SELLTHROUGH_SQL_FILE
from processes.faire_qty import SQL_FILE as FAIRE_QTY_SQL_FILE
from processes.faire_orders import SQL_FILE as FAIRE_ORDERS_SQL_FILE


FRONTLIST_DIR = Path(r"G:\SALES\2026 Sales Reports\Frontlist Tracking")


def resolve_frontlist_tracking_path(source_dir: Path = FRONTLIST_DIR) -> Path:
    files = [path for path in source_dir.glob("*.xlsx") if not path.name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"No Frontlist Tracking workbook found in: {source_dir}")
    return max(files, key=lambda path: path.stat().st_mtime)


def load_frontlist_isbns(source_path: Path | None = None) -> pd.DataFrame:
    resolved_path = source_path or resolve_frontlist_tracking_path()
    df = pd.read_excel(
        resolved_path,
        header=5,
        usecols=["ISBN-13"],
        dtype={"ISBN-13": "object"},
        engine="openpyxl",
    )

    result = df.rename(columns={"ISBN-13": "ISBN"}).copy()
    result["ISBN"] = normalize_isbn_series(result["ISBN"])
    result = result.dropna(subset=["ISBN"]).drop_duplicates(subset=["ISBN"]).reset_index(drop=True)
    return result


def dedupe_source_on_isbn(df: pd.DataFrame) -> pd.DataFrame:
    if df["ISBN"].is_unique:
        return df

    aggregations: dict[str, str] = {}
    for col in df.columns:
        if col == "ISBN":
            continue
        if pd.api.types.is_numeric_dtype(df[col]):
            aggregations[col] = "sum"
        else:
            aggregations[col] = "first"

    return df.groupby("ISBN", as_index=False).agg(aggregations)


def fill_missing_metric_values(df: pd.DataFrame) -> pd.DataFrame:
    protected_cols = {"ISBN", "Reprint Due Date"}
    protected_prefixes = ("IngramUpdated_", "BNUpdated_")

    for col in df.columns:
        if col in protected_cols or col.startswith(protected_prefixes):
            continue

        converted = pd.to_numeric(df[col], errors="coerce")
        original_non_null = df[col].notna().sum()
        converted_non_null = converted.notna().sum()

        if original_non_null == converted_non_null:
            filled = converted.fillna(0)
            if (filled % 1 == 0).all():
                df[col] = filled.astype("Int64")
            else:
                df[col] = filled

    return df


def build_metadata_sheet(frontlist_path: Path, combined: pd.DataFrame) -> pd.DataFrame:
    ingram_path = resolve_ingram_daily_report_path()
    barnes_noble_path = resolve_barnes_noble_weekly_path()
    inventory_detail_path = resolve_inventory_detail_path()
    amazon_preorders_path = AMAZON_PREORDERS_PATH

    _, ingram_report_date = build_modified_date_header(ingram_path)
    _, bn_report_date = build_bn_date_header(barnes_noble_path)

    amz_last_week_value = ""
    if "AmzLastWeek" in combined.columns:
        non_null = combined["AmzLastWeek"].dropna()
        if not non_null.empty:
            amz_last_week_value = pd.to_datetime(non_null.iloc[0]).strftime("%m/%d/%Y")

    rows = [
        {
            "Source": "Frontlist Tracking",
            "FileName": frontlist_path.name,
            "ReportDate": "",
            "ModifiedDate": datetime.fromtimestamp(frontlist_path.stat().st_mtime).strftime("%m/%d/%Y"),
        },
        {
            "Source": "Ingram",
            "FileName": ingram_path.name,
            "ReportDate": ingram_report_date,
            "ModifiedDate": datetime.fromtimestamp(ingram_path.stat().st_mtime).strftime("%m/%d/%Y"),
        },
        {
            "Source": "Barnes & Noble",
            "FileName": barnes_noble_path.name,
            "ReportDate": bn_report_date,
            "ModifiedDate": datetime.fromtimestamp(barnes_noble_path.stat().st_mtime).strftime("%m/%d/%Y"),
        },
        {
            "Source": "Inventory Detail",
            "FileName": inventory_detail_path.name,
            "ReportDate": "",
            "ModifiedDate": datetime.fromtimestamp(inventory_detail_path.stat().st_mtime).strftime("%m/%d/%Y"),
        },
        {
            "Source": "Amazon Preorders",
            "FileName": amazon_preorders_path.name,
            "ReportDate": "",
            "ModifiedDate": datetime.fromtimestamp(amazon_preorders_path.stat().st_mtime).strftime("%m/%d/%Y"),
        },
        {
            "Source": "Amazon Sellthrough SQL",
            "FileName": AMAZON_SELLTHROUGH_SQL_FILE.name,
            "ReportDate": amz_last_week_value,
            "ModifiedDate": datetime.fromtimestamp(AMAZON_SELLTHROUGH_SQL_FILE.stat().st_mtime).strftime("%m/%d/%Y"),
        },
        {
            "Source": "Faire Qty SQL",
            "FileName": FAIRE_QTY_SQL_FILE.name,
            "ReportDate": "",
            "ModifiedDate": datetime.fromtimestamp(FAIRE_QTY_SQL_FILE.stat().st_mtime).strftime("%m/%d/%Y"),
        },
        {
            "Source": "Faire Orders SQL",
            "FileName": FAIRE_ORDERS_SQL_FILE.name,
            "ReportDate": "",
            "ModifiedDate": datetime.fromtimestamp(FAIRE_ORDERS_SQL_FILE.stat().st_mtime).strftime("%m/%d/%Y"),
        },
    ]

    return pd.DataFrame(rows)


def build_frontlist_main() -> tuple[pd.DataFrame, pd.DataFrame, Path]:
    frontlist_path = resolve_frontlist_tracking_path()
    combined = load_frontlist_isbns(frontlist_path)

    source_frames = [
        load_inventory_detail_cached()[0],
        load_amazon_preorders_cached()[0],
        load_amazon_sellthrough(),
        load_faire_qty(),
        load_faire_orders(),
        load_ingram_daily_report_cached()[0],
        load_barnes_noble_weekly_cached()[0],
    ]

    for source_df in source_frames:
        combined = combined.merge(dedupe_source_on_isbn(source_df), on="ISBN", how="left")

    metadata_df = build_metadata_sheet(frontlist_path, combined)
    drop_cols = [
        col for col in combined.columns
        if col == "AmzLastWeek" or col.startswith("IngramUpdated_") or col.startswith("BNUpdated_")
    ]
    if drop_cols:
        combined = combined.drop(columns=drop_cols)

    combined = fill_missing_metric_values(combined)

    return combined, metadata_df, frontlist_path


def save_frontlist_main_output(
    df: pd.DataFrame,
    metadata_df: pd.DataFrame,
    as_of: datetime | None = None,
) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = as_of or datetime.now()
    default_path = OUTPUT_DIR / f"frontlist_main_{timestamp.strftime('%Y_%m_%d')}.xlsx"

    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.asksaveasfilename(
            title="Save Frontlist Main As",
            initialdir=str(OUTPUT_DIR),
            initialfile=default_path.name,
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        root.destroy()
    except Exception:
        selected = ""

    output_path = Path(selected) if selected else default_path
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="FrontlistMain", index=False)
        metadata_df.to_excel(writer, sheet_name="SourceDates", index=False)
    return output_path


def main() -> None:
    df, metadata_df, frontlist_path = build_frontlist_main()
    output_path = save_frontlist_main_output(df, metadata_df)

    print(f"Loaded Frontlist source: {frontlist_path}")
    print(f"Rows in combined output: {len(df)}")
    print(df.head(20).to_string(index=False))
    print("\nSource date summary:")
    print(metadata_df.to_string(index=False))
    print(f"\nSaved output to: {output_path}")


if __name__ == "__main__":
    main()
