from __future__ import annotations

import argparse
import csv
from importlib.util import module_from_spec, spec_from_file_location
import re
import sys
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

AMAZON_SQL_UPLOAD_DIR = REPO_ROOT / "amazon_sql_upload"
if str(AMAZON_SQL_UPLOAD_DIR) not in sys.path:
    sys.path.insert(0, str(AMAZON_SQL_UPLOAD_DIR))

from asin_manual_key import asin_isbn_manual_key  # noqa: E402
from asin_removal_list import asins_to_delete_list  # noqa: E402
from load_catalog import df_catalog  # noqa: E402
from load_ebs_isbn_key import isbn_key  # noqa: E402
from load_ypticod import load_ypticod  # noqa: E402


def load_shared_process_paths():
    shared_path = REPO_ROOT / "paths" / "process_paths.py"
    spec = spec_from_file_location("_shared_process_paths", shared_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load shared process paths from {shared_path}")
    module = module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


process_paths = load_shared_process_paths()

MONTHLY_SALES_FILE_PATTERN = re.compile(
    r"^Sales_ASIN_Manufacturing_Retail_UnitedStates_Monthly_"
    r"(?P<start>\d{1,2}-\d{1,2}-\d{4})_(?P<end>\d{1,2}-\d{1,2}-\d{4})\.csv$",
    re.IGNORECASE,
)
SALES_COLUMNS = [
    "ASIN",
    "Product Title",
    "Ordered Revenue",
    "Ordered Units",
    "Shipped Revenue",
    "Shipped Units",
    "Shipped COGS",
    "Customer Returns",
]
METADATA_COLUMNS = [
    "Period",
    "Report Year",
    "Report Month",
    "Report Start Date",
    "Report End Date",
    "Source File",
]
OUTPUT_COLUMNS = [
    *METADATA_COLUMNS,
    "ASIN",
    "ISBN",
    "Product Title",
    "Ordered Revenue",
    "Ordered Units",
    "Shipped Revenue",
    "Shipped Units",
    "Shipped COGS",
    "Customer Returns",
]
NUMERIC_COLUMNS = [
    "Ordered Revenue",
    "Ordered Units",
    "Shipped Revenue",
    "Shipped Units",
    "Shipped COGS",
    "Customer Returns",
]


@dataclass(frozen=True)
class MonthlySalesFile:
    path: Path
    period: str
    report_year: int
    report_month: int
    start_date: pd.Timestamp
    end_date: pd.Timestamp


def resolve_source_root(source_root: Path | None = None) -> Path:
    if source_root is not None:
        return source_root
    preferred = process_paths.AMAZON_MONTHLY_SALES_ROOT
    if preferred.exists():
        return preferred
    fallback = process_paths.AMAZON_MONTHLY_SALES_FALLBACK_ROOT
    if fallback.exists():
        return fallback
    return preferred


def output_file_for_source(source_root: Path, output_file: Path | None = None) -> Path:
    if output_file is not None:
        return output_file
    return source_root / "cache" / process_paths.AMAZON_MONTHLY_SALES_PARQUET_NAME


def parse_monthly_sales_file(path: Path) -> MonthlySalesFile | None:
    match = MONTHLY_SALES_FILE_PATTERN.match(path.name)
    if not match:
        return None

    start_date = pd.to_datetime(match.group("start"), format="%m-%d-%Y", errors="raise")
    end_date = pd.to_datetime(match.group("end"), format="%m-%d-%Y", errors="raise")
    period = f"{start_date.year}{start_date.month:02d}"
    return MonthlySalesFile(
        path=path,
        period=period,
        report_year=start_date.year,
        report_month=start_date.month,
        start_date=pd.Timestamp(start_date).normalize(),
        end_date=pd.Timestamp(end_date).normalize(),
    )


def discover_monthly_sales_files(source_root: Path) -> list[MonthlySalesFile]:
    if not source_root.exists():
        raise FileNotFoundError(f"Amazon monthly sales folder not found: {source_root}")
    if not source_root.is_dir():
        raise NotADirectoryError(f"Amazon monthly sales path is not a folder: {source_root}")

    files_by_period: dict[str, MonthlySalesFile] = {}
    for path in sorted(source_root.glob("*.csv"), key=lambda item: item.name.lower()):
        if path.name.startswith("~$"):
            continue
        parsed = parse_monthly_sales_file(path)
        if parsed is None:
            continue
        existing = files_by_period.get(parsed.period)
        if existing is None or parsed.path.stat().st_mtime > existing.path.stat().st_mtime:
            files_by_period[parsed.period] = parsed
    return sorted(files_by_period.values(), key=lambda item: item.period)


def detect_csv_header_row(file_path: Path, max_rows: int = 12) -> int:
    expected = {column.lower() for column in SALES_COLUMNS}
    best_row = 0
    best_score = 0

    with file_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        for row_idx, row in enumerate(reader):
            if row_idx >= max_rows:
                break
            values = {str(value).strip().lower() for value in row if str(value).strip()}
            score = len(values & expected)
            if score > best_score:
                best_score = score
                best_row = int(row_idx)
            if score == len(expected):
                break
    return best_row


def clean_numeric_column(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype("string")
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def read_monthly_sales_file(monthly_file: MonthlySalesFile) -> tuple[pd.DataFrame, list[str]]:
    header_row = detect_csv_header_row(monthly_file.path)
    df = pd.read_csv(monthly_file.path, header=header_row, dtype=object, encoding="utf-8-sig")
    df.columns = [str(column).strip() for column in df.columns]

    missing = [column for column in SALES_COLUMNS if column not in df.columns]
    if "ASIN" not in df.columns:
        raise ValueError(f"ASIN column is required in {monthly_file.path}")

    present_columns = [column for column in SALES_COLUMNS if column in df.columns]
    df = df.loc[:, present_columns].copy()
    for column in SALES_COLUMNS:
        if column not in df.columns:
            df[column] = 0 if column in NUMERIC_COLUMNS else ""

    df["ASIN"] = df["ASIN"].astype("string").str.strip().str.zfill(10)
    df["Product Title"] = df["Product Title"].astype("string").str.strip()
    for column in NUMERIC_COLUMNS:
        df[column] = clean_numeric_column(df[column])

    df.insert(0, "Source File", monthly_file.path.name)
    df.insert(0, "Report End Date", monthly_file.end_date)
    df.insert(0, "Report Start Date", monthly_file.start_date)
    df.insert(0, "Report Month", monthly_file.report_month)
    df.insert(0, "Report Year", monthly_file.report_year)
    df.insert(0, "Period", monthly_file.period)
    return df, missing


def normalize_isbn(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    digits = "".join(char for char in text if char.isdigit())
    if not digits:
        return ""
    if len(digits) < 13:
        return digits.zfill(13)
    if len(digits) > 13 and digits.startswith("0"):
        return digits[-13:]
    return digits[:13]


def safe_load_catalog() -> pd.DataFrame:
    try:
        catalog = df_catalog()
    except Exception as exc:
        print(f"Warning: could not load latest Amazon catalog for ISBN fallback: {exc}")
        return pd.DataFrame(columns=["ASIN", "EAN", "ISBN_Amz", "Model Number"])
    for column in ["ASIN", "EAN", "ISBN_Amz", "Model Number"]:
        if column not in catalog.columns:
            catalog[column] = ""
    return catalog[["ASIN", "EAN", "ISBN_Amz", "Model Number"]].drop_duplicates(subset=["ASIN"])


def safe_load_ebs_isbn_key() -> pd.DataFrame:
    try:
        df_isbn = isbn_key()
    except Exception as exc:
        print(f"Warning: could not load EBS ISBN key for ISBN validation: {exc}")
        return pd.DataFrame(columns=["ISBN"])
    if "ISBN" not in df_isbn.columns:
        return pd.DataFrame(columns=["ISBN"])
    df_isbn = df_isbn.copy()
    df_isbn["ISBN"] = df_isbn["ISBN"].map(normalize_isbn)
    return df_isbn


def add_isbn(df: pd.DataFrame) -> pd.DataFrame:
    df_ypticod = load_ypticod()
    df_isbn = safe_load_ebs_isbn_key()
    df_catalog = safe_load_catalog()

    isbn_set = set(df_isbn["ISBN"].dropna().astype(str).unique().tolist())
    isbn_set.discard("")
    isbn_set.discard("ISBN")

    working = df.copy()
    existing_isbn = (
        working.pop("ISBN").map(normalize_isbn)
        if "ISBN" in working.columns
        else pd.Series("", index=working.index, dtype="string")
    )

    merged = working.merge(df_ypticod, on="ASIN", how="left")
    merged = merged.merge(df_catalog, on="ASIN", how="left")

    isbn_col = pd.Series(existing_isbn.to_numpy(), index=merged.index).fillna("")
    ypticod_values = merged["ISBN"].map(normalize_isbn)
    mask = isbn_col == ""
    isbn_col = pd.Series(np.where(mask, ypticod_values, isbn_col), index=merged.index)
    for fallback_column in ["EAN", "ISBN_Amz", "Model Number"]:
        fallback_values = merged[fallback_column].map(normalize_isbn)
        mask = (isbn_col == "") & fallback_values.isin(isbn_set)
        isbn_col = pd.Series(np.where(mask, fallback_values, isbn_col), index=merged.index)

    merged["ISBN"] = np.where(isbn_col == "", "NO_ISBN", isbn_col)
    merged = merged[~merged["ASIN"].isin(asins_to_delete_list)]
    merged = merged[~merged["Product Title"].fillna("").str.lower().str.endswith("anglais")]
    merged["ISBN"] = merged.apply(lambda row: asin_isbn_manual_key.get(row["ASIN"], row["ISBN"]), axis=1)
    return merged


def standardize_for_parquet(df: pd.DataFrame) -> pd.DataFrame:
    standardized = df.copy()
    standardized["Period"] = standardized["Period"].astype("string").str.strip()
    standardized["Report Year"] = pd.to_numeric(standardized["Report Year"], errors="coerce").astype("Int64")
    standardized["Report Month"] = pd.to_numeric(standardized["Report Month"], errors="coerce").astype("Int64")
    standardized["Report Start Date"] = pd.to_datetime(standardized["Report Start Date"], errors="coerce")
    standardized["Report End Date"] = pd.to_datetime(standardized["Report End Date"], errors="coerce")
    for column in ["ASIN", "ISBN", "Product Title", "Source File"]:
        standardized[column] = standardized[column].astype("string").str.strip()
    for column in NUMERIC_COLUMNS:
        standardized[column] = pd.to_numeric(standardized[column], errors="coerce").fillna(0)
    return standardized


def read_existing_history(history_file: Path, before_period: str) -> pd.DataFrame:
    if not history_file.exists():
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    legacy = pd.read_parquet(history_file)
    missing_columns = [column for column in OUTPUT_COLUMNS if column not in legacy.columns]
    for column in missing_columns:
        legacy[column] = 0 if column in NUMERIC_COLUMNS else ""

    legacy = legacy.loc[:, OUTPUT_COLUMNS].copy()
    legacy["Period"] = legacy["Period"].astype("string").str.strip()
    legacy = legacy[legacy["Period"] < before_period]
    if legacy.empty:
        return legacy

    legacy = add_isbn(legacy)
    legacy = legacy.loc[:, OUTPUT_COLUMNS]
    return standardize_for_parquet(legacy)


def build_monthly_sales(
    source_root: Path | None = None,
    output_file: Path | None = None,
    preserve_history: bool = True,
) -> tuple[pd.DataFrame, dict[str, list[str]], Path, Path]:
    resolved_source_root = resolve_source_root(source_root)
    resolved_output_file = output_file_for_source(resolved_source_root, output_file)
    monthly_files = discover_monthly_sales_files(resolved_source_root)
    if not monthly_files:
        raise FileNotFoundError(f"No Amazon monthly sales CSV files found under {resolved_source_root}")

    frames: list[pd.DataFrame] = []
    missing_by_file: dict[str, list[str]] = {}
    for monthly_file in monthly_files:
        df, missing = read_monthly_sales_file(monthly_file)
        frames.append(df)
        if missing:
            missing_by_file[monthly_file.path.name] = missing

    combined = pd.concat(frames, ignore_index=True, sort=False)
    combined = add_isbn(combined)
    combined = combined.loc[:, OUTPUT_COLUMNS]
    combined = standardize_for_parquet(combined)

    if preserve_history:
        first_new_period = str(combined["Period"].min())
        legacy = read_existing_history(resolved_output_file, first_new_period)
        if not legacy.empty:
            combined = pd.concat([legacy, combined], ignore_index=True, sort=False)
            combined = combined.sort_values(["Period", "ASIN", "Source File"], kind="stable").reset_index(drop=True)
            combined = standardize_for_parquet(combined)

    resolved_output_file.parent.mkdir(parents=True, exist_ok=True)
    combined.to_parquet(resolved_output_file, index=False, engine="pyarrow")
    return combined, missing_by_file, resolved_source_root, resolved_output_file


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Compile Amazon monthly sales CSVs to parquet.")
    parser.add_argument("--source-root", type=Path, help="Monthly sales CSV folder.")
    parser.add_argument("--output", type=Path, help="Output parquet file.")
    parser.add_argument("--no-history", action="store_true", help="Compile only the CSV folder and do not preserve older periods from the existing parquet.")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    combined, missing_by_file, source_root, output_file = build_monthly_sales(
        source_root=args.source_root,
        output_file=args.output,
        preserve_history=not args.no_history,
    )

    print("Saved Amazon monthly sales parquet.")
    print(f"  Source root: {source_root}")
    print(f"  Output file: {output_file}")
    print(f"  Rows:        {len(combined):,}")
    print(f"  Columns:     {len(combined.columns):,}")
    print(f"  Periods:     {', '.join(sorted(combined['Period'].dropna().astype(str).unique()))}")
    no_isbn_count = int((combined["ISBN"] == "NO_ISBN").sum())
    print(f"  NO_ISBN:     {no_isbn_count:,}")
    if missing_by_file:
        print("Files missing expected monthly sales columns:")
        for filename, missing in missing_by_file.items():
            print(f"  - {filename}: {', '.join(missing)}")


if __name__ == "__main__":
    main()
