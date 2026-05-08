from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from paths import process_paths  # noqa: E402


CONFIG_FILE = Path(__file__).with_name("monthly_customer_orders_config.json")
DEFAULT_CONFIG = {
    "start_year": 2023,
    "end_year": None,
    "included_periods": [],
    "excluded_periods": [],
}
MONTHLY_FILE_PATTERN = re.compile(
    r"^(?P<updated>UPDATED_)?(?P<use>USE_)?(?:(?P<month>\d{2})_)?Sales_(?:ASIN_)?"
    r"Manufacturing_Retail_UnitedStates_Monthly_"
    r"(?P<start>\d{1,2}-\d{1,2}-\d{4})_(?P<end>\d{1,2}-\d{1,2}-\d{4})"
    r"(?: \(\d+\))?\.(?:xlsx|csv)$",
    re.IGNORECASE,
)
SELL_THROUGH_FILE_PATTERN = re.compile(
    r"^SELL THROUGH_(?P<year>20\d{2})_(?P<month>\d{2})_[A-Z]{3}\.xlsx$",
    re.IGNORECASE,
)
IMPORTANT_COLUMNS = [
    "ASIN",
    "ISBN",
    "Ordered Revenue",
    "Ordered Units",
    "Shipped Revenue",
    "Shipped COGS",
    "Shipped Units",
    "Unfilled Customer Ordered Units",
    "Confirmed Units",
    "Net Shipped GMS",
    "Binding",
]
METADATA_COLUMNS = [
    "Period",
    "Report Year",
    "Report Month",
    "Report Start Date",
    "Report End Date",
    "Source File",
    "Source Sheet",
]
TEXT_LIKE_COLUMNS = {
    "Period",
    "ASIN",
    "ISBN",
    "Product Title",
    "Brand Code",
    "Format (From YPTICOD)",
    "AMZ GL",
    "Brand",
    "Category",
    "Subcategory",
    "Parent ASIN",
    "UPC",
    "EAN",
    "ISBN.1",
    "Model Number",
    "Manufacturer Code",
    "Binding",
    "Color",
    "Replenishment Code",
    "Source File",
    "Source Sheet",
}
DATE_LIKE_COLUMNS = {"Report Start Date", "Report End Date", "Release Date"}
COLUMN_ALIASES = {
    "ISBN-13": "ISBN",
}


@dataclass(frozen=True)
class MonthlyFile:
    path: Path
    period: str
    report_year: int
    report_month: int
    start_date: pd.Timestamp
    end_date: pd.Timestamp
    is_updated: bool


def normalize_period(period: str) -> str:
    text = str(period).strip()
    if not re.fullmatch(r"20\d{4}", text):
        raise ValueError("Period must use yyyymm format, such as 202604.")
    month = int(text[-2:])
    if month < 1 or month > 12:
        raise ValueError("Period month must be between 01 and 12.")
    return text


def load_config() -> dict[str, object]:
    if not CONFIG_FILE.exists():
        return dict(DEFAULT_CONFIG)
    with CONFIG_FILE.open("r", encoding="utf-8") as config_file:
        raw_config = json.load(config_file)
    config = dict(DEFAULT_CONFIG)
    config.update(raw_config)
    config["included_periods"] = sorted({normalize_period(period) for period in config.get("included_periods", [])})
    config["excluded_periods"] = sorted({normalize_period(period) for period in config.get("excluded_periods", [])})
    return config


def save_config(config: dict[str, object]) -> None:
    normalized = dict(DEFAULT_CONFIG)
    normalized.update(config)
    normalized["included_periods"] = sorted({normalize_period(period) for period in normalized.get("included_periods", [])})
    normalized["excluded_periods"] = sorted({normalize_period(period) for period in normalized.get("excluded_periods", [])})
    CONFIG_FILE.write_text(json.dumps(normalized, indent=2) + "\n", encoding="utf-8")


def include_period(period: str) -> dict[str, object]:
    normalized_period = normalize_period(period)
    config = load_config()
    included = set(config.get("included_periods", []))
    excluded = set(config.get("excluded_periods", []))
    excluded.discard(normalized_period)
    start_year = config.get("start_year")
    end_year = config.get("end_year")
    period_year = int(normalized_period[:4])
    if (
        not isinstance(start_year, int)
        or period_year < start_year
        or (isinstance(end_year, int) and period_year > end_year)
    ):
        included.add(normalized_period)
    config["included_periods"] = sorted(included)
    config["excluded_periods"] = sorted(excluded)
    save_config(config)
    return config


def exclude_period(period: str) -> dict[str, object]:
    normalized_period = normalize_period(period)
    config = load_config()
    included = set(config.get("included_periods", []))
    excluded = set(config.get("excluded_periods", []))
    included.discard(normalized_period)
    excluded.add(normalized_period)
    config["included_periods"] = sorted(included)
    config["excluded_periods"] = sorted(excluded)
    save_config(config)
    return config


def parse_monthly_file(path: Path) -> MonthlyFile | None:
    match = MONTHLY_FILE_PATTERN.match(path.name)
    if "pw-only" in path.name.lower():
        return None
    if not match:
        sell_through_match = SELL_THROUGH_FILE_PATTERN.match(path.name)
        if not sell_through_match:
            return None
        report_year = int(sell_through_match.group("year"))
        report_month = int(sell_through_match.group("month"))
        start_date = pd.Timestamp(year=report_year, month=report_month, day=1)
        end_date = start_date + pd.offsets.MonthEnd(0)
        return MonthlyFile(
            path=path,
            period=f"{report_year}{report_month:02d}",
            report_year=report_year,
            report_month=report_month,
            start_date=start_date.normalize(),
            end_date=pd.Timestamp(end_date).normalize(),
            is_updated=False,
        )

    start_date = pd.to_datetime(match.group("start"), format="%m-%d-%Y", errors="raise")
    end_date = pd.to_datetime(match.group("end"), format="%m-%d-%Y", errors="raise")
    report_month = int(match.group("month") or start_date.month)
    period = f"{start_date.year}{report_month:02d}"
    return MonthlyFile(
        path=path,
        period=period,
        report_year=start_date.year,
        report_month=report_month,
        start_date=pd.Timestamp(start_date).normalize(),
        end_date=pd.Timestamp(end_date).normalize(),
        is_updated=bool(match.group("updated")),
    )


def discover_year_folders(
    source_root: Path,
    year: int | None = None,
    all_years: bool = False,
    start_year: int | None = None,
    end_year: int | None = None,
    included_periods: list[str] | None = None,
) -> list[Path]:
    if not source_root.exists():
        raise FileNotFoundError(f"Amazon monthly customer orders root not found: {source_root}")

    if all_years or start_year is not None or end_year is not None:
        folders = [path for path in source_root.iterdir() if path.is_dir() and re.fullmatch(r"20\d{2}", path.name)]
        if start_year is not None:
            folders = [path for path in folders if int(path.name) >= start_year]
        if end_year is not None:
            folders = [path for path in folders if int(path.name) <= end_year]
        folder_by_year = {path.name: path for path in folders}
        for period in included_periods or []:
            folder = source_root / period[:4]
            if folder.exists() and folder.is_dir():
                folder_by_year[folder.name] = folder
        return sorted(folder_by_year.values(), key=lambda path: path.name)

    selected_year = year or datetime.now().year
    folder = source_root / str(selected_year)
    if not folder.exists():
        raise FileNotFoundError(f"Amazon monthly customer orders year folder not found: {folder}")
    if not folder.is_dir():
        raise NotADirectoryError(f"Amazon monthly customer orders year path is not a folder: {folder}")
    return [folder]


def discover_monthly_files(
    source_root: Path,
    year: int | None = None,
    all_years: bool = False,
    start_year: int | None = None,
    end_year: int | None = None,
    included_periods: list[str] | None = None,
    excluded_periods: list[str] | None = None,
) -> list[MonthlyFile]:
    files_by_period: dict[str, MonthlyFile] = {}
    for folder in discover_year_folders(
        source_root,
        year=year,
        all_years=all_years,
        start_year=start_year,
        end_year=end_year,
        included_periods=included_periods,
    ):
        for path in sorted(folder.iterdir(), key=lambda value: value.name.lower()):
            if not path.is_file() or path.suffix.lower() not in {".xlsx", ".csv"}:
                continue
            if path.name.startswith("~$"):
                continue
            parsed = parse_monthly_file(path)
            if parsed is not None:
                existing = files_by_period.get(parsed.period)
                if existing is None:
                    files_by_period[parsed.period] = parsed
                    continue
                parsed_priority = (parsed.is_updated, parsed.path.stat().st_mtime)
                existing_priority = (existing.is_updated, existing.path.stat().st_mtime)
                if parsed_priority > existing_priority:
                    files_by_period[parsed.period] = parsed
    included = set(included_periods or [])
    excluded = set(excluded_periods or [])
    if included:
        files_by_period = {
            period: monthly_file
            for period, monthly_file in files_by_period.items()
            if period in included
            or (
                (start_year is None or monthly_file.report_year >= start_year)
                and (end_year is None or monthly_file.report_year <= end_year)
            )
        }
    for period in excluded:
        files_by_period.pop(period, None)
    return sorted(files_by_period.values(), key=lambda item: (item.period, item.path.name.lower()))


def clean_columns(columns: pd.Index) -> list[str]:
    return [str(column).strip() for column in columns]


def detect_header_row(excel_file: pd.ExcelFile, sheet_name: str, max_rows: int = 12) -> int | None:
    preview = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=max_rows, dtype=object)
    important_lower = {column.lower() for column in IMPORTANT_COLUMNS}
    best_row = None
    best_score = 0
    for row_idx, row in preview.iterrows():
        values = {str(value).strip().lower() for value in row.dropna()}
        score = len(values & important_lower)
        if score > best_score:
            best_score = score
            best_row = int(row_idx)
    return best_row if best_score >= 2 else None


def read_detail_sheet(excel_file: pd.ExcelFile) -> tuple[pd.DataFrame, str]:
    best: tuple[int, pd.DataFrame, str] | None = None
    important_lower = {column.lower() for column in IMPORTANT_COLUMNS}
    for sheet_name in excel_file.sheet_names:
        header_row = detect_header_row(excel_file, sheet_name)
        if header_row is None:
            continue
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_row, dtype=object)
        df.columns = clean_columns(df.columns)
        score = len({column.lower() for column in df.columns} & important_lower)
        if best is None or score > best[0]:
            best = (score, df, sheet_name)
    if best is None:
        raise ValueError(f"Could not find a detail sheet with recognizable headers in {excel_file.io}")
    return best[1], best[2]


def detect_csv_header_row(file_path: Path, max_rows: int = 12) -> int:
    preview = pd.read_csv(file_path, header=None, nrows=max_rows, dtype=object, encoding="utf-8-sig")
    important_lower = {column.lower() for column in IMPORTANT_COLUMNS}
    best_row = 0
    best_score = 0
    for row_idx, row in preview.iterrows():
        values = {str(value).strip().lower() for value in row.dropna()}
        score = len(values & important_lower)
        if score > best_score:
            best_score = score
            best_row = int(row_idx)
    return best_row


def read_csv_detail(file_path: Path) -> pd.DataFrame:
    header_row = detect_csv_header_row(file_path)
    df = pd.read_csv(file_path, header=header_row, dtype=object, encoding="utf-8-sig")
    df.columns = clean_columns(df.columns)
    return df


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


def read_monthly_file(monthly_file: MonthlyFile) -> tuple[pd.DataFrame, list[str]]:
    if monthly_file.path.suffix.lower() == ".csv":
        df = read_csv_detail(monthly_file.path)
        sheet_name = ""
    else:
        excel_file = pd.ExcelFile(monthly_file.path)
        df, sheet_name = read_detail_sheet(excel_file)
    df.rename(columns=COLUMN_ALIASES, inplace=True)

    missing = [column for column in IMPORTANT_COLUMNS if column not in df.columns]
    df.insert(0, "Source Sheet", sheet_name)
    df.insert(0, "Source File", monthly_file.path.name)
    df.insert(0, "Report End Date", monthly_file.end_date)
    df.insert(0, "Report Start Date", monthly_file.start_date)
    df.insert(0, "Report Month", monthly_file.report_month)
    df.insert(0, "Report Year", monthly_file.report_year)
    df.insert(0, "Period", monthly_file.period)

    if "ISBN" in df.columns:
        df["ISBN"] = df["ISBN"].map(normalize_isbn)
    if "ASIN" in df.columns:
        df["ASIN"] = df["ASIN"].astype("string").str.strip()

    return df, missing


def standardize_for_parquet(df: pd.DataFrame) -> pd.DataFrame:
    standardized = df.copy()
    for column in standardized.columns:
        if column in DATE_LIKE_COLUMNS:
            standardized[column] = pd.to_datetime(standardized[column], errors="coerce")
            continue
        if column in TEXT_LIKE_COLUMNS:
            standardized[column] = standardized[column].astype("string").str.strip()
            continue
        if pd.api.types.is_object_dtype(standardized[column]):
            parsed = pd.to_numeric(standardized[column], errors="coerce")
            non_null = standardized[column].notna().sum()
            parsed_non_null = parsed.notna().sum()
            if non_null and parsed_non_null / non_null >= 0.8:
                standardized[column] = parsed.astype("Float64")
            else:
                standardized[column] = standardized[column].astype("string").str.strip()
    return standardized


def build_monthly_customer_orders(
    source_root: Path = process_paths.AMAZON_MONTHLY_CUSTOMER_ORDERS_ROOT,
    output_file: Path = process_paths.AMAZON_MONTHLY_CUSTOMER_ORDERS_PARQUET,
    year: int | None = None,
    all_years: bool = False,
    start_year: int | None = None,
    end_year: int | None = None,
    included_periods: list[str] | None = None,
    excluded_periods: list[str] | None = None,
) -> tuple[pd.DataFrame, dict[str, list[str]]]:
    monthly_files = discover_monthly_files(
        source_root,
        year=year,
        all_years=all_years,
        start_year=start_year,
        end_year=end_year,
        included_periods=included_periods,
        excluded_periods=excluded_periods,
    )
    if not monthly_files:
        scope = "all year folders" if all_years else str(year or datetime.now().year)
        raise FileNotFoundError(f"No Amazon monthly customer orders files found for {scope} under {source_root}")

    frames: list[pd.DataFrame] = []
    missing_by_file: dict[str, list[str]] = {}
    for monthly_file in monthly_files:
        df, missing = read_monthly_file(monthly_file)
        frames.append(df)
        if missing:
            missing_by_file[monthly_file.path.name] = missing

    combined = pd.concat(frames, ignore_index=True, sort=False)
    ordered_columns = [
        *METADATA_COLUMNS,
        *[column for column in IMPORTANT_COLUMNS if column in combined.columns],
        *[
            column
            for column in combined.columns
            if column not in METADATA_COLUMNS and column not in IMPORTANT_COLUMNS
        ],
    ]
    combined = combined.loc[:, ordered_columns]
    combined = standardize_for_parquet(combined)

    output_file.parent.mkdir(parents=True, exist_ok=True)
    combined.to_parquet(output_file, index=False, engine="pyarrow")
    return combined, missing_by_file


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Compile Amazon monthly customer orders workbooks to parquet.")
    parser.add_argument(
        "command",
        nargs="?",
        default="compile",
        choices=["compile", "status", "include", "exclude"],
        help="Compile parquet, show configured periods, include a period, or exclude a period.",
    )
    parser.add_argument("period", nargs="?", help="Period for include/exclude, in yyyymm format.")
    parser.add_argument("--year", type=int, help="Year folder to compile.")
    parser.add_argument("--all-years", action="store_true", help="Compile all numeric year folders under the source root.")
    parser.add_argument("--start-year", type=int, help="First year folder to include.")
    parser.add_argument("--end-year", type=int, help="Last year folder to include.")
    parser.add_argument("--source-root", type=Path, default=process_paths.AMAZON_MONTHLY_CUSTOMER_ORDERS_ROOT)
    parser.add_argument("--output", type=Path, default=process_paths.AMAZON_MONTHLY_CUSTOMER_ORDERS_PARQUET)
    return parser


def configured_period_args(args: argparse.Namespace) -> tuple[int | None, int | None, list[str], list[str]]:
    if args.year or args.all_years or args.start_year or args.end_year:
        return args.start_year, args.end_year, [], []

    config = load_config()
    start_year = config.get("start_year")
    end_year = config.get("end_year")
    return (
        start_year if isinstance(start_year, int) else None,
        end_year if isinstance(end_year, int) else None,
        list(config.get("included_periods", [])),
        list(config.get("excluded_periods", [])),
    )


def print_status(source_root: Path) -> None:
    config = load_config()
    start_year = config.get("start_year")
    end_year = config.get("end_year")
    included_periods = list(config.get("included_periods", []))
    excluded_periods = list(config.get("excluded_periods", []))
    files = discover_monthly_files(
        source_root,
        start_year=start_year if isinstance(start_year, int) else None,
        end_year=end_year if isinstance(end_year, int) else None,
        included_periods=included_periods,
        excluded_periods=excluded_periods,
    )
    print("Amazon Monthly Customer Orders")
    print(f"  Source root:       {source_root}")
    print(f"  Output parquet:    {process_paths.AMAZON_MONTHLY_CUSTOMER_ORDERS_PARQUET}")
    print(f"  Default start year:{start_year}")
    print(f"  Default end year:  {end_year or 'current/open'}")
    print(f"  Included periods:  {', '.join(included_periods) if included_periods else 'none'}")
    print(f"  Excluded periods:  {', '.join(excluded_periods) if excluded_periods else 'none'}")
    print(f"  Active periods:    {', '.join(file.period for file in files) if files else 'none'}")


def main() -> None:
    args = build_parser().parse_args()
    if args.command == "status":
        print_status(args.source_root)
        return
    if args.command == "include":
        if not args.period:
            raise ValueError("include requires a yyyymm period.")
        include_period(args.period)
        print(f"Included period: {normalize_period(args.period)}")
        print_status(args.source_root)
        return
    if args.command == "exclude":
        if not args.period:
            raise ValueError("exclude requires a yyyymm period.")
        exclude_period(args.period)
        print(f"Excluded period: {normalize_period(args.period)}")
        print_status(args.source_root)
        return

    start_year, end_year, included_periods, excluded_periods = configured_period_args(args)
    combined, missing_by_file = build_monthly_customer_orders(
        source_root=args.source_root,
        output_file=args.output,
        year=args.year,
        all_years=args.all_years,
        start_year=start_year,
        end_year=end_year,
        included_periods=included_periods,
        excluded_periods=excluded_periods,
    )

    print(f"Saved Amazon monthly customer orders parquet: {args.output}")
    print(f"Rows:    {len(combined):,}")
    print(f"Columns: {len(combined.columns):,}")
    print(f"Periods: {', '.join(sorted(combined['Period'].dropna().astype(str).unique()))}")
    if missing_by_file:
        print("Files missing important columns:")
        for filename, missing in missing_by_file.items():
            print(f"  - {filename}: {', '.join(missing)}")
    else:
        print("All important columns were present in every source file.")


if __name__ == "__main__":
    main()
