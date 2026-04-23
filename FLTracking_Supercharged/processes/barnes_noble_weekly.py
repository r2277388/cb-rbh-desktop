from pathlib import Path
import re
import sys
import warnings

import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.append(str(Path(__file__).resolve().parents[1]))
sys.path.append(str(REPO_ROOT / "bn_rolling_reports"))

from cache_utils import build_source_signature, load_cached_dataframe
from config import BARNES_NOBLE_DIR, BARNES_NOBLE_GLOB, OUTPUT_DIR
from isbn_utils import normalize_isbn_series
from rolling_paths import inventory_cache_file, sales_cache_file


warnings.filterwarnings(
    "ignore",
    message="Cannot parse header or footer so it will be ignored",
    category=UserWarning,
    module="openpyxl.worksheet.header_footer",
)


def resolve_barnes_noble_weekly_path(
    source_dir: Path = BARNES_NOBLE_DIR,
    filename_glob: str = BARNES_NOBLE_GLOB,
) -> Path:
    report_files = list(source_dir.glob(filename_glob))
    if not report_files:
        raise FileNotFoundError(
            f"No Barnes & Noble file matching '{filename_glob}' found in: {source_dir}"
        )
    return max(report_files, key=lambda path: path.stat().st_mtime)


def build_bn_date_header(source_path: Path) -> tuple[str, str]:
    match = re.search(r"\((\d{6})\)", source_path.name)
    if not match:
        raise ValueError(f"Could not parse date from Barnes & Noble filename: {source_path.name}")

    parsed = pd.to_datetime(match.group(1), format="%m%d%y")
    return f"BNUpdated_{parsed.strftime('%m_%d_%Y')}", parsed.strftime("%m/%d/%Y")


def build_bn_date_header_from_week(week: pd.Timestamp) -> tuple[str, str]:
    parsed = pd.Timestamp(week)
    return f"BNUpdated_{parsed.strftime('%m_%d_%Y')}", parsed.strftime("%m/%d/%Y")


def _legacy_cache_path(cache_path: Path) -> Path:
    return REPO_ROOT / "bn_rolling_reports" / "cache" / cache_path.name


def _resolve_existing_cache_path(cache_path: Path) -> Path:
    if cache_path.exists():
        return cache_path

    legacy_path = _legacy_cache_path(cache_path)
    if legacy_path.exists():
        return legacy_path

    raise FileNotFoundError(
        f"Barnes & Noble rolling cache not found: {cache_path}. "
        f"Legacy local cache also not found: {legacy_path}"
    )


def resolve_barnes_noble_parquet_paths() -> tuple[Path, Path]:
    return (
        _resolve_existing_cache_path(sales_cache_file),
        _resolve_existing_cache_path(inventory_cache_file),
    )


def _validate_barnes_noble_parquet_source(sales_path: Path, inventory_path: Path) -> pd.Timestamp:
    sales_df = pd.read_parquet(sales_path, columns=["Week", "ISBN", "qty"])
    inventory_df = pd.read_parquet(inventory_path, columns=["Week", "ISBN", "OH_Stores", "OH_DC"])
    if sales_df.empty:
        raise ValueError(f"Barnes & Noble sales cache is empty: {sales_path}")
    if inventory_df.empty:
        raise ValueError(f"Barnes & Noble inventory cache is empty: {inventory_path}")

    latest_sales_week = pd.to_datetime(sales_df["Week"], errors="coerce").max()
    if pd.isna(latest_sales_week):
        raise ValueError(f"B&N sales cache has no valid Week values: {sales_path}")
    return pd.Timestamp(latest_sales_week)


def _build_parquet_signature(sales_path: Path, inventory_path: Path) -> dict:
    return {
        "source": "bn_rolling_reports parquet",
        "sales": build_source_signature(sales_path),
        "inventory": build_source_signature(inventory_path),
    }


def _normalize_cached_isbns(series: pd.Series) -> pd.Series:
    def normalize_cached_value(value: object) -> str | None:
        if pd.isna(value):
            return None
        clean = re.sub(r"[-\s]", "", str(value).strip())
        if not clean or not clean.isdigit():
            return None
        if len(clean) in {12, 13}:
            return clean
        if len(clean) < 13:
            return clean.zfill(13)
        return None

    return series.map(normalize_cached_value).astype("object")


def load_barnes_noble_weekly_excel(source_path: Path | None = None) -> pd.DataFrame:
    resolved_path = source_path or resolve_barnes_noble_weekly_path()
    date_header, date_value = build_bn_date_header(resolved_path)

    df = pd.read_excel(
        resolved_path,
        header=None,
        skiprows=6,
        usecols=[4, 12, 14, 24],
        names=["ISBN", "B&N_OH_Store", "B&N_OH_DC", "B&N_LTD"],
        dtype={"ISBN": "object"},
        engine="openpyxl",
    )

    result = df.copy()
    result["ISBN"] = normalize_isbn_series(result["ISBN"])
    result = result.dropna(subset=["ISBN"])

    for col in ["B&N_OH_Store", "B&N_OH_DC", "B&N_LTD"]:
        result[col] = (
            pd.to_numeric(result[col], errors="coerce")
            .fillna(0)
            .round()
            .astype("Int64")
        )

    result[date_header] = date_value

    ordered_cols = ["ISBN", date_header, "B&N_OH_Store", "B&N_OH_DC", "B&N_LTD"]
    return result[ordered_cols]


def load_barnes_noble_weekly_from_parquet(
    sales_path: Path | None = None,
    inventory_path: Path | None = None,
) -> pd.DataFrame:
    resolved_sales_path, resolved_inventory_path = (
        (sales_path, inventory_path)
        if sales_path is not None and inventory_path is not None
        else resolve_barnes_noble_parquet_paths()
    )

    sales_df = pd.read_parquet(resolved_sales_path)
    inventory_df = pd.read_parquet(resolved_inventory_path)

    if sales_df.empty:
        raise ValueError(f"Barnes & Noble sales cache is empty: {resolved_sales_path}")
    if inventory_df.empty:
        raise ValueError(f"Barnes & Noble inventory cache is empty: {resolved_inventory_path}")

    required_sales_cols = {"Week", "ISBN", "qty"}
    required_inventory_cols = {"Week", "ISBN", "OH_Stores", "OH_DC"}
    missing_sales = sorted(required_sales_cols.difference(sales_df.columns))
    missing_inventory = sorted(required_inventory_cols.difference(inventory_df.columns))
    if missing_sales:
        raise ValueError(f"B&N sales cache is missing required columns: {missing_sales}")
    if missing_inventory:
        raise ValueError(f"B&N inventory cache is missing required columns: {missing_inventory}")

    sales = sales_df.loc[:, ["Week", "ISBN", "qty"]].copy()
    sales["Week"] = pd.to_datetime(sales["Week"], errors="coerce")
    sales["ISBN"] = _normalize_cached_isbns(sales["ISBN"].astype("string"))
    sales["qty"] = pd.to_numeric(sales["qty"], errors="coerce").fillna(0)
    sales = sales.dropna(subset=["Week", "ISBN"])
    latest_sales_week = sales["Week"].max()
    if pd.isna(latest_sales_week):
        raise ValueError(f"B&N sales cache has no valid Week values: {resolved_sales_path}")

    ltd = (
        sales.groupby("ISBN", as_index=False)["qty"]
        .sum()
        .rename(columns={"qty": "B&N_LTD"})
    )

    inventory = inventory_df.loc[:, ["Week", "ISBN", "OH_Stores", "OH_DC"]].copy()
    inventory["Week"] = pd.to_datetime(inventory["Week"], errors="coerce")
    inventory["ISBN"] = _normalize_cached_isbns(inventory["ISBN"].astype("string"))
    inventory = inventory.dropna(subset=["Week", "ISBN"])
    inventory_week = (
        latest_sales_week
        if (inventory["Week"] == latest_sales_week).any()
        else inventory["Week"].max()
    )
    latest_inventory = inventory[inventory["Week"] == inventory_week].copy()
    latest_inventory["OH_Stores"] = pd.to_numeric(latest_inventory["OH_Stores"], errors="coerce").fillna(0)
    latest_inventory["OH_DC"] = pd.to_numeric(latest_inventory["OH_DC"], errors="coerce").fillna(0)
    on_hand = (
        latest_inventory.groupby("ISBN", as_index=False)[["OH_Stores", "OH_DC"]]
        .sum()
        .rename(columns={"OH_Stores": "B&N_OH_Store", "OH_DC": "B&N_OH_DC"})
    )

    date_header, date_value = build_bn_date_header_from_week(latest_sales_week)
    result = ltd.merge(on_hand, on="ISBN", how="outer")
    result[date_header] = date_value

    for col in ["B&N_OH_Store", "B&N_OH_DC", "B&N_LTD"]:
        result[col] = (
            pd.to_numeric(result[col], errors="coerce")
            .fillna(0)
            .round()
            .astype("Int64")
        )

    ordered_cols = ["ISBN", date_header, "B&N_OH_Store", "B&N_OH_DC", "B&N_LTD"]
    return result[ordered_cols]


def load_barnes_noble_weekly(source_path: Path | None = None) -> pd.DataFrame:
    if source_path is not None and source_path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        return load_barnes_noble_weekly_excel(source_path)

    try:
        return load_barnes_noble_weekly_from_parquet()
    except Exception as parquet_error:
        print(f"Warning: B&N parquet cache unavailable; falling back to weekly Excel. {parquet_error}")
        return load_barnes_noble_weekly_excel()


def get_barnes_noble_parquet_metadata() -> dict[str, object]:
    sales_path, inventory_path = resolve_barnes_noble_parquet_paths()
    latest_sales_week = _validate_barnes_noble_parquet_source(sales_path, inventory_path)

    return {
        "sales_path": sales_path,
        "inventory_path": inventory_path,
        "report_date": pd.Timestamp(latest_sales_week).strftime("%m/%d/%Y"),
        "modified_date": datetime_from_latest_mtime(sales_path, inventory_path),
        "sales_modified_date": datetime_from_mtime(sales_path),
        "inventory_modified_date": datetime_from_mtime(inventory_path),
    }


def get_barnes_noble_excel_metadata() -> dict[str, object]:
    source_path = resolve_barnes_noble_weekly_path()
    _, report_date = build_bn_date_header(source_path)
    return {
        "source_type": "excel_fallback",
        "source_path": source_path,
        "report_date": report_date,
        "modified_date": datetime_from_mtime(source_path),
    }


def get_barnes_noble_source_metadata() -> dict[str, object]:
    try:
        metadata = get_barnes_noble_parquet_metadata()
        metadata["source_type"] = "parquet"
        return metadata
    except Exception as parquet_error:
        metadata = get_barnes_noble_excel_metadata()
        metadata["fallback_reason"] = str(parquet_error)
        return metadata


def datetime_from_latest_mtime(*paths: Path) -> str:
    latest_mtime = max(path.stat().st_mtime for path in paths)
    return pd.Timestamp.fromtimestamp(latest_mtime).strftime("%m/%d/%Y %I:%M %p")


def datetime_from_mtime(path: Path) -> str:
    return pd.Timestamp.fromtimestamp(path.stat().st_mtime).strftime("%m/%d/%Y %I:%M %p")


def load_barnes_noble_weekly_cached(
    source_path: Path | None = None,
) -> tuple[pd.DataFrame, bool, Path]:
    cache_path = OUTPUT_DIR / "barnes_noble_weekly_selected_columns.pkl"
    if source_path is not None and source_path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        signature = build_source_signature(source_path, extra={"source": "weekly excel"})
        build_func = lambda: load_barnes_noble_weekly_excel(source_path)
        df, cache_hit = load_cached_dataframe(
            cache_path,
            signature,
            build_func,
        )
        return df, cache_hit, cache_path

    try:
        sales_path, inventory_path = resolve_barnes_noble_parquet_paths()
        _validate_barnes_noble_parquet_source(sales_path, inventory_path)
        signature = _build_parquet_signature(sales_path, inventory_path)
        build_func = lambda: load_barnes_noble_weekly_from_parquet(sales_path, inventory_path)
        df, cache_hit = load_cached_dataframe(
            cache_path,
            signature,
            build_func,
        )
        return df, cache_hit, cache_path
    except Exception as parquet_error:
        print(f"Warning: B&N parquet cache unavailable; falling back to weekly Excel. {parquet_error}")
        try:
            excel_path = resolve_barnes_noble_weekly_path()
        except Exception:
            raise parquet_error
        signature = build_source_signature(excel_path, extra={"source": "weekly excel fallback"})
        df, cache_hit = load_cached_dataframe(
            cache_path,
            signature,
            lambda: load_barnes_noble_weekly_excel(excel_path),
        )
        return df, cache_hit, cache_path


def save_barnes_noble_weekly_output(df: pd.DataFrame) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / "barnes_noble_weekly_selected_columns.pkl"
    df.to_pickle(output_path)
    return output_path


def main() -> None:
    source_metadata = get_barnes_noble_source_metadata()
    df, cache_hit, output_path = load_barnes_noble_weekly_cached()

    if source_metadata["source_type"] == "parquet":
        print(f"Loaded sales cache: {source_metadata['sales_path']}")
        print(f"Loaded inventory cache: {source_metadata['inventory_path']}")
    else:
        print(f"Loaded weekly Excel fallback: {source_metadata['source_path']}")
    print(f"Used cached pickle: {cache_hit}")
    print(f"Rows after removing null ISBN values: {len(df)}")
    print(df.head(20).to_string(index=False))
    print(f"\nSaved output to: {output_path}")


if __name__ == "__main__":
    main()
