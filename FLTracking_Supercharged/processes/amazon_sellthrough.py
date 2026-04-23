import sys
from pathlib import Path
import re

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.append(str(REPO_ROOT))
sys.path.append(str(Path(__file__).resolve().parents[1]))

from config import OUTPUT_DIR, SQL_DIR
from isbn_utils import normalize_isbn_series
from amazon_rolling_reports.paths import units_shipped_pickle_file
from shared.db.connection import get_connection
from shared.db.query_runner import fetch_data_from_db


SQL_FILE = SQL_DIR / "amazon_sellthrough_latest.sql"


def _legacy_cache_path(cache_path: Path) -> Path:
    return REPO_ROOT / "amazon_rolling_reports" / cache_path.name


def _latest_local_backup(cache_path: Path) -> Path | None:
    backup_dir = REPO_ROOT / "amazon_rolling_reports" / "backups"
    backups = sorted(
        backup_dir.glob(f"{cache_path.stem}_*{cache_path.suffix}"),
        key=lambda path: path.stat().st_mtime,
    )
    return backups[-1] if backups else None


def resolve_amazon_units_shipped_cache_path() -> Path:
    if units_shipped_pickle_file.exists():
        return units_shipped_pickle_file

    legacy_path = _legacy_cache_path(units_shipped_pickle_file)
    if legacy_path.exists():
        return legacy_path

    backup_path = _latest_local_backup(units_shipped_pickle_file)
    if backup_path is not None:
        return backup_path

    raise FileNotFoundError(
        f"Amazon rolling units shipped cache not found: {units_shipped_pickle_file}. "
        f"Legacy local cache also not found: {legacy_path}"
    )


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


def _parse_history_column(column: object) -> pd.Timestamp | None:
    if not isinstance(column, str):
        return None
    parsed = pd.to_datetime(column, format="%m-%d-%Y", errors="coerce")
    if pd.isna(parsed):
        return None
    return pd.Timestamp(parsed)


def _latest_week_from_columns(columns: pd.Index) -> pd.Timestamp:
    weeks = [week for week in (_parse_history_column(column) for column in columns) if week is not None]
    if not weeks:
        raise ValueError("Amazon rolling units shipped cache has no weekly history columns.")
    return max(weeks)


def _load_amazon_sellthrough_cache(source_path: Path | None = None) -> pd.DataFrame:
    resolved_path = source_path or resolve_amazon_units_shipped_cache_path()
    df = pd.read_pickle(resolved_path)

    required_columns = ["ISBN", "LTD", "OH"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(
            "Amazon rolling units shipped cache is missing required columns: "
            + ", ".join(missing_columns)
        )

    latest_week = _latest_week_from_columns(df.columns)
    result = df[required_columns].rename(
        columns={
            "LTD": "AmzUnitShipped_LTD",
            "OH": "AmzOnHand",
        }
    ).copy()
    result["AmzLastWeek"] = latest_week
    result = result.dropna(subset=["ISBN"])
    result["ISBN"] = _normalize_cached_isbns(result["ISBN"])
    result = result.dropna(subset=["ISBN"])
    result["AmzUnitShipped_LTD"] = (
        pd.to_numeric(result["AmzUnitShipped_LTD"], errors="coerce")
        .fillna(0)
        .round()
        .astype("Int64")
    )
    result["AmzOnHand"] = (
        pd.to_numeric(result["AmzOnHand"], errors="coerce")
        .fillna(0)
        .round()
        .astype("Int64")
    )

    return result[["ISBN", "AmzUnitShipped_LTD", "AmzLastWeek", "AmzOnHand"]]


def load_amazon_sellthrough_sql(sql_path: Path = SQL_FILE) -> pd.DataFrame:
    query = sql_path.read_text(encoding="utf-8").strip()
    engine = get_connection()
    df = fetch_data_from_db(engine, query)

    required_columns = ["ISBN", "AmzUnitShipped_LTD", "AmzLastWeek", "AmzOnHand"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(
            "Amazon sellthrough query is missing required columns: "
            + ", ".join(missing_columns)
        )

    result = df[required_columns].copy()
    result = result.dropna(subset=["ISBN"])
    result["ISBN"] = normalize_isbn_series(result["ISBN"])
    result = result.dropna(subset=["ISBN"])

    return result


def load_amazon_sellthrough(source_path: Path | None = None) -> pd.DataFrame:
    try:
        return _load_amazon_sellthrough_cache(source_path)
    except Exception as cache_error:
        print(f"Warning: Amazon rolling cache unavailable; falling back to SQL. {cache_error}")
        return load_amazon_sellthrough_sql()


def get_amazon_sellthrough_cache_metadata() -> dict[str, object]:
    cache_path = resolve_amazon_units_shipped_cache_path()
    df = pd.read_pickle(cache_path)
    required_columns = ["ISBN", "LTD", "OH"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(
            "Amazon rolling units shipped cache is missing required columns: "
            + ", ".join(missing_columns)
        )
    latest_week = _latest_week_from_columns(df.columns)
    return {
        "cache_path": cache_path,
        "report_date": latest_week.strftime("%m/%d/%Y"),
        "modified_date": pd.Timestamp.fromtimestamp(cache_path.stat().st_mtime).strftime("%m/%d/%Y %I:%M %p"),
    }


def get_amazon_sellthrough_source_metadata() -> dict[str, object]:
    try:
        metadata = get_amazon_sellthrough_cache_metadata()
        metadata["source_type"] = "cache"
        return metadata
    except Exception as cache_error:
        return {
            "source_type": "sql_fallback",
            "sql_path": SQL_FILE,
            "report_date": "",
            "modified_date": pd.Timestamp.fromtimestamp(SQL_FILE.stat().st_mtime).strftime("%m/%d/%Y %I:%M %p"),
            "fallback_reason": str(cache_error),
        }


def save_amazon_sellthrough_output(df: pd.DataFrame) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / "amazon_sellthrough_latest.pkl"
    df.to_pickle(output_path)
    return output_path


def main() -> None:
    source_metadata = get_amazon_sellthrough_source_metadata()
    df = load_amazon_sellthrough()
    output_path = save_amazon_sellthrough_output(df)

    if source_metadata["source_type"] == "cache":
        print(f"Loaded Amazon rolling units shipped cache: {source_metadata['cache_path']}")
    else:
        print(f"Loaded Amazon sellthrough SQL fallback: {source_metadata['sql_path']}")
    print(f"Rows after removing null ISBN values: {len(df)}")
    print(df.head(20).to_string(index=False))
    print(f"\nSaved output to: {output_path}")


if __name__ == "__main__":
    main()
