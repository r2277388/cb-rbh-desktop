from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from cache_utils import build_source_signature, load_cached_dataframe
from config import BOOKSHOP_PREORDERS_DIR, BOOKSHOP_PREORDERS_GLOB, OUTPUT_DIR
from isbn_utils import normalize_isbn_series


def resolve_bookshop_preorders_path(
    source_dir: Path = BOOKSHOP_PREORDERS_DIR,
    filename_glob: str = BOOKSHOP_PREORDERS_GLOB,
) -> Path:
    report_files = [path for path in source_dir.glob(filename_glob) if not path.name.startswith("~$")]
    if not report_files:
        raise FileNotFoundError(
            f"No Bookshop preorders file matching '{filename_glob}' found in: {source_dir}"
        )
    return max(report_files, key=lambda path: path.stat().st_mtime)


def _extract_year_from_filename(source_path: Path) -> int | None:
    match = re.search(r"(\d{2})\.(\d{2})\.(\d{2})", source_path.name)
    if not match:
        return None
    return 2000 + int(match.group(3))


def parse_bookshop_report_date(source_path: Path) -> str:
    header_df = pd.read_csv(source_path, nrows=0)
    columns = list(header_df.columns)
    if len(columns) < 4:
        return ""

    raw_header = str(columns[3]).strip()
    match = re.search(r"(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?$", raw_header)
    if not match:
        return ""

    month = int(match.group(1))
    day = int(match.group(2))
    year_text = match.group(3)
    if year_text:
        year = int(year_text)
        if year < 100:
            year += 2000
    else:
        year = _extract_year_from_filename(source_path) or pd.Timestamp.now().year

    return f"{month:02d}/{day:02d}/{year:04d}"


def load_bookshop_preorders(source_path: Path | None = None) -> pd.DataFrame:
    resolved_path = source_path or resolve_bookshop_preorders_path()
    df = pd.read_csv(
        resolved_path,
        usecols=[0, 3],
        dtype={0: "object"},
    )

    result = df.copy()
    result.columns = ["ISBN", "BookshopPreOrders"]
    result = result.dropna(subset=["ISBN"])
    result["ISBN"] = normalize_isbn_series(result["ISBN"])
    result = result.dropna(subset=["ISBN"])
    result["BookshopPreOrders"] = (
        pd.to_numeric(result["BookshopPreOrders"], errors="coerce")
        .fillna(0)
        .round()
        .astype("Int64")
    )
    return result


def load_bookshop_preorders_cached(
    source_path: Path | None = None,
) -> tuple[pd.DataFrame, bool, Path]:
    resolved_path = source_path or resolve_bookshop_preorders_path()
    cache_path = OUTPUT_DIR / "bookshop_preorders_selected_columns.pkl"
    signature = build_source_signature(resolved_path)
    df, cache_hit = load_cached_dataframe(
        cache_path,
        signature,
        lambda: load_bookshop_preorders(resolved_path),
    )
    return df, cache_hit, cache_path


def save_bookshop_preorders_output(df: pd.DataFrame) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / "bookshop_preorders_selected_columns.pkl"
    df.to_pickle(output_path)
    return output_path


def main() -> None:
    source_path = resolve_bookshop_preorders_path()
    df, cache_hit, output_path = load_bookshop_preorders_cached(source_path)
    report_date = parse_bookshop_report_date(source_path)

    print(f"Loaded source: {source_path}")
    if report_date:
        print(f"Report date: {report_date}")
    print(f"Used cached pickle: {cache_hit}")
    print(f"Rows after removing null ISBN values: {len(df)}")
    print(df.head(20).to_string(index=False))
    print(f"\nSaved output to: {output_path}")


if __name__ == "__main__":
    main()
