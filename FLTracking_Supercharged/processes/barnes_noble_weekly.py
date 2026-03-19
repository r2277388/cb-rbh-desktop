from pathlib import Path
import re
import sys
import warnings

import pandas as pd


sys.path.append(str(Path(__file__).resolve().parents[1]))

from cache_utils import build_source_signature, load_cached_dataframe
from config import BARNES_NOBLE_DIR, BARNES_NOBLE_GLOB, OUTPUT_DIR
from isbn_utils import normalize_isbn_series


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


def load_barnes_noble_weekly(source_path: Path | None = None) -> pd.DataFrame:
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


def load_barnes_noble_weekly_cached(
    source_path: Path | None = None,
) -> tuple[pd.DataFrame, bool, Path]:
    resolved_path = source_path or resolve_barnes_noble_weekly_path()
    cache_path = OUTPUT_DIR / "barnes_noble_weekly_selected_columns.pkl"
    signature = build_source_signature(resolved_path)
    df, cache_hit = load_cached_dataframe(
        cache_path,
        signature,
        lambda: load_barnes_noble_weekly(resolved_path),
    )
    return df, cache_hit, cache_path


def save_barnes_noble_weekly_output(df: pd.DataFrame) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / "barnes_noble_weekly_selected_columns.pkl"
    df.to_pickle(output_path)
    return output_path


def main() -> None:
    source_path = resolve_barnes_noble_weekly_path()
    df, cache_hit, output_path = load_barnes_noble_weekly_cached(source_path)

    print(f"Loaded source: {source_path}")
    print(f"Used cached pickle: {cache_hit}")
    print(f"Rows after removing null ISBN values: {len(df)}")
    print(df.head(20).to_string(index=False))
    print(f"\nSaved output to: {output_path}")


if __name__ == "__main__":
    main()
