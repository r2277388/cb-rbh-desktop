from pathlib import Path
import sys
import warnings

import pandas as pd


sys.path.append(str(Path(__file__).resolve().parents[1]))

from cache_utils import build_source_signature, load_cached_dataframe
from config import INGRAM_REPORT_DIR, INGRAM_REPORT_GLOB, OUTPUT_DIR
from isbn_utils import normalize_isbn_series

warnings.filterwarnings(
    "ignore",
    message="Cannot parse header or footer so it will be ignored",
    category=UserWarning,
    module="openpyxl.worksheet.header_footer",
)


REQUIRED_COLUMNS = [
    "EAN",
    "Avg 4wk Sales",
    "Ingram On Hand",
    "Total Ingram's Customer BO",
]

RENAMED_COLUMNS = {
    "EAN": "ISBN",
    "Ingram On Hand": "IngramOH",
    "Total Ingram's Customer BO": "IngramPreOrders",
}


def resolve_ingram_daily_report_path(
    source_dir: Path = INGRAM_REPORT_DIR,
    filename_glob: str = INGRAM_REPORT_GLOB,
) -> Path:
    report_files = list(source_dir.glob(filename_glob))
    if not report_files:
        raise FileNotFoundError(
            f"No Ingram Daily Report file matching '{filename_glob}' found in: {source_dir}"
    )
    return max(report_files, key=lambda path: path.stat().st_mtime)


def build_modified_date_header(source_path: Path) -> tuple[str, str]:
    modified_at = pd.Timestamp(source_path.stat().st_mtime, unit="s")
    display_date = modified_at.strftime("%m/%d/%Y")
    header = f"IngramUpdated_{modified_at.strftime('%m_%d_%Y')}"
    return header, display_date


def load_ingram_daily_report(source_path: Path | None = None) -> pd.DataFrame:
    resolved_path = source_path or resolve_ingram_daily_report_path()
    modified_header, modified_value = build_modified_date_header(resolved_path)

    df = pd.read_excel(
        resolved_path,
        header=2,
        usecols=REQUIRED_COLUMNS,
        dtype={"EAN": "object"},
        engine="openpyxl",
    )

    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_columns:
        raise ValueError(
            "Ingram Daily Report is missing required columns: "
            + ", ".join(missing_columns)
        )

    result = df[REQUIRED_COLUMNS].copy()
    result["EAN"] = normalize_isbn_series(result["EAN"])
    result = result.dropna(subset=["EAN"])
    result = result.rename(columns=RENAMED_COLUMNS)
    result["Ingram4WkSales"] = (
        pd.to_numeric(result["Avg 4wk Sales"], errors="coerce").fillna(0).mul(4).round().astype("Int64")
    )
    result["IngramOH"] = pd.to_numeric(result["IngramOH"], errors="coerce").fillna(0).astype("Int64")
    result["IngramPreOrders"] = (
        pd.to_numeric(result["IngramPreOrders"], errors="coerce").fillna(0).astype("Int64")
    )
    result[modified_header] = modified_value

    return result[["ISBN", modified_header, "Ingram4WkSales", "IngramOH", "IngramPreOrders"]]


def load_ingram_daily_report_cached(
    source_path: Path | None = None,
) -> tuple[pd.DataFrame, bool, Path]:
    resolved_path = source_path or resolve_ingram_daily_report_path()
    cache_path = OUTPUT_DIR / "ingram_daily_report_selected_columns.pkl"
    signature = build_source_signature(resolved_path)
    df, cache_hit = load_cached_dataframe(
        cache_path,
        signature,
        lambda: load_ingram_daily_report(resolved_path),
    )
    return df, cache_hit, cache_path


def save_ingram_daily_report_output(df: pd.DataFrame) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / "ingram_daily_report_selected_columns.pkl"
    df.to_pickle(output_path)
    return output_path


def main() -> None:
    source_path = resolve_ingram_daily_report_path()
    df, cache_hit, output_path = load_ingram_daily_report_cached(source_path)

    print(f"Loaded source: {source_path}")
    print(f"Used cached pickle: {cache_hit}")
    print(f"Rows after removing null ISBN values: {len(df)}")
    print(df.head(20).to_string(index=False))
    print(f"\nSaved output to: {output_path}")


if __name__ == "__main__":
    main()
