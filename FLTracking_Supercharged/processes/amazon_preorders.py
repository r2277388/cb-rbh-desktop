from pathlib import Path

import pandas as pd

from cache_utils import build_source_signature, load_cached_dataframe
from config import AMAZON_PREORDERS_PATH, AMAZON_PREORDERS_SHEET, OUTPUT_DIR
from isbn_utils import normalize_isbn_series


REQUIRED_COLUMNS = ["ISBN", "Orders"]
RENAMED_COLUMNS = {"Orders": "AmzPreOrders"}


def load_amazon_preorders(
    source_path: Path = AMAZON_PREORDERS_PATH,
    sheet_name: str = AMAZON_PREORDERS_SHEET,
) -> pd.DataFrame:
    if not source_path.exists():
        raise FileNotFoundError(f"Amazon preorders file not found: {source_path}")

    df = pd.read_excel(
        source_path,
        sheet_name=sheet_name,
        usecols=REQUIRED_COLUMNS,
        dtype={"ISBN": "object"},
        engine="openpyxl",
    )

    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_columns:
        raise ValueError(
            "Amazon preorders is missing required columns: "
            + ", ".join(missing_columns)
        )

    result = df[REQUIRED_COLUMNS].copy()
    result = result.dropna(subset=["ISBN"])
    result["ISBN"] = normalize_isbn_series(result["ISBN"])
    result = result.dropna(subset=["ISBN"])
    result = result.rename(columns=RENAMED_COLUMNS)

    return result


def load_amazon_preorders_cached(
    source_path: Path = AMAZON_PREORDERS_PATH,
    sheet_name: str = AMAZON_PREORDERS_SHEET,
) -> tuple[pd.DataFrame, bool, Path]:
    cache_path = OUTPUT_DIR / "amazon_preorders_selected_columns.pkl"
    signature = build_source_signature(source_path, {"sheet_name": sheet_name})
    df, cache_hit = load_cached_dataframe(
        cache_path,
        signature,
        lambda: load_amazon_preorders(source_path=source_path, sheet_name=sheet_name),
    )
    return df, cache_hit, cache_path


def save_amazon_preorders_output(df: pd.DataFrame) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / "amazon_preorders_selected_columns.pkl"
    df.to_pickle(output_path)
    return output_path


def main() -> None:
    df, cache_hit, output_path = load_amazon_preorders_cached()

    print(f"Loaded source: {AMAZON_PREORDERS_PATH} [{AMAZON_PREORDERS_SHEET}]")
    print(f"Used cached pickle: {cache_hit}")
    print(f"Rows after removing null ISBN values: {len(df)}")
    print(df.head(20).to_string(index=False))
    print(f"\nSaved output to: {output_path}")


if __name__ == "__main__":
    main()
