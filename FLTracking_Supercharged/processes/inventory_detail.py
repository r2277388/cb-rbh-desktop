from pathlib import Path

import pandas as pd

from cache_utils import build_source_signature, load_cached_dataframe
from config import INVENTORY_DETAIL_DIR, INVENTORY_DETAIL_GLOB, OUTPUT_DIR
from isbn_utils import normalize_isbn_series


REQUIRED_COLUMNS = [
    "ISBN",
    "Reprint Due Date",
    "Reprint Quantity",
    "Available To Sell",
    "Frozen",
    "Reprint Freeze",
]

def resolve_inventory_detail_path(
    source_dir: Path = INVENTORY_DETAIL_DIR,
    filename_glob: str = INVENTORY_DETAIL_GLOB,
) -> Path:
    inventory_files = list(source_dir.glob(filename_glob))
    if not inventory_files:
        raise FileNotFoundError(
            f"No Inventory Detail file matching '{filename_glob}' found in: {source_dir}"
        )
    return inventory_files[0]


def load_inventory_detail(source_path: Path | None = None) -> pd.DataFrame:
    resolved_path = source_path or resolve_inventory_detail_path()

    df = pd.read_excel(
        resolved_path,
        usecols=REQUIRED_COLUMNS,
        dtype={"ISBN": "object"},
        engine="openpyxl",
    )

    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_columns:
        raise ValueError(
            "Inventory Detail is missing required columns: "
            + ", ".join(missing_columns)
        )

    result = df[REQUIRED_COLUMNS].copy()
    result = result.dropna(subset=["ISBN"])
    result["ISBN"] = normalize_isbn_series(result["ISBN"])
    result = result.dropna(subset=["ISBN"])

    return result


def load_inventory_detail_cached(source_path: Path | None = None) -> tuple[pd.DataFrame, bool, Path]:
    resolved_path = source_path or resolve_inventory_detail_path()
    cache_path = OUTPUT_DIR / "inventory_detail_selected_columns.pkl"
    signature = build_source_signature(resolved_path)
    df, cache_hit = load_cached_dataframe(
        cache_path,
        signature,
        lambda: load_inventory_detail(resolved_path),
    )
    return df, cache_hit, cache_path


def save_inventory_detail_output(df: pd.DataFrame) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / "inventory_detail_selected_columns.pkl"
    df.to_pickle(output_path)
    return output_path


def main() -> None:
    source_path = resolve_inventory_detail_path()
    df, cache_hit, output_path = load_inventory_detail_cached(source_path)

    print(f"Loaded source: {source_path}")
    print(f"Used cached pickle: {cache_hit}")
    print(f"Rows after removing null ISBN values: {len(df)}")
    print(df.head(20).to_string(index=False))
    print(f"\nSaved output to: {output_path}")


if __name__ == "__main__":
    main()
