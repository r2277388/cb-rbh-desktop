import pandas as pd
from pathlib import Path

from paths import ATELIER_AMAZON_INVENTORY_FOLDER

INVENTORY_NA_VALUES = ["â€”", "—", ""]
NUMERIC_SUMMARY_EXCLUSIONS = {"ASIN", "ISBN", "EAN", "UPC", "Model Number"}


def get_latest_inventory_file(folder_path: Path, pattern: str) -> Path:
    """Return the latest inventory file in the folder matching the given pattern."""
    files = list(folder_path.glob(pattern))

    if not files:
        raise FileNotFoundError(f"No files found in {folder_path} with pattern {pattern}")

    return max(files, key=lambda path: path.stat().st_mtime)


def _normalize_numeric_series(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.strip()
        .replace(INVENTORY_NA_VALUES, pd.NA)
        .str.replace(",", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def print_inventory_numeric_summary(file_path: Path) -> None:
    """Print a compact summary for numeric-looking columns in the raw inventory CSV."""
    df_raw = pd.read_csv(
        file_path,
        skiprows=1,
        na_values=INVENTORY_NA_VALUES,
        low_memory=False,
        dtype=str,
    )

    print()
    print("Inventory file summary:")
    print(f"  Source file: {file_path}")
    print(f"  Rows:        {len(df_raw):,}")

    numeric_summaries: list[tuple[str, int, float, float, float]] = []
    for column in df_raw.columns:
        if column in NUMERIC_SUMMARY_EXCLUSIONS:
            continue

        numeric_values = _normalize_numeric_series(df_raw[column])
        non_null = int(numeric_values.notna().sum())
        if non_null == 0:
            continue

        numeric_summaries.append(
            (
                column,
                non_null,
                float(numeric_values.sum()),
                float(numeric_values.min()),
                float(numeric_values.max()),
            )
        )

    if not numeric_summaries:
        print("  Numeric columns with data: none found")
        print()
        return

    print("  Numeric columns with data:")
    for column, non_null, total, min_value, max_value in numeric_summaries:
        print(
            f"    {column}: nonblank={non_null:,}, "
            f"sum={total:,.0f}, min={min_value:,.0f}, max={max_value:,.0f}"
        )
    print()


def read_inventory_file(file_path: Path, columns: list[str]) -> pd.DataFrame:
    """Read the inventory CSV file and return a DataFrame."""
    return pd.read_csv(
        file_path,
        skiprows=1,
        na_values=INVENTORY_NA_VALUES,
        usecols=columns,
        low_memory=False,
        dtype={
            "ASIN": object,
            "Inventory Qty": int,
        },
    )


def rename_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename specific columns in the DataFrame."""
    return df.rename(
        columns={
            "ASIN": "ASIN",
            "Unfilled Customer Ordered Units": "Orders",
        }
    )


def process_inventory_data(folder_path: str, pattern: str, columns: list[str]) -> pd.DataFrame:
    """Load and process the latest inventory export."""
    folder = Path(folder_path)
    file_inventory = get_latest_inventory_file(folder, pattern)
    print_inventory_numeric_summary(file_inventory)
    df_inventory = read_inventory_file(file_inventory, columns)
    return rename_columns(df_inventory)


def data_inventory() -> pd.DataFrame:
    folder_path = ATELIER_AMAZON_INVENTORY_FOLDER
    pattern = "*Inventory*csv"
    columns = ["ASIN", "Unfilled Customer Ordered Units"]

    df = process_inventory_data(str(folder_path), pattern, columns)
    df["Orders"] = df["Orders"].replace(",", "", regex=True).fillna(0).astype(int)
    df = df[df["Orders"] > 0]

    return df


def main() -> None:
    df = data_inventory()
    print(df.info())
    print(df.head())


if __name__ == "__main__":
    main()
