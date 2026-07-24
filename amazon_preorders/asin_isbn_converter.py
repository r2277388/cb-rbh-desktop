import pandas as pd
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from loader_ebs_item import data_item
from loader_catalog import data_catalog
from shared.amazon_metadata import resolve_isbn_series

def asin_isbn_conversion() -> pd.DataFrame:
    df_item = data_item()
    df_catalog = data_catalog()

    """Convert ASIN to ISBN using a prioritized check of ISBN-13, EAN, and Model Number."""
    df_catalog["ISBN"] = resolve_isbn_series(
        df_catalog,
        df_item,
        ["ISBN-13", "EAN", "Model Number"],
    )
    return df_catalog.dropna(subset=["ISBN"])[
        ["ASIN", "ISBN", "Release Date"]
    ].reset_index(drop=True)

def main():
    df = asin_isbn_conversion()
    print(df.info())
    print(df.head())

if __name__ == '__main__':
    main()