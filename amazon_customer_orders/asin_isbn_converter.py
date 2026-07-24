import pandas as pd
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from loader.loader_item import upload_item
from loader.loader_catalog import data_catalog
from loader.loader_ypticod import get_cleaned_ypticod
from shared.amazon_metadata import resolve_isbn_series

def asin_isbn_conversion() -> pd.DataFrame:
    # This top part of the code provides the ASIN, ISBN and Release Date from the
    # Amazon catalog report.
    df_item = upload_item()
    df_catalog = data_catalog()

    """Convert ASIN to ISBN using a prioritized check of ISBN-13, EAN, and Model Number."""
    df_catalog["ISBN"] = resolve_isbn_series(
        df_catalog,
        df_item,
        ["ISBN-13", "EAN", "Model Number"],
    )
    df = df_catalog.dropna(subset=["ISBN"])[
        ["ASIN", "ISBN", "Release Date"]
    ].reset_index(drop=True)

    # This gives us ASIN and ISBN from the ypticod table as well as the OSD from a SQL table.
    df_ypticod = get_cleaned_ypticod()
    df_ypticod = df_ypticod[['ASIN', 'ISBN', 'Release Date']]
    df_combined = pd.concat([df_ypticod,df], ignore_index=True)
    df_combined = df_combined.drop_duplicates(subset=['ASIN','ISBN'], keep='first')

    return df_combined

def main():
    df = asin_isbn_conversion()
    print(df.info())
    print(df.head())

if __name__ == '__main__':
    main()