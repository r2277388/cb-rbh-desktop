import pandas as pd
import numpy as np

from loader.loader_item import upload_item
from loader.loader_catalog import data_catalog

def asin_isbn_conversion() -> pd.DataFrame:
    df_item = upload_item()
    df_catalog = data_catalog()    
    
    """Convert ASIN to ISBN using a prioritized check of ISBN-13, EAN, and Model Number."""
    isbn_list = df_item['ISBN'].unique()
    
    # Start with ISBN-13 and use combine_first to fill with EAN, then Model Number
    df_catalog['ISBN'] = (
        df_catalog['ISBN-13'].where(df_catalog['ISBN-13'].isin(isbn_list))
        .combine_first(df_catalog['EAN'].where(df_catalog['EAN'].isin(isbn_list)))
        .combine_first(df_catalog['Model Number'].where(df_catalog['Model Number'].isin(isbn_list)))
    )
    
    # Drop rows where ISBN is NaN
    df_catalog = df_catalog.dropna(subset=['ISBN'])

    # Create the ASIN to ISBN DataFrame
    df = df_catalog[['ASIN', 'ISBN', 'Release Date']].reset_index(drop=True)

    return df

def main():
    df = asin_isbn_conversion()
    print(df.info())
    print(df.head())

if __name__ == '__main__':
    main()