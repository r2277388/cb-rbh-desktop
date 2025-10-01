import pandas as pd
import numpy as np
import time

from load_ypticod import load_ypticod
from combine_files import combine_weekly_files
from load_ebs_isbn_key import isbn_key

from asin_removal_list import asins_to_delete_list

def asin_isbn_conversion():
    df_combined = combine_weekly_files()
    df_ypticod = load_ypticod()
    df_isbn = isbn_key()
    
    isbn_set = set(
        df_isbn['ISBN']
        .dropna()
        .unique().tolist()
        )
    isbn_set.discard("ISBN")  # Remove the header if present

    # Merge combined weekly data with ypticod data on ASIN
    df_merged = df_combined.merge(df_ypticod, on='ASIN', how='left')

    # Start with ISBN from ypticod
    isbn_col = df_merged['ISBN']

    # Fill missing ISBNs with EAN if in isbn_set
    mask = isbn_col.isna() & df_merged['EAN'].isin(isbn_set)
    isbn_col = np.where(mask, df_merged['EAN'], isbn_col)

    # Fill remaining missing ISBNs with ISBN_Amz if in isbn_set
    mask = pd.isna(isbn_col) & df_merged['ISBN_Amz'].isin(isbn_set)
    isbn_col = np.where(mask, df_merged['ISBN_Amz'], isbn_col)

    # Fill remaining missing ISBNs with Model Number if in isbn_set
    mask = pd.isna(isbn_col) & df_merged['Model Number'].isin(isbn_set)
    isbn_col = np.where(mask, df_merged['Model Number'], isbn_col)

    # Fill any remaining missing ISBNs with "NO_ISBN"
    isbn_col = np.where(pd.isna(isbn_col), "NO_ISBN", isbn_col)

    df_merged['ISBN'] = isbn_col
    
    # Remove rows with ASINs in asins_to_delete_list
    df_merged = df_merged[~df_merged['ASIN'].isin(asins_to_delete_list)]

    # Remove rows where Product Title ends with 'anglais'
    df_merged = df_merged[~df_merged['Product Title'].fillna('').str.lower().str.endswith('anglais')]

    float_cols = df_merged.select_dtypes(include='float64').columns
    df_merged[float_cols] = df_merged[float_cols].fillna(0)
    df_merged = df_merged[~(df_merged[float_cols].eq(0).all(axis=1))]

    return df_merged

def main():
    start_time = time.time()  # Start timer
    df = asin_isbn_conversion()

    print(df.info())
    print(df.head(20))
    
    # Show count and top 20 rows where ISBN is "NO_ISBN"
    no_isbn_count = (df['ISBN'] == "NO_ISBN").sum()
    print(f'Rows with NO_ISBN: {no_isbn_count}')
    cols_to_show = ['ASIN', 'Product Title', 'EAN', 'ISBN_Amz', 'Model Number', 'ISBN']
    print(df[df['ISBN'] == "NO_ISBN"][cols_to_show].head(20))
    
    end_time = time.time()  # End timer
    elapsed = end_time - start_time
    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)
    print(f"Total runtime: {minutes} minutes, {seconds} seconds.")
    
if __name__ == "__main__":
    main()