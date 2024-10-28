import pandas as pd
import numpy as np
from pathlib import Path
import os
import glob

def load_sales_data(folder_path: Path, file_glob_pattern: str, columns: list) -> pd.DataFrame:
    files = glob.glob(str(folder_path) + file_glob_pattern)
    most_recent_file = max(files, key=os.path.getctime)
    df = pd.read_csv(most_recent_file, skiprows=1, na_values='—', usecols=columns)
    return df

def map_columns(df: pd.DataFrame, column_mapping: dict) -> pd.DataFrame:
    """Rename columns using the provided mapping dictionary."""
    df = df.rename(columns=column_mapping)
    return df

def clean_and_convert_columns(df: pd.DataFrame) -> pd.DataFrame:
    df['Ordered Revenue'] = df['Ordered Revenue'].replace(r'[$,]', '', regex=True).fillna(0).astype(float)
    df['Ordered Revenue - Prior Period'] = df['Ordered Revenue - Prior Period'].replace(r'[%,]', '', regex=True).fillna(0).astype(float) / 100
    df['Ordered Revenue - Same Period Last Year'] = df['Ordered Revenue - Same Period Last Year'].replace(r'[%,]', '', regex=True).fillna(0).astype(float) / 100
    df['Ordered Units'] = df['Ordered Units'].replace(',', '', regex=True).fillna(0).astype(int)
    df['Ordered Units - Prior Period'] = df['Ordered Units - Prior Period'].replace(r'[%,]', '', regex=True).fillna(0).astype(float) / 100
    df['Ordered Units - Same Period Last Year'] = df['Ordered Units - Same Period Last Year'].replace(r'[%,]', '', regex=True).fillna(0).astype(float) / 100
    return df

def filter_nonzero_rows(df: pd.DataFrame) -> pd.DataFrame:
    # Filter out rows where Ordered Revenue or Ordered Units is 0 or NaN
    df_filtered = df[(df['Ordered Revenue'].notna()) & 
                     (df['Ordered Units'].notna()) &
                     (df['Ordered Revenue'] != 0) & 
                     (df['Ordered Units'] != 0)
                     & (df['Ordered Revenue - Prior Period']!= -1)
                     & (df['Ordered Revenue - Same Period Last Year']!= -1)
                     & (df['Ordered Units - Same Period Last Year']!= -1)
                     & (df['Ordered Units - Prior Period']!= -1)
                     ]
    return df_filtered

def uploader_weeklysales() -> pd.DataFrame:
    folder_path = Path(r'G:\SALES\Amazon\RBH\DOWNLOADED_FILES')
    file_glob_sales_weekly = r'\*Sales*Weekly*csv'
    cols_sales_weekly = [
        'ASIN', 'Ordered Revenue', 'Ordered Revenue - Prior Period (%)', 'Ordered Revenue - Same Period Last Year (%)',
        'Ordered Units', 'Ordered Units - Prior Period (%)', 'Ordered Units - Same Period Last Year (%)'
    ]
    
    # When Vendor Central makes updates to the columns, update the mapping here
    column_mapping = {
        'ASIN': 'ASIN',
        'Ordered Revenue': 'Ordered Revenue',
        'Ordered Revenue - Prior Period (%)': 'Ordered Revenue - Prior Period',
        'Ordered Revenue - Same Period Last Year (%)': 'Ordered Revenue - Same Period Last Year',
        'Ordered Units': 'Ordered Units',
        'Ordered Units - Prior Period (%)': 'Ordered Units - Prior Period',
        'Ordered Units - Same Period Last Year (%)': 'Ordered Units - Same Period Last Year'
    }
    
    df_sales_weekly = load_sales_data(folder_path, file_glob_sales_weekly, cols_sales_weekly)
    df_sales_weekly = map_columns(df_sales_weekly, column_mapping)
    df_sales_weekly = clean_and_convert_columns(df_sales_weekly)

    # Apply your custom logic to set values to 0 where needed
    df_sales_weekly = filter_nonzero_rows(df_sales_weekly)

    # Calculate the OR and OU columns gracefully using np.divide
    df_sales_weekly['or_pp'] = np.divide(
        df_sales_weekly['Ordered Revenue'], 
        1 + df_sales_weekly['Ordered Revenue - Prior Period'],
        out=np.zeros_like(df_sales_weekly['Ordered Revenue'], dtype=float), 
        where=(df_sales_weekly['Ordered Revenue - Prior Period'] != 0)
    )

    df_sales_weekly['or_ly'] = np.divide(
        df_sales_weekly['Ordered Revenue'], 
        1 + df_sales_weekly['Ordered Revenue - Same Period Last Year'],
        out=np.zeros_like(df_sales_weekly['Ordered Revenue'], dtype=float), 
        where=(df_sales_weekly['Ordered Revenue - Same Period Last Year'] != 0)
    )

    df_sales_weekly['ou_pp'] = np.divide(
        df_sales_weekly['Ordered Units'], 
        1 + df_sales_weekly['Ordered Units - Prior Period'],
        out=np.zeros_like(df_sales_weekly['Ordered Units'], dtype=float), 
        where=(df_sales_weekly['Ordered Units - Prior Period'] != 0)
    )

    df_sales_weekly['ou_ly'] = np.divide(
        df_sales_weekly['Ordered Units'], 
        1 + df_sales_weekly['Ordered Units - Same Period Last Year'],
        out=np.zeros_like(df_sales_weekly['Ordered Units'], dtype=float), 
        where=(df_sales_weekly['Ordered Units - Same Period Last Year'] != 0)
    )

    return df_sales_weekly

def main():
    pd.options.display.float_format = '{:.2f}'.format
    df = uploader_weeklysales()
    print(df.info())
    print(df.head())

    # Calculate and print the totals for each column
    totals = df.sum(numeric_only=True)
    print("\nColumn Totals:")
    print(totals)

if __name__ == '__main__':
    main()