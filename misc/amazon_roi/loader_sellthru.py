import pandas as pd
import numpy as np
from pathlib import Path
from functions import asin_isbn_key

# Define a path for the cache file
cache_file_path = Path('compiled_data_cache.pkl')

def compile_csv_files(folder_path: Path = Path('vc_sellthru'), output_file: str = None, use_cache: bool = True) -> pd.DataFrame:
    if use_cache and cache_file_path.exists():
        print("Loading data from cache...")
        return pd.read_pickle(cache_file_path)

    print("Processing files...")
    asin_isbn_df = asin_isbn_key()
    dataframes = []

    for file in folder_path.glob('*.csv'):
        file_stem = file.stem
        date_part = file_stem.split('_')[-2]
        period = pd.to_datetime(date_part).strftime('%Y%m')

        usecols = [
            'ASIN', 'Ordered Revenue', 'Ordered Revenue - Prior Period',
            'Ordered Revenue - Same Period Last Year',
            'Customer Returns','Customer Returns - Prior Period', 'Customer Returns - Same Period Last Year'
            ]
        df = pd.read_csv(file, skiprows=1, usecols=usecols)

        # Convert monetary columns
        df['Period'] = period
        df['Ordered Revenue'] = df['Ordered Revenue'].replace('[\$,]', '', regex=True).astype(float)
        df['Customer Returns'] = df['Customer Returns'].replace('[,]', '', regex=True).astype(float)

        # Convert percentage columns
        columns_with_percent = [
            'Ordered Revenue - Prior Period',
            'Ordered Revenue - Same Period Last Year',
            'Customer Returns - Prior Period',
            'Customer Returns - Same Period Last Year'
        ]

        for col in columns_with_percent:
            df[col] = df[col].replace('[\%,]', "", regex=True).replace('', np.nan).astype(float)

        dataframes.append(df)

    combined_df = pd.concat(dataframes, ignore_index=True)
    combined_df = combined_df.merge(asin_isbn_df, on='ASIN', how='left')
    combined_df.dropna(subset=['ISBN'], inplace=True)

    # Save the processed dataframe to cache
    combined_df.to_pickle(cache_file_path)
    print(f"Data cached to {cache_file_path}")
    
    if output_file:
        combined_df.to_csv(output_file, index=False)
        print(f"All files have been compiled into {output_file}")
    
    return combined_df

def sellthru_preparation(use_cache: bool = True):
    df = compile_csv_files(use_cache=use_cache)
    
    # Drop 'ASIN' column
    df.drop(columns=['ASIN'], inplace=True)

    # Ensure the columns exist before reordering
    if all(col in df.columns for col in ['ISBN', 'Period']):
        df = df[['ISBN', 'Period'] + [col for col in df.columns if col not in ['ISBN', 'Period']]]

    # Drop rows with missing 'Ordered Revenue'
    df.dropna(subset=['Ordered Revenue'], inplace=True)
    
    # Convert percentage columns to decimal and calculate revenue/returns
    percentage_columns = [
        'Ordered Revenue - Prior Period',
        'Ordered Revenue - Same Period Last Year',
        'Customer Returns - Prior Period',
        'Customer Returns - Same Period Last Year'
    ]
    
    for col in percentage_columns:
        df[col] = df[col].fillna(0) / 100.0
        
    df['Ordered Revenue - Prior Period'] = df['Ordered Revenue'] * (1 + df['Ordered Revenue - Prior Period'])
    df['Ordered Revenue - Same Period Last Year'] = df['Ordered Revenue'] * (1 + df['Ordered Revenue - Same Period Last Year'])
    df['Customer Returns - Prior Period'] = df['Customer Returns'] * (1 + df['Customer Returns - Prior Period'])
    df['Customer Returns - Same Period Last Year'] = df['Customer Returns'] * (1 + df['Customer Returns - Same Period Last Year'])

    # Drop unnecessary columns
    columns_to_drop = ['Customer Returns','Customer Returns - Prior Period','Customer Returns - Same Period Last Year']
    df.drop(columns=columns_to_drop,inplace=True)

    return df

def main():
    sellthru = sellthru_preparation(use_cache=True)
    print(sellthru.head())
    print(sellthru.info())
    
if __name__ == '__main__':
    main()