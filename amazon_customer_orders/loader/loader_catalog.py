import pandas as pd
from pathlib import Path
import os

def get_latest_file(folder_path: Path, pattern: str) -> Path:
    """Return the latest file in the folder matching the given pattern."""
    files = list(folder_path.glob(pattern))
    if not files:
        raise FileNotFoundError(f"No files found in {folder_path} with pattern {pattern}")
    latest_file = max(files, key=os.path.getctime)
    return latest_file

def read_catalog_file(file_path: Path, columns: list, date_column: str) -> pd.DataFrame:
    """Read the catalog CSV file and return a DataFrame with parsed dates."""
    df = pd.read_csv(
        file_path,
        skiprows=1,
        na_values='—',
        usecols=columns
    )
    
    # Convert the date column to datetime using the correct format
    df[date_column] = pd.to_datetime(df[date_column], errors='coerce').dt.normalize()
    
    return df

def process_latest_catalog(folder_path: str, pattern: str, columns: list, date_column: str) -> pd.DataFrame:
    """Fetch the latest catalog file, read it, and return the DataFrame."""
    folder_path = Path(folder_path)
    latest_file = get_latest_file(folder_path, pattern)
    df = read_catalog_file(latest_file, columns, date_column)
    return df
    
def data_catalog():
    folder_path = r'G:\SALES\Amazon\RBH\DOWNLOADED_FILES'
    pattern = '*Catalog*csv'
    columns = ['ASIN', 'EAN', 'ISBN-13', 'Model Number', 'Release Date']
    date_column = 'Release Date'
    
    df_catalog = process_latest_catalog(folder_path, pattern, columns, date_column)
    return df_catalog
    
def main():
    df = data_catalog()
    print(df.info())
    print(df.head())

if __name__ == "__main__":
    main()