import pandas as pd
from pathlib import Path
import os

def get_latest_inventory_file(folder_path: Path, pattern: str) -> Path:
    """Return the latest inventory file in the folder matching the given pattern."""
    files = list(folder_path.glob(pattern))
    
    # Check if files were found to avoid errors
    if not files:
        raise FileNotFoundError(f"No files found in {folder_path} with pattern {pattern}")
    
    # Get the latest file based on creation time
    latest_file = max(files, key=os.path.getctime)
    return latest_file

def read_inventory_file(file_path: Path, columns: list) -> pd.DataFrame:
    """Read the inventory CSV file and return a DataFrame."""
    df = pd.read_csv(
        file_path,
        skiprows=1,
        na_values='—',
        usecols=columns,
        dtype={
            'ASIN': object,
            'Inventory Qty': int
        }
    )
    return df


def rename_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename specific columns in the DataFrame."""
    rename_dict = {
        'ASIN': 'ASIN',
        'Unfilled Customer Ordered Units': 'Orders'
    }
    df = df.rename(columns=rename_dict)
    return df

def process_inventory_data(folder_path: str, pattern: str, columns: list) -> pd.DataFrame:
    """Main function to load, process, and return the inventory DataFrame."""
    folder_path = Path(folder_path)
    
    # Get the latest inventory file
    file_inventory = get_latest_inventory_file(folder_path, pattern)
    
    # Read the inventory file
    df_inventory = read_inventory_file(file_inventory, columns)
    
    # Rename columns
    df_inventory = rename_columns(df_inventory)
    
    return df_inventory

def data_inventory():
    # Define the folder path and pattern for inventory filesW
    folder_path = r'G:\SALES\Amazon\RBH\DOWNLOADED_FILES'
    pattern = '*Inventory*csv'
    columns = ['ASIN', 'Unfilled Customer Ordered Units']

    # Process the inventory data and print the first few rows
    df = process_inventory_data(folder_path, pattern, columns)
    df['Orders'] = df['Orders'].replace(',',"",regex=True).fillna(0).astype(int)
    df = df[df['Orders']>0]
    
    return df

def main():
    df =  data_inventory()
    print(df.info())
    print(df.head())

if __name__ == "__main__":
    main()