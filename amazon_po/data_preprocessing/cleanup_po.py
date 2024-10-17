import sys
import pandas as pd
import numpy as np
from pathlib import Path

# Add the parent directory to the sys.path so Python can find functions.py
sys.path.append(str(Path(__file__).resolve().parent.parent))

# File paths and connection strings
# Note that the downloaded file must be saved to this location

# Path components
base_dir = Path(r'G:\SALES\Amazon\PURCHASE ORDERS\atelier\po_analysis')
file_name = 'PurchaseOrderItems.csv'
PO_FILE_PATH = base_dir / file_name

def upload_po(file=PO_FILE_PATH) -> pd.DataFrame:
    """
    Upload the Purchase Order file from Vendor Central and select relevant columns.
    
    Parameters:
    - file: Path to the Purchase Order file.
    
    Returns:
    - DataFrame with selected columns and appropriate data types.
    """
    columns = ['ASIN', 'External ID', 'Accepted quantity (units)', 'Requested quantity (units)', 
               'Total accepted cost', 'Cost price','Total requested cost']
    
    return pd.read_csv(file,
                       usecols=columns,
                       dtype={'External ID': 'object',
                              'ASIN':'object',
                              'Total accepted cost': 'float',
                              'Total requested cost':'float'})

def rename_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rename columns for better clarity and consistency.
    Vendor Central tends to change names often so this is a way to keep names more consistent.
    
    Parameters:
    - df: DataFrame to rename columns in.
    
    Returns:
    - DataFrame with renamed columns.
    """
    col_dict = {'External ID': 'ISBN',
                'Accepted quantity (units)':'Accepted Quantity',
                'Requested quantity (units)':'Requested quantity',
                'Total requested cost':'Requested Cost'
                }
    
    return df.rename(columns=col_dict)

def po_clean(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean and aggregate the Purchase Order data.
    
    Parameters:
    - df: DataFrame to clean.
    
    Returns:
    - Cleaned and aggregated DataFrame.
    """
    df = df.groupby(by=['ASIN', 'ISBN']).agg(
        {'Accepted Quantity': 'sum', 
         'Requested quantity': 'sum', 
         'Total accepted cost': 'sum',
         'Requested Cost':'sum', 
         'Cost price': 'mean'}
    ).reset_index()
    
    df['ISBN'] = df['ISBN'].str.zfill(13)
    df['ASIN'] = df['ASIN'].str.zfill(10)
    return df

def get_cleaned_po() -> pd.DataFrame:
    """
    Upload, clean, and prepare the Purchase Order file for further analysis.
    
    Returns:
    - Cleaned DataFrame ready for analysis.
    """
    df = upload_po()
    df = rename_columns(df)
    df = po_clean(df)
    return df

def main():
    """
    Main function to execute the Purchase Order data preparation.
    """
    df = get_cleaned_po()
    print(df.info())
    print(df.sum())
    print(df.head())

if __name__ == "__main__":
    main()