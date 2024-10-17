import sys
from pathlib import Path
import pandas as pd
import numpy as np

# Add the parent directory to the sys.path so Python can find functions.py
sys.path.append(str(Path(__file__).resolve().parent.parent))

# File Path
# Path components
base_dir = Path(fr'G:\OPS\Inventory\Daily\Finance_Only')
file_name = 'Inventory Detail.xlsx'

AVAIL_FILE_PATH = base_dir / file_name

def upload_inventory_detail():
    return pd.read_excel(AVAIL_FILE_PATH
                        ,usecols=['ISBN','Available To Sell','On Sale Date','Reprint Due Date','Reprint Quantity']
                        ,dtype={'ISBN': 'object'}
                        ,engine='openpyxl')

def avail_clean(df)-> pd.DataFrame:
    df = df.loc[~df.ISBN.isnull()].reset_index(drop=True)
    df['ISBN'] = df['ISBN'].astype(str).apply(lambda x: x.zfill(13))
    df['Reprint Quantity'] = df['Reprint Quantity'].fillna(0)
    return df

def get_cleaned_inv():
    df = upload_inventory_detail()
    df = avail_clean(df)
    return df

if __name__ == "__main__":
    df = get_cleaned_inv()
    print(df.head())