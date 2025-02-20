import sys
from pathlib import Path
import pandas as pd
import numpy as np

# Add the parent directory to the sys.path so Python can find functions.py
sys.path.append(str(Path(__file__).resolve().parent.parent))

# File
# Path components
base_dir = Path(fr'G:\SALES\Amazon\RBH\DOWNLOADED_FILES')

# Find the file that starts with "Catalog"
catalog_files = list(base_dir.glob("Catalog*.csv"))
if not catalog_files:
    raise FileNotFoundError(f"No 'csv' file starting with 'Catalog' found in the following folder {base_dir}.")
file_name = catalog_files[0] # finding the catalog file as long as it starts with Catalog.

CAT_FILE_PATH = base_dir / file_name

def upload_catalog(file = CAT_FILE_PATH):
    return pd.read_csv(file
                        ,skiprows=1
                        ,usecols=['EAN','ASIN','ISBN','Product Group']
                        ,na_values=['UNKNOWN','—']
                        ,dtype={'EAN': 'object','ISBN':'object'}
                        )

def cat_clean(df)-> pd.DataFrame:
    df = df.copy()
    filt1 = (df['EAN'].isnull()) & (df['ISBN'].isnull())
    df = df.loc[~filt1]
    #df.rename(columns={"EAN": "ISBN"}, errors="raise", inplace=True)
    df['ISBN'] = np.where(df['EAN'].isna(), df['ISBN'], df['EAN'])
    #df.drop(['ISBN-13'], axis=1, inplace=True)
    df['ISBN'] = df['ISBN'].str.zfill(13)
    df['ASIN'] = df['ASIN'].str.zfill(10)
    return df

def get_cleaned_catalog():
    df = upload_catalog()
    df = cat_clean(df)
    df = df.drop_duplicates(subset='ISBN', keep="first")
    return df

# Optional: You can include this to run the script independently if needed
if __name__ == "__main__":
    df = get_cleaned_catalog()
    print(df.head())