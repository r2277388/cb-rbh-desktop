import sys
import pandas as pd
from pathlib import Path

# Add the parent directory to the sys.path so Python can find functions.py
sys.path.append(str(Path(__file__).resolve().parent.parent))

# Path components
base_dir = Path(fr'\\sierra\groups\READTHIS\FINANCE\SALES\23sales')
file_name = 'Oracle YPTICOD.xlsx'

ypticod_list = list(base_dir.glob("Oracle YPTICOD*.xlsx"))
if not ypticod_list:
    raise FileNotFoundError(f"No 'xlsx' file found that starts with 'Oracle Ypticode' in {base_dir}.")
file_name = ypticod_list[0]  # Get the first match

# File paths and connection strings
YPTICOD_FILE_PATH = base_dir / file_name

def upload_ypticod(file = YPTICOD_FILE_PATH):
    return pd.read_excel(file
                        ,usecols=['ISBN','ISBN10']
                        ,dtype={'ISBN': 'object','ISBN10':'object'}
                        ,engine='openpyxl')

def clean_ypticod(df)-> pd.DataFrame:
    df['ISBN'] = df['ISBN'].astype(str).str.zfill(13)
    df['ISBN10'] = df['ISBN10'].astype(str).str.zfill(10)
    df.rename(columns={'ISBN10': 'ASIN'}, inplace=True)
    return df

def get_cleaned_ypticod():
    df = upload_ypticod()
    df = clean_ypticod(df)
    return df

if __name__ == "__main__":
    df = get_cleaned_ypticod()
    print(df.head())