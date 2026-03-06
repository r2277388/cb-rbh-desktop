import sys
import pandas as pd
from pathlib import Path

# Add the parent directory to sys.path
sys.path.append(str(Path(__file__).resolve().parent))
sys.path.append(str(Path(__file__).resolve().parent.parent))
from loader_osd import upload_osd

# Path components
base_dir = Path(fr'J:\Metadata Reports')
file_name = 'Oracle YPTICOD.xlsx'
    
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
    
    df_osd = upload_osd()
    df = df.merge(df_osd, on='ISBN', how='left')    
    df.rename(columns={'osd': 'Release Date'}, inplace=True)
    return df

if __name__ == "__main__":
    df = get_cleaned_ypticod()
    print(df.head())