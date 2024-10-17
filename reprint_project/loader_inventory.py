import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta
import joblib

CACHE_DIR = Path('E:/My Drive/code/reprint_project/cache')
CACHE_FILE = CACHE_DIR / 'inventory_cache.pkl'
CACHE_EXPIRY_HOURS = 10

def load_data(file_path, columns):
    dtype = {'ISBN': str}
    return pd.read_excel(file_path,dtype=dtype, usecols=columns, engine='openpyxl')

def filter_data(df):
    df.dropna(subset=['ISBN','Season','Status'], inplace=True)
    df = df[df['Publisher']=='Chronicle']
    df.drop(columns=['Publisher'], inplace=True)
    
    df['ISBN'] = df['ISBN'].str.zfill(13)
    
    return df

def load_inventory(reload=False):
    if not reload and CACHE_FILE.exists():
        cache_time = datetime.fromtimestamp(CACHE_FILE.stat().st_mtime)
        if datetime.now() - cache_time < timedelta(hours=CACHE_EXPIRY_HOURS):
            return joblib.load(CACHE_FILE)
    
    folder_location = Path('G:/OPS/Inventory/Daily/Finance_Only')
    file_name = 'Inventory Detail.xlsx'
    location = folder_location / file_name  

    columns = ['ISBN','Publisher','Status','Season','On Sale Date',
               'Reprint Due Date','Reprint Quantity','Available To Sell',
               'Frozen','Reprint Freeze']

    df = load_data(location, columns)
    df = filter_data(df)
    
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    joblib.dump(df, CACHE_FILE)
    return df

def main():
    df = load_inventory()
    print(df.info())

if __name__ == '__main__':
    main()