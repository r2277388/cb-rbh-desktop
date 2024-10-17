import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
import joblib

CACHE_DIR = Path('E:/My Drive/code/reprint_project/cache')
CACHE_FILE = CACHE_DIR / 'hachette_order_status_cache.pkl'
CACHE_EXPIRY_HOURS = 10

def load_data(file_path, columns):
    dtype = {'ISBN': str}
    return pd.read_excel(file_path, usecols=columns, dtype=dtype, engine='openpyxl')

def filter_data(df):
    df.dropna(subset=['ISBN','Title'], inplace=True)
    df = df[df['Publisher Name']=='Chronicle Books LLC']
    df = df[df['Title'].str.contains('Display',na=False)]
    df.drop(columns=['Publisher Name','Title'], inplace=True)
    
    df['ISBN'] = df['ISBN'].str.zfill(13)
    
    return df

def load_hachette_orders(reload=False):
	if not reload and CACHE_FILE.exists():
		cache_time = datetime.fromtimestamp(CACHE_FILE.stat().st_mtime)
		if datetime.now() - cache_time < timedelta(hours=CACHE_EXPIRY_HOURS):
			return joblib.load(CACHE_FILE)
	
	file_path = Path("G:/OPS/Inventory/Daily/Finance_Only/Hachette Order Status.xlsx")
	columns = ['ISBN','Title','Publisher Name','Order Status','Entered Date','Release Date','Quantity']

	df = load_data(file_path, columns)
	df = filter_data(df)
	
	CACHE_DIR.mkdir(parents=True, exist_ok=True)
	joblib.dump(df, CACHE_FILE)
	return df

def main():
	df = load_hachette_orders()
	print(df.head(10))

if __name__ == '__main__':
	main()