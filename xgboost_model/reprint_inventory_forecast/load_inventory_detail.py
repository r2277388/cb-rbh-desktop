import pandas as pd
import numpy as np

def fetch_inventory_detail():
    
    inventory_detail = r'G:\OPS\Inventory\Daily\Finance_Only\Inventory Detail.xlsx'
    
    try:
        df = pd.read_excel(inventory_detail
                        ,usecols=['ISBN','Publisher','Reprint Due Date','Reprint Quantity','Available To Sell','Frozen','Reprint Freeze']
                        ,parse_dates = ['Reprint Due Date']
                        ,dtype={'ISBN':'str','Publisher':'str'}
        )
        numeric_cols = ['Reprint Quantity','Available To Sell','Frozen','Reprint Freeze']
        
        # Make sure numeric columns are numeric
        df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric,errors='coerce')
        
        # Remove rows with out any quantity data at all.
        df[numeric_cols] = df[numeric_cols].replace(0,np.nan)
        df = df.dropna(subset=numeric_cols,how='all')
        
        # Make sure ISBN's have retained their 13 digits
        df['ISBN'] = df['ISBN'].str.zfill(13)
        
        # drop rows with missing ISBN's
        df = df.dropna(subset='ISBN')
        
        # filter for CB only
        df = df.loc[df['Publisher'] == 'Chronicle']  
        
        df.reset_index(drop=True, inplace=True)
        
        return df
    
    except FileNotFoundError:
        print(f'File not found: {inventory_detail}')
        return None
    except Exception as e:
        print(f'Error: {e}')
        return None

if __name__ == '__main__':
    
    df = fetch_inventory_detail()
    print(df.info())
    print(df.head())
