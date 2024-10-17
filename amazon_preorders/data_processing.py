import pandas as pd
import numpy as np
import datetime

from loader_ebs_item import data_item
from loader_inventory import data_inventory
from asin_isbn_converter import asin_isbn_conversion

def merge_catalog_inventory() -> pd.DataFrame:
    df_item = data_item()
    df_inventory = data_inventory()
    df_catalog = asin_isbn_conversion()
    
    """Merge catalog and inventory DataFrames on ASIN and clean up the resulting DataFrame."""
    df = pd.merge(df_inventory, df_catalog, on='ASIN', how='inner')
    df['Orders'] = df['Orders'].fillna(0)
    
    df = pd.merge(df,df_item,on='ISBN',how='inner')
    df = df[df['Release Date'] > datetime.datetime.today()]
    
    df = df[['ASIN', 'ISBN', 'title', 'publisher', 'Release Date', 'Orders']]
    
    df = df.sort_values(by='Orders',ascending=False)
    
    return df

if __name__ == '__main__':
    df = merge_catalog_inventory()
    print(df.head())