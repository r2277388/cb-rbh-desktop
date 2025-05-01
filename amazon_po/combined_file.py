#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import numpy as np
from datetime import datetime as dt
from functools import cache

# Import custom functions
from data_preprocessing.cleanup_catalog import get_cleaned_catalog
from data_preprocessing.cleanup_ebs_item import get_cleaned_item
from data_preprocessing.cleanup_inventorydetail import get_cleaned_inv
from data_preprocessing.cleanup_po_v2 import get_cleaned_po

def merge_clean_tables(df_po, df_item, df_avail)-> pd.DataFrame:
    df_po['ISBN'] = df_po['ISBN'].str.zfill(13)
    df_avail['ISBN'] = df_avail['ISBN'].str.zfill(13)
    df_item['ISBN'] = df_item['ISBN'].str.zfill(13)
    
    df_po = df_po.loc[:, ['ASIN', 'ISBN', 'Product Group'
                          ,'Accepted Quantity', 'Requested quantity'
                          ,'Total accepted cost', 'Cost price'
                          ,'Requested Cost']
                      ]
    df_item = df_item.loc[:, ['ISBN', 'Publisher', 'pgrp', 'title', 'pub', 'price']]
    
    df = pd.merge(df_po, df_item, how='left', left_on='ISBN', right_on='ISBN').fillna(0)
    df['Delta'] = df['Accepted Quantity'] - df['Requested quantity']
    df = pd.merge(df, df_avail, how='left', on='ISBN')
    df['Lost Sales'] = (df['Requested quantity'] * df['Cost price']) - df['Total accepted cost']
    
    return df.loc[:, ['ASIN', 'ISBN', 'title', 'Publisher', 'pgrp', 'pub', 'Product Group', 
                      'Available To Sell', 'Reprint Due Date', 'Reprint Quantity', 
                      'Requested quantity', 'Accepted Quantity', 'Delta', 'Total accepted cost'
                      ,'Lost Sales','Requested Cost']
                  ]

# Output Paths
def get_merged_files():
    # Upload clean dataframes
    df_po = get_cleaned_po()
    df_cat = get_cleaned_catalog()
    df_avail = get_cleaned_inv()
    df_item = get_cleaned_item()

    # Merge df_po dataframe and the amazon catalog file
    df_cat_po = pd.merge(df_po, df_cat, how='left', on='ISBN').fillna(0)
    df_cat_po.rename(columns={'ASIN_x':'ASIN'},inplace=True)
    
    # Adding some more columns
    df_combined = merge_clean_tables(df_cat_po, df_item, df_avail)
    df_combined = df_combined[df_combined['Publisher'] != 0]
    return df_combined

def main():
    df = get_merged_files()
    print(df.info())
    print(df.head())
    print('Unique Publisher:', df['Publisher'].unique())
    
if __name__ == "__main__":
    main()