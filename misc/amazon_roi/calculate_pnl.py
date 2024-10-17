import pandas as pd
import numpy as np
from pathlib import Path

from loader_ams import upload_ams_raw
from loader_unit_costs import upload_uc
from loader_ebs_item import upload_item_raw

def load_and_merge_data():
    df_item = upload_item_raw()
    df_ams = upload_ams_raw()
    df_uc = upload_uc()

    # Merging the data
    df = df_ams.merge(df_item, on='ISBN', how='inner').merge(df_uc, on='ISBN', how='inner')

    # Selecting and ordering the relevant columns
    columns = ['ISBN', 'Title', 'pub', 'pgrp', 'price', 'SeasonOnly', 'pub_date', 'UC', 'Period', 'PT', 'FT', 'RF',
               'AMS Spend', 'AMS Sales', 'AMS Units']
    df = df[columns]

    return df

def calculate_metrics(df, amazon_discount=0.54, return_rate=0.013, royalty_rate=0.12, 
                      fulfillment_rate=0.025, fulfillment_return_rate=0.035):
    
    # Net Sales Calculations
    df['Retail'] = df['AMS Units'] * df['price']
    df['Gross'] = df['Retail'] * (1 - amazon_discount)
    df['Returns'] = -df['Gross'] * return_rate
    df['Net Sales'] = df['Gross'] + df['Returns']

    # Cost of Sales Calculations
    df['COGS'] = df['AMS Units'] * df['UC'] * (1 - return_rate)
    df['Royalties'] = np.where(df['RF'] == 'R', df['Net Sales'] * royalty_rate, 0)
    df['Fulfillment'] = df['Gross'] * fulfillment_rate + df['Returns'] * fulfillment_return_rate
    df['COS'] = df['AMS Spend'] + df['Fulfillment'] + df['Royalties'] + df['COGS']

    # Gross Margin and Gross Margin %
    df['Gross Margin'] = df['Net Sales'] - df['COS']
    df['GM %'] = df['Gross Margin'] / df['Gross']
    
    # Handle division by zero in GM %
    df['GM %'] = df['GM %'].where(df['Gross'] != 0, 0)
    
    return df

def main():
    df = load_and_merge_data()
    df = calculate_metrics(df)
    
    # Save to Excel with a specific file name using a with statement
    file_path = Path(r"C:\Users\rbh\Desktop\output_data.xlsx")
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    
    print(df.head())
    df.info()

if __name__ == '__main__':
    main()