import pandas as pd
import numpy as np
from pathlib import Path

# Function to load the data
def load_data(file_path):
    df = pd.read_excel(file_path
                       ,skiprows=3
                       ,na_values=['']
                       #,sheet_name='ALL_Consolidated_Inventories_V_'
                       ,keep_default_na=False)
    df = df.fillna(0)
    df.columns = [
        "Item", "pgrp", "bkft", "title", "price", "ship", "ans", "uc_cbc", "uc_hbg", 
        "uc_cbp", "uc_all", "val_cbc", "units_cbc", "val_hbg", "units_hbg", 
        "val_cbp", "units_cbp", "total_Units", "total_value"
    ]
    return df

# Function to apply filters
def apply_filters(df, pgrp_filter, bkft_filter):
    df = df[df['pgrp'].isin(pgrp_filter)]
    df = df[df['bkft'].isin(bkft_filter)]
    df.reset_index(drop=True, inplace=True)
    return df

# Main function to execute the steps
def consolidate_inventory():
    file_consolided_inventory = Path(r"F:\2024\Analysis\invobs\Q4\All_Consolidated_Inventories_dec24.xlsx")
    
    # Load data
    df = load_data(file_consolided_inventory)
    
    # Filters
    pgrp_filter = ['ART', 'LIF', 'CHL', 'ENT', 'FWN', 'PTC', 'RID', 'GAM', 'BAR-LIF', 'BAR-ENT', 'BAR-ART', 'CCB', 'CPB','CPA']
    bkft_filter = ['BK', 'FT','CP','RP']
    
    # Apply filters
    df_filtered = apply_filters(df, pgrp_filter, bkft_filter)
    
    # Rename column correctly
    df_filtered = df_filtered.rename(columns={'Item': 'ISBN'})
    
    # Ensure ISBN is 13 digits
    df_filtered['ISBN'] = df_filtered['ISBN'].astype(str).str.zfill(13)
    
    return df_filtered

def main():
    df = consolidate_inventory()  # Call only once and store the result
    print(df.info())              # Print info for the dataframe
    print(df.head())              # Print the first few rows of the dataframe
    print(df[df['ISBN']=='9781797233086'])   # Print the number of unique ISBNs
if __name__ == "__main__":
    main()
