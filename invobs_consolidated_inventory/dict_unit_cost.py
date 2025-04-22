import pandas as pd
import numpy as np
from load_consolidated_inventory import consolidate_inventory

def df_to_nested_dict():
    df = consolidate_inventory()
    df = df[['ISBN', 'val_cbc', 'val_hbg', 'val_cbp','units_cbc','units_hbg','units_cbp']]
    
    df['uc_cbc'] = df['val_cbc'] / df['units_cbc']
    df['uc_hbg'] = df['val_hbg'] / df['units_hbg']
    df['uc_cbp'] = df['val_cbp'] / df['units_cbp']
    
    # Replace infinities and NaNs
    df.replace([np.inf, -np.inf], np.nan, inplace=True)
    # Optionally fill NaNs with a default value, like 0
    df.fillna(0, inplace=True)
    
    # Use apply to create a nested dictionary where ISBN is the key
    result_dict = df.set_index('ISBN').apply(
        lambda row: {'uc_cbc': row['uc_cbc'], 'uc_hbg': row['uc_hbg'], 'uc_cbp': row['uc_cbp']},
        axis=1
    ).to_dict()

    return result_dict

def main():
    
    dict_uc = df_to_nested_dict()

    print(dict_uc['0810073340527'])
    
if __name__ == "__main__":
    main()