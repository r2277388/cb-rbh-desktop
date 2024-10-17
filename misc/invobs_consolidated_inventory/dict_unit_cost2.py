import pandas as pd
import numpy as np
from load_consolidated_inventory import consolidate_inventory

def df_to_nested_dict():
    df = consolidate_inventory()
    df = df[['ISBN', 'val_cbc', 'val_hbg', 'val_cbp', 'units_cbc', 'units_hbg', 'units_cbp']]
    
    df['uc_cbc'] = df['val_cbc'] / df['units_cbc']
    df['uc_hbg'] = df['val_hbg'] / df['units_hbg']
    df['uc_cbp'] = df['val_cbp'] / df['units_cbp']
    
    # Replace infinities and NaNs
    df.replace([np.inf, -np.inf], np.nan, inplace=True)
    df.fillna(0, inplace=True)
    
    # Define a function to adjust unit costs
    def adjust_unit_cost(row):
        uc_cbc, uc_hbg, uc_cbp = row['uc_cbc'], row['uc_hbg'], row['uc_cbp']
        
        # Check if any unit cost is zero and replace it with the max of the others
        if uc_cbc == 0.0:
            row['uc_cbc'] = max(uc_hbg, uc_cbp)
        if uc_hbg == 0.0:
            row['uc_hbg'] = max(uc_cbc, uc_cbp)
        if uc_cbp == 0.0:
            row['uc_cbp'] = max(uc_cbc, uc_hbg)

        return row

    # Apply the adjust_unit_cost function to each row
    df = df.apply(adjust_unit_cost, axis=1)
    
    # Create the nested dictionary
    result_dict = df.set_index('ISBN').apply(
        lambda row: {'uc_cbc': row['uc_cbc'], 'uc_hbg': row['uc_hbg'], 'uc_cbp': row['uc_cbp']},
        axis=1
    ).to_dict()

    return result_dict

def main():
    dict_uc = df_to_nested_dict()
    print(dict_uc['0810073340527'])  # Example for checking

if __name__ == "__main__":
    main()
