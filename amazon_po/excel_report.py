import pandas as pd
from pathlib import Path
from datetime import datetime
from combined_file import get_merged_files

def load_data():
    """
    Loads the merged data from the combined_file module.
    """
    return get_merged_files()

def aggregate_data(df):
    """
    Aggregates data by 'Publisher' and 'pgrp'.
    
    Parameters:
        df (pd.DataFrame): The dataframe containing the purchase order data.
    
    Returns:
        agg_pub (pd.DataFrame): Aggregated data by Publisher.
        agg_pgrp (pd.DataFrame): Aggregated data by product group (pgrp).
    """
    agg_pub = df.groupby('Publisher').agg({'Total accepted cost': 'sum'}).reset_index()
    agg_pgrp = df.groupby('pgrp').agg({'Total accepted cost': 'sum'}).reset_index()
    return agg_pub, agg_pgrp

def generate_filename(base_path):
    """
    Generates a formatted filename in the format 'yyyy_mm_dd_Sunday_Month.xlsx'.
    
    Parameters:
        base_path (str or Path): The base directory where the file should be saved.
    
    Returns:
        filename (Path): The complete path with the formatted filename.
    """
    now = datetime.now()
    # Format the date as 'yyyy_mm_dd', day of the week, and full month name
    formatted_date = now.strftime('%Y_%m_%d')+"_Sunday_Monday"
    
    return Path(base_path) / f"{formatted_date}.xlsx"

def save_to_excel(df, agg_pub, agg_pgrp, filename):
    """
    Saves the raw data and aggregated data into an Excel file with multiple sheets.
    
    Parameters:
        df (pd.DataFrame): The raw purchase order data.
        agg_pub (pd.DataFrame): Aggregated data by Publisher.
        agg_pgrp (pd.DataFrame): Aggregated data by product group.
        filename (Path): The full path for saving the Excel file.
    """
    with pd.ExcelWriter(filename) as writer:
        df.to_excel(writer, sheet_name='Raw_PO', index=False)
        agg_pub.to_excel(writer, sheet_name='Publisher', index=False)
        agg_pgrp.to_excel(writer, sheet_name='Product Group', index=False)

def main():
    """
    Main function that orchestrates loading, aggregating, and saving the purchase order data.
    """
    df = load_data()
    agg_pub, agg_pgrp = aggregate_data(df)
    filename = generate_filename(r'G:\SALES\Amazon\PURCHASE ORDERS\2025')
    save_to_excel(df, agg_pub, agg_pgrp, filename)

if __name__ == '__main__':
    main()