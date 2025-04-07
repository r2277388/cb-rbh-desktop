# combine_files.py
import pandas as pd
import load_sales
import load_reserve
import load_midas
from paths import output_path, folder_path
import tkinter as tk
from tkinter import messagebox

def combine_data():
    """Combines sales, reserve, and Midas data into a single DataFrame."""
    # Load individual dataframes
    df_sales = load_sales.load_sales_data()
    df_reserve = load_reserve.load_reserve_data()
    df_midas = load_midas.load_midas_data()
    
    # Merge dataframes
    df_combined = pd.merge(df_sales, df_reserve, on='ISBN', how='outer')
    df_combined = pd.merge(df_combined, df_midas, on='ISBN', how='outer')
    df_combined.fillna(0, inplace=True)

    # Rename columns for consistency
    column_dict = {
        'PUB-PRICE': 'Price',
        'DEL-QTY': 'Sales',
        'RESERVED QTY': 'Reserve',
        'Warehouse Stock': 'Avail',
        'Consignment Stock': 'Consignment',
        'All Due Quantity': 'BackOrder'
    }
    df_combined.rename(columns=column_dict, inplace=True)

    return df_combined

df = combine_data()

df.head(20)