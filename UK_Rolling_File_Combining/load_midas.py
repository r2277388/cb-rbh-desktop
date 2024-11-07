# load_midas.py
import pandas as pd
from paths import folder_path
import glob
import os

def load_midas_data():
    """Loads and transforms data from the Midas Excel file."""
    files_midas = glob.glob(os.path.join(folder_path, "*.xlsx"))
    df_midas_raw = pd.read_excel(files_midas[0], skiprows=2, engine='openpyxl', na_values=0)
    df_midas = df_midas_raw[['Product', 'Warehouse Stock', 'Consignment Stock', 'All Due Quantity']].copy()
    df_midas['ISBN'] = df_midas['Product'].str[-13:]
    df_midas = df_midas.drop(columns=['Product'])
    df_midas.set_index('ISBN', inplace=True)
    return df_midas
