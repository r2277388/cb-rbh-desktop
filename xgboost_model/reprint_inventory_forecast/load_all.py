import pandas as pd
import numpy as np
from datetime import datetime
import xgboost as xgb
import os

from load_inventory_detail import fetch_inventory_detail
from load_saldet import fetch_sales_detail, get_connection
from load_orders import fetch_hachette_orders

# Define file paths
data_path = "saved_data"
os.makedirs(data_path, exist_ok=True)  # Ensure the directory exists

FILE_PATHS = {
    "saldet": f"{data_path}/df_saldet.parquet",
    "orders": f"{data_path}/df_orders.parquet",
    "inventory": f"{data_path}/df_inventory.parquet",
        }

# Function to load or fetch and save DataFrame
def load_or_fetch(file_path, fetch_function, requires_engine=False,force_refresh=False):
    """Load DataFrame from a file if it exists; otherwise, fetch and save it."""
    if not force_refresh and os.path.exists(file_path):
        print(f"Loading data from {file_path}")
        return pd.read_parquet(file_path)
    
    # Print a message before fetching the data
    print(f"⚡ Currently Fetching {file_path.split('/')[-1].replace('_', ' ').replace('.parquet', '').title()} Data...")
    
    engine = get_connection() if requires_engine else None

    df = fetch_function(engine) if requires_engine else fetch_function()
    
    if df is not None and not df.empty:
        print(f"Saving data to {file_path}")
        df.to_parquet(file_path)
        
    else:
        print(f"Failed to fetch data for {file_path}")
    return df

def load_all_data(force_refresh = False):
    # Load or fetch data
    df_saldet = load_or_fetch(FILE_PATHS["saldet"], fetch_sales_detail, requires_engine=True, force_refresh=force_refresh)
    df_orders = load_or_fetch(FILE_PATHS["orders"], fetch_hachette_orders, requires_engine=True, force_refresh=force_refresh)
    df_inventory = load_or_fetch(FILE_PATHS["inventory"], fetch_inventory_detail, requires_engine=False, force_refresh=force_refresh)

    return {"saldet": df_saldet, "orders": df_orders, "inventory": df_inventory}

if __name__ == "__main__":
    # Load all data
    data = load_all_data(force_refresh=True)
    
    for name,df in data.items():
        if df is not None:
            print(f"{name}:")
            print(df.info())
            print(df.head())
            print("\n")