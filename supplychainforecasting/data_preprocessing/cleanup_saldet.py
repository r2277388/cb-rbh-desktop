# Getting title metadata from the ebs.item table
import sys
from pathlib import Path
import pandas as pd
import numpy as np
from sqlalchemy import create_engine
from sql_queries.queries import sql_5y_sales

# Add the parent directory to the sys.path so Python can find functions.py
# sys.path.append(str(Path(__file__).resolve().parent.parent))

def get_connection():
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

# Create SQL Connection
engine = get_connection()

def upload_item(engine) -> pd.DataFrame:
    with engine.connect() as connection:
        df = pd.read_sql_query(sql_5y_sales(), connection)
    return df

def get_sales_data(engine, use_cache=True, pkl_path="sales_data.pkl"):
    """
    Loads sales data from pickle if available and use_cache is True,
    otherwise runs the SQL query and saves the result to pickle.
    """
    if use_cache:
        try:
            df = pd.read_pickle(pkl_path)
            print(f"Loaded data from {pkl_path}")
            return df
        except FileNotFoundError:
            print(f"No pickle found at {pkl_path}, running SQL query...")
    # If not using cache or pickle not found, run SQL
    df = upload_item(engine)
    df.to_pickle(pkl_path)
    print(f"Saved data to {pkl_path}")
    return df

if __name__ == "__main__":
    # Set use_cache to True to load from pickle if available, False to always run SQL
    df = get_sales_data(engine, use_cache=False)
    print(df.ISBN.nunique(), "unique ISBNs found")
    print(df.head())
    print(f"Data shape: {df.shape}")
    
### To run this script: python -m data_preprocessing.cleanup_saldet