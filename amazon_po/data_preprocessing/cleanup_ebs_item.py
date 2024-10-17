# Getting title metadata from the ebs.item table
import sys
from pathlib import Path
import pandas as pd
import numpy as np
from sqlalchemy import create_engine
from sql_queries.queries import item_sql

# Add the parent directory to the sys.path so Python can find functions.py
sys.path.append(str(Path(__file__).resolve().parent.parent))

def get_connection():
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

# Create SQL Connection
engine = get_connection()

def upload_item(engine) -> pd.DataFrame:
    with engine.connect() as connection:
        df = pd.read_sql_query(item_sql(), connection)
    return df

def tidyup_item(df: pd.DataFrame) -> pd.DataFrame:
    df['pub'] = df['pub'].astype('datetime64[ns]')
    return df

def get_cleaned_item() -> pd.DataFrame:
    df = upload_item(engine)  # Pass the engine to upload_item
    df = tidyup_item(df)
    return df

# Optional: You can include this to run the script independently if needed
if __name__ == "__main__":
    df = get_cleaned_item()
    print(df.head())