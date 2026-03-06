# functions.py
# This file contains the functions that will be used in the main script
import sys
from pathlib import Path
from sqlalchemy import text
from datetime import datetime
import os
import pickle
from queries import query_viz_daily
import pandas as pd

# Ensure repo root is importable when this script is executed by file path.
ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from shared.db import get_connection as shared_get_connection


def get_connection():
    return shared_get_connection()

def get_greeting():
    now = datetime.now()
    current_hour = now.hour
    current_day = now.strftime('%A')

    if current_hour < 12:
        greeting = "Good morning"
    elif 12 <= current_hour < 17:
        greeting = "Good afternoon"
    else:
        greeting = "Good evening"

    return f"{greeting}, happy {current_day}"

if __name__ == "__main__":
    engine = get_connection()
    try:
        with engine.connect() as conn:
            result = conn.execute(text("SELECT GETDATE()"))  # Use text for the SQL query
            for row in result:
                print(f"Connection successful, current database date/time: {row[0]}")
    except Exception as e:
        print(f"Connection failed: {e}")
        
def load_data(ty=None,ly=None):
    """Connects to the database and loads the daily data."""
    engine = get_connection()
    # Define the filename with the current date
    today = datetime.today().strftime('%Y-%m-%d')
    filename = f"data_{today}.pkl"

    # Set default years if not provided
    ty = ty or datetime.today().year
    ly = ly or (ty - 1)

    # Check if the file already exists
    if os.path.exists(filename):
        # Load the data from the pickle file
        with open(filename, 'rb') as file:
            df = pickle.load(file)
    else:
        # Load the data from the database
        query_daily = query_viz_daily(ty=ty,ly=ly)
        df = pd.read_sql_query(query_daily, engine)
        # Save the data to a pickle file
        with open(filename, 'wb') as file:
            pickle.dump(df, file)

    return df
