import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.append(str(REPO_ROOT))

from functions import get_connection, fetch_data_from_db
from paths import customer_orders_pickle_file
from query_datecheck import check_date
import pandas as pd
from datetime import datetime
import os
from shared.bookscan_calendar import bookscan_week

def data_datecheck():
    engine = get_connection()
    query = check_date()
    df = fetch_data_from_db(engine, query)
    return df

def lastdate_formats():
    df_date = data_datecheck()
    if not df_date.empty:
        last_date = df_date.iloc[0, 0]
        # Convert to datetime if not already
        if not isinstance(last_date, datetime):
            last_date = pd.to_datetime(last_date)
        formatted = last_date.strftime("%m_%d_%Y")
        bookscan = bookscan_week(last_date)
        return formatted, bookscan.week, str(bookscan.year)

def lastdate_display():
    df_date = data_datecheck()
    if not df_date.empty:
        last_date = df_date.iloc[0, 0]
        # Convert to datetime if not already
        if not isinstance(last_date, datetime):
            last_date = pd.to_datetime(last_date)
        formatted = last_date.strftime("%A, %Y-%m-%d")
        days_old = (datetime.now()-last_date).days
        print(f"The last SQL table update was made {days_old} days ago on {formatted}.")
    else:
        print("No dates found in the [WEEK] field ofthe SQL table: [CBQ2].[cb].[Sellthrough_Amazon]")
        
def get_pickle_last_modified(filename):
    if os.path.exists(filename):
        ts = os.path.getmtime(filename)
        return datetime.fromtimestamp(ts)
    else:
        return None
    
def display_pickle_last_modified(filename):
    last_modified = get_pickle_last_modified(filename)
    if last_modified:
        days_old = (datetime.now() - last_modified).days
        formatted = last_modified.strftime("%A, %Y-%m-%d")
        print(f"The data was pickled from the SQL table {days_old} days ago on {formatted}.")
    else:
        print(f"{filename} does not exist.")
        
def main():
    lastdate_display()
    display_pickle_last_modified(customer_orders_pickle_file)

if __name__ == "__main__":
    main()
