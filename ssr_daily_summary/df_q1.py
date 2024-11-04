# -*- coding: utf-8 -*-

# %% Imports
import pandas as pd
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta
import time
from queries import query1
from sqlalchemy import text
from functions import get_connection

# %% Main Function

def test_connection():
    engine = get_connection()
    try:
        with engine.connect() as conn:
            result = conn.execute(text("SELECT GETDATE()"))
            for row in result:
                print(f"Connection successful, current database date/time: {row[0]}")
        return True
    except Exception as e:
        print(f"Connection failed: {e}")
        return False

def main():
    start = time.time()
    engine = get_connection()

    try:
        prior_day_str = input('Please enter the prior day (yyyy-mm-dd): ')
        prior_day_dt = dt.strptime(prior_day_str, '%Y-%m-%d')
    except ValueError:
        print("Invalid date format. Please enter the date in yyyy-mm-dd format.")
        return

    # Correct date formatting for the variables
    prior_day = prior_day_dt.strftime('%Y-%m-%d')  # For date comparisons
    prior_day_ly = prior_day_dt - relativedelta(years=1)
    tp = prior_day_dt.strftime('%Y%m')  # Format for 'YYYYMM'
    tply = prior_day_ly.strftime('%Y%m')  # Format for 'YYYYMM'
       
    print('Processing, please wait...')
    print()

    # Generate and print the formatted query for debugging
    formatted_query = query1(prior_day, tp, tply)
    # print("Formatted Query:")
    # print(formatted_query)

    try:
        df_q1 = pd.read_sql_query(text(formatted_query), engine)
        if df_q1.empty:
            print("Query executed but returned an empty DataFrame.")
        else:
            print(df_q1.head())
    except Exception as e:
        print(f"Error while executing the query: {e}")
        return    
     
    # print(df_q1.head())

    end = time.time()
    total_running_time = end - start
    minutes, seconds = divmod(total_running_time, 60)
    print(f'Process completed in {int(minutes)} minutes and {int(seconds)} seconds.')
    print()

if __name__ == "__main__":
    # test_connection()

    main()