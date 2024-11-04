# -*- coding: utf-8 -*-
"""
Created on Sat Aug  3 18:31:57 2024

@author: RBH
"""

# %% Imports
import pandas as pd
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta
import time
from tqdm import tqdm
from sqlalchemy import text  # Import text
from queries import (query1, query2, query3, query4, query5, query6, query7,
                      query8, query9)
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

    prior_day = prior_day_dt.isoformat()
    prior_day_ly = prior_day_dt - relativedelta(years=1)
    tp = prior_day_dt.strftime('%Y%m')
    tply = prior_day_ly.strftime('%Y%m')

    print('Processing, please wait...')
    print()

    # Progress bar for remaining queries
    with tqdm(total=9, desc="Running queries", leave=False, ncols=100) as pbar:
        df_q1 = pd.read_sql_query(text(query1(tp, tply, prior_day)), engine)
        pbar.update(1)

        df_q2 = pd.read_sql_query(text(query2(tp, prior_day)), engine)
        pbar.update(1)

        df_q3 = pd.read_sql_query(text(query3(tp, prior_day)), engine)
        df_q3 = df_q3[df_q3['order'] < 11].iloc[:, 1:-1]
        pbar.update(1)

        df_q4 = pd.read_sql_query(text(query4(tp)), engine)
        df_q4 = df_q4[df_q4['order'] < 11].iloc[:, 1:-1]
        pbar.update(1)

        df_q5 = pd.read_sql_query(text(query5(tp)), engine)
        pbar.update(1)

        df_q6 = pd.read_sql_query(text(query6(prior_day)), engine)
        pbar.update(1)

        df_q7 = pd.read_sql_query(text(query7(prior_day)), engine)
        pbar.update(1)

        df_q8 = pd.read_sql_query(text(query8(prior_day)), engine)
        pbar.update(1)

        df_q9 = pd.read_sql_query(text(query9(prior_day)), engine)
        pbar.update(1)

    # Saving to file
    path = 'G:\\SALES\\2024 Sales Reports\\SSR\\SSR_Template\\rbh_daily_py.xlsx'

    with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
        df_q1.to_excel(writer, sheet_name='df_q1', index=False)
        # df_q2.to_excel(writer, sheet_name='df_q2', index=False)
        # df_q3.to_excel(writer, sheet_name='df_q3', index=False)
        # df_q4.to_excel(writer, sheet_name='df_q4', index=False)
        # df_q5.to_excel(writer, sheet_name='df_q5', index=False)
        # df_q6.to_excel(writer, sheet_name='df_q6', index=False)
        # df_q7.to_excel(writer, sheet_name='df_q7', index=False)
        # df_q8.to_excel(writer, sheet_name='df_q8', index=False)
        # df_q9.to_excel(writer, sheet_name='df_q9', index=False)

    print(df_q1.head())

    end = time.time()
    total_running_time = end - start
    minutes, seconds = divmod(total_running_time, 60)
    print(f'Process completed in {int(minutes)} minutes and {int(seconds)} seconds.')
    print(f'The new file is located here: {path}')
    print()

if __name__ == "__main__":
    ## main()
    test_connection()