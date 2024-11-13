# -*- coding: utf-8 -*-
"""
Created on Sat Aug  3 18:31:57 2024

@author: RBH
"""

# %% Imports
import pandas as pd
import time
from tqdm import tqdm
from queries import (query1, query2, query3, query4, query5, query6, query7, query8, query9)
from sqlalchemy import text
from functions import get_connection
from paths import saved_query_location
from variables import get_variables

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
        # Call get_variables() from variables.py
        variables = get_variables(use_current_date=False)
        if not variables:
            print("Failed to retrieve variables.")
            return

        # Unpack variables returned from get_variables()
        prior_day, prior_day_ly, tp, typ1, tply, lyp1 = variables
        
        print('Processing, please wait...')
        print(f'Prior Day: {prior_day}')
        print(f'This Period: {tp}')
        print(f'This Period_LY: {tply}')
        print()

        # Progress bar for running queries
        with tqdm(total=9, desc="Running queries", leave=False, ncols=100) as pbar:
            try:
                df_q1 = pd.read_sql_query(query1(prior_day, tp, tply), engine)
            except Exception as e:
                print(f"Query 1 failed: {e}")
            pbar.update(1)

            try:
                df_q2 = pd.read_sql_query(query2(tp, prior_day), engine)
            except Exception as e:
                print(f"Query 2 failed: {e}")
            pbar.update(1)

            try:
                df_q3 = pd.read_sql_query(query3(tp, prior_day), engine)
                df_q3 = df_q3[df_q3['order'] < 11].iloc[:, 1:-1]
            except Exception as e:
                print(f"Query 3 failed: {e}")
            pbar.update(1)

            try:
                df_q4 = pd.read_sql_query(query4(tp), engine)
                df_q4 = df_q4[df_q4['order'] < 11].iloc[:, 1:-1]
            except Exception as e:
                print(f"Query 4 failed: {e}")
            pbar.update(1)

            try:
                df_q5 = pd.read_sql_query(query5(tp), engine)
            except Exception as e:
                print(f"Query 5 failed: {e}")
            pbar.update(1)

            try:
                df_q6 = pd.read_sql_query(query6(prior_day), engine)
            except Exception as e:
                print(f"Query 6 failed: {e}")
            pbar.update(1)

            try:
                df_q7 = pd.read_sql_query(query7(prior_day), engine)
            except Exception as e:
                print(f"Query 7 failed: {e}")
            pbar.update(1)

            try:
                df_q8 = pd.read_sql_query(query8(prior_day), engine)
            except Exception as e:
                print(f"Query 8 failed: {e}")
            pbar.update(1)

            try:
                df_q9 = pd.read_sql_query(query9(prior_day), engine)
            except Exception as e:
                print(f"Query 9 failed: {e}")
            pbar.update(1)

        # Saving results to an Excel file
        path = saved_query_location()
        
        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
            df_q1.to_excel(writer, sheet_name='df_q1', index=False)
            df_q2.to_excel(writer, sheet_name='df_q2', index=False)
            df_q3.to_excel(writer, sheet_name='df_q3', index=False)
            df_q4.to_excel(writer, sheet_name='df_q4', index=False)
            df_q5.to_excel(writer, sheet_name='df_q5', index=False)
            df_q6.to_excel(writer, sheet_name='df_q6', index=False)
            df_q7.to_excel(writer, sheet_name='df_q7', index=False)
            df_q8.to_excel(writer, sheet_name='df_q8', index=False)
            df_q9.to_excel(writer, sheet_name='df_q9', index=False)

    finally:
        # Ensure the engine connection is closed
        engine.dispose()

    # Output runtime
    end = time.time()
    minutes, seconds = divmod(end - start, 60)
    print(f'Process completed in {int(minutes)} minutes and {int(seconds)} seconds.')
    print(f'The new file is located here: {path}')
    print()

if __name__ == "__main__":
    main()