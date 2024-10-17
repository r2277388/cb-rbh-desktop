# main_script.py

#!/usr/bin/env python
# coding: utf-8


import pandas as pd
import time
from tqdm import tqdm

from queries import query_saldet
from functions import (
                        get_connection,
                        load_pickled_data,
                        get_period_info,
                        remove_periods,
                        concatenate_data,
                        check_combination,
                        save_pickle,
                        get_greeting
                        )

def main():
    start_time = time.time()

    # file paths
    file_path = 'E:\\My Drive\\Colab Notebooks\\cb_forecasting\\df_pickle.pkl'
    folder_path = 'E:\\My Drive\\Colab Notebooks\\cb_forecasting\\'
    filename = 'df_pickle.pkl'

    # Get connection to Microsoft SQL Server
    engine = get_connection()
    
    # unpickle saved off saldet
    df_pickled = load_pickled_data(file_path)
    
    # obtain current period and prior period
    current_period, previous_period = get_period_info()
    
    # remove the current and prior periods from the pickeled saldet.
    df_pickled = remove_periods(df_pickled, [current_period, previous_period])
    
    # query new data from >= previous_period
    print()
    print("Running the query...")
    df_additional = pd.read_sql_query(query_saldet(previous_period),engine)

    df_combo = concatenate_data(df_pickled, df_additional)
    df_combo['period'] = df_combo['period'].astype('category')
    df_combo['ssr'] = df_combo['ssr'].astype('category')
    
    check_combination(df_combo, df_pickled, df_additional)
    save_pickle(df_combo, folder_path, filename)

    end_time = time.time()
    total_time = end_time-start_time
    minutes, seconds = divmod(total_time,60)

    print(f"Process completed in {int(minutes)} minutes and {int(seconds)} seconds.")
    print(f"The updated file located: {folder_path}{filename} ")
    print()

if __name__ == "__main__":
    main()