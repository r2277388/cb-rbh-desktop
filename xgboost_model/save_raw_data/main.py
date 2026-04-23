# main_script.py

#!/usr/bin/env python
# coding: utf-8
import sys
from pathlib import Path

# Add the parent directory (code_xgboost) to sys.path
sys.path.append(str(Path(__file__).resolve().parent.parent))

import pandas as pd
import time
from paths import (
    DATAWAREHOUSE_PARQUET_PATH,
    DATAWAREHOUSE_PICKLE_PATH,
    LOCAL_PARQUET_PATH,
    LOCAL_PICKLE_PATH,
    get_pickle_path,
)

from queries import query_saldet
from functions import (
                        get_connection,
                        load_pickled_data,
                        get_period_info,
                        remove_periods,
                        concatenate_data,
                        check_combination,
                        save_pickle,
                        save_parquet
                        )

def main():
    start_time = time.time()

    # Get connection to Microsoft SQL Server
    engine = get_connection()
    
    source_pickle_path = get_pickle_path()
    df_pickled = load_pickled_data(source_pickle_path)
    
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
    
    save_pickle(df_combo, LOCAL_PICKLE_PATH.parent, LOCAL_PICKLE_PATH.name)
    save_pickle(df_combo, DATAWAREHOUSE_PICKLE_PATH.parent, DATAWAREHOUSE_PICKLE_PATH.name)

    save_parquet(df_combo, LOCAL_PARQUET_PATH.parent, LOCAL_PARQUET_PATH.name)
    save_parquet(df_combo, DATAWAREHOUSE_PARQUET_PATH.parent, DATAWAREHOUSE_PARQUET_PATH.name)

    end_time = time.time()
    total_time = end_time-start_time
    minutes, seconds = divmod(total_time,60)

    print(f"Process completed in {int(minutes)} minutes and {int(seconds)} seconds.")
    print(
        "The updated files are saved at:\n"
        f"- Source pickle used: {source_pickle_path}\n"
        f"- Local pickle: {LOCAL_PICKLE_PATH}\n"
        f"- DataWarehouse pickle: {DATAWAREHOUSE_PICKLE_PATH}\n"
        f"- Local parquet: {LOCAL_PARQUET_PATH}\n"
        f"- DataWarehouse parquet: {DATAWAREHOUSE_PARQUET_PATH}"
    )
    print()

if __name__ == "__main__":
    main()
