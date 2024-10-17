from pathlib import Path
import joblib
import pandas as pd
from sqlalchemy import create_engine
from queries.query_saldet import query_saldet

# Define the cache directory and file
CACHE_DIR = Path('E:/My Drive/code/reprint_project/cache_saldet')
CACHE_FILE = CACHE_DIR / 'saldet_cache.pkl'

def get_connection():
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def get_previous_periods(months):
    now = pd.to_datetime('today')
    start_period = (now - pd.DateOffset(months=months)).strftime('%Y%m')
    end_period = now.strftime('%Y%m')
    return start_period, end_period

def upload_saldet(reload=False) -> pd.DataFrame:
    # Load existing cached data if it exists and reload is not requested
    if not reload and CACHE_FILE.exists():
        return joblib.load(CACHE_FILE)

    # Define the period range for the new data (last 3 months)
    start_period, end_period = get_previous_periods(3)

    # Query new data for the last 3 periods
    engine = get_connection()
    with engine.connect() as connection:
        df_new = pd.read_sql_query(query_saldet(start_period, end_period), connection)
    
    # # Ensure the new data only contains the expected columns
    # expected_columns = ['ISBN', 'WeekStartDate', 'qty']
    # df_new = df_new[expected_columns]

    # Save the new data to the cache
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    joblib.dump(df_new, CACHE_FILE)
    return df_new

def main():
    df = upload_saldet()
    df_sorted = df.sort_values(by='WeekStartDate', ascending=False)
    print(df_sorted.info())
    print(df_sorted.head())

if __name__ == '__main__':
    main()