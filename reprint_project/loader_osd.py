import pandas as pd
from pathlib import Path
import joblib
from sqlalchemy import create_engine
from queries.query_osd import query_osd

# Define a separate cache directory and file for OSD data
CACHE_DIR_OSD = Path('E:/My Drive/code/reprint_project/cache')
CACHE_FILE_OSD = CACHE_DIR_OSD / 'osd_cache.pkl'

def get_connection():
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def upload_osd(reload=False) -> pd.DataFrame:
    # Load existing cached data if it exists and reload is not requested
    if not reload and CACHE_FILE_OSD.exists():
        return joblib.load(CACHE_FILE_OSD)

    # Query new data
    engine = get_connection()
    with engine.connect() as connection:
        df_new = pd.read_sql_query(query_osd(), connection)

    # Save the new data to the cache
    CACHE_DIR_OSD.mkdir(parents=True, exist_ok=True)
    joblib.dump(df_new, CACHE_FILE_OSD)
    return df_new

def main():
    df = upload_osd()
    print(df.info())
    print(df.head())

if __name__ == '__main__':
    main()