from sqlalchemy import create_engine
import pandas as pd

def get_connection() -> create_engine:
    """Establish a connection to the database and return the engine."""
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def fetch_data_from_db(engine: create_engine, query: str) -> pd.DataFrame:
    """Fetch data from the database using the provided query."""
    raw_connection = engine.raw_connection()
    try:
        cursor = raw_connection.cursor()
        try:
            cursor.execute(query)
            columns = None
            rows = None
            while True:
                if cursor.description:
                    columns = [col[0] for col in cursor.description]
                    rows = cursor.fetchall()
                    break
                if not cursor.nextset():
                    raise RuntimeError("Query did not return a result set.")
        finally:
            cursor.close()
    finally:
        raw_connection.close()

    if columns is None or rows is None:
        return pd.DataFrame()

    return pd.DataFrame.from_records(rows, columns=columns)

def save_to_pickle(df, filename):
    df.to_pickle(filename)
    print(f"Pickle File saved to {filename}")
    
def save_to_excel(df,filename):
    with pd.ExcelWriter(filename,engine='xlsxwriter') as writer:
        print('Saving to Excel... this will take a moment...')
        df.to_excel(writer, index=False)
    print(f"Excel saved to: {filename}")