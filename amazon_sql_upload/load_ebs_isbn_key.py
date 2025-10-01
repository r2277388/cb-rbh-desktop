from sqlalchemy import create_engine
from query_item_key import item_sql
import pandas as pd

def get_connection( ) -> create_engine:
    """Establish a connection to the database and return the engine."""
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def fetch_data_from_db(engine: create_engine, query: str) -> pd.DataFrame:
    """Fetch data from the database using the provided query."""
    with engine.connect() as connection:
        df = pd.read_sql_query(query, connection)
    return df

def isbn_key(query: str = item_sql()):
    engine = get_connection()
    df = fetch_data_from_db(engine, query)
    
    # this is due to the 21 titles in CBQ that don't have 13 digit ISBNs
    df['ISBN'] = df['ISBN'].str.zfill(13)
    
    return df

def main():
    df = isbn_key()
    
    print(df.head())
    
if __name__ == "__main__":
    main()