from sqlalchemy import create_engine
from queries import item_sql 
import pandas as pd

def get_connection() -> create_engine:
    """Establish a connection to the database and return the engine."""
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def fetch_data_from_db(engine: create_engine, query: str) -> pd.DataFrame:
    """Fetch data from the database using the provided query."""
    with engine.connect() as connection:
        df = pd.read_sql_query(query, connection)
    return df

def data_item():
    engine = get_connection()
    query = item_sql()
    df = fetch_data_from_db(engine, query)
    return df

def main():
    engine = get_connection()
    query = item_sql()
    df = fetch_data_from_db(engine, query)
    print(df.head())
    
if __name__ == "__main__":
    main()