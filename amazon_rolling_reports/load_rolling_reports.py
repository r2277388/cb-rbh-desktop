from sqlalchemy import create_engine
from query_co import sql_co
from query_us import sql_us
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

def data_co():
    engine = get_connection()
    query = sql_co()
    df = fetch_data_from_db(engine, query)
    return df

def data_us():
    engine = get_connection()
    query = sql_us()
    df = fetch_data_from_db(engine, query)
    return df

def save_to_pickle(df, filename):
    df.to_pickle(filename)
    print(f"✅ Data saved to {filename}")

def main():
    df_co = data_co()
    save_to_pickle(df_co, "customer_orders_co.pkl")

    df_us = data_us()
    save_to_pickle(df_us, "customer_orders_us.pkl")

if __name__ == "__main__":
    main()