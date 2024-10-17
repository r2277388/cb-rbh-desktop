import pandas as pd
import numpy as np
from sqlalchemy import create_engine
from sqlalchemy.exc import SQLAlchemyError  # For error handling
from queries import saldet

def get_connection():
    try:
        engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
        return engine
    except SQLAlchemyError as e:
        print(f"Error connecting to the database: {e}")
        return None

# Create SQL Connection
engine = get_connection()

def upload_saldet(engine) -> pd.DataFrame:
    if engine is None:
        raise ConnectionError("Failed to create a database connection.")
    try:
        with engine.connect() as connection:
            df = pd.read_sql_query(saldet(), connection)
        return df
    except SQLAlchemyError as e:
        print(f"Error executing the SQL query: {e}")
        return pd.DataFrame()  # Return an empty DataFrame in case of error

# Retrieve data and save it as a pickle file
df = upload_saldet(engine)
if not df.empty:
    df.to_pickle('pickled_saldet.pkl')  # Added .pkl extension for clarity
    print("Data saved successfully!")
else:
    print("Data was not saved due to an error or empty DataFrame.")

