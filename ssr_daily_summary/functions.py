# functions.py
# This file contains the functions that will be used in the main script
from sqlalchemy import create_engine, text
from datetime import datetime as dt

def get_connection():
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def get_greeting():
    now = dt.now()
    current_hour = now.hour
    current_day = now.strftime('%A')

    if current_hour < 12:
        greeting = "Good morning"
    elif 12 <= current_hour < 17:
        greeting = "Good afternoon"
    else:
        greeting = "Good evening"

    return f"{greeting}, happy {current_day}"

if __name__ == "__main__":
    engine = get_connection()
    try:
        with engine.connect() as conn:
            result = conn.execute(text("SELECT GETDATE()"))  # Use text for the SQL query
            for row in result:
                print(f"Connection successful, current database date/time: {row[0]}")
    except Exception as e:
        print(f"Connection failed: {e}")