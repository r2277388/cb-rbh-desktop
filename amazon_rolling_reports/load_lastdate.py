from functions import get_connection, fetch_data_from_db
from query_datecheck import check_date
import pandas as pd
from datetime import datetime

def data_datecheck():
    engine = get_connection()
    query = check_date()
    df = fetch_data_from_db(engine, query)
    return df

def lastdate_formats():
    df_date = data_datecheck()
    if not df_date.empty:
        last_date = df_date.iloc[0, 0]
        # Convert to datetime if not already
        if not isinstance(last_date, datetime):
            last_date = pd.to_datetime(last_date)
        formatted = last_date.strftime("%m_%d_%Y")
        full_year = last_date.strftime("%Y")
        week_number = last_date.isocalendar()[1]  # ISO week number
        return formatted, week_number, full_year

def lastdate_display():
    df_date = data_datecheck()
    if not df_date.empty:
        last_date = df_date.iloc[0, 0]
        # Convert to datetime if not already
        if not isinstance(last_date, datetime):
            last_date = pd.to_datetime(last_date)
        formatted = last_date.strftime("%A, %Y-%m-%d")
        days_old = (datetime.now()-last_date).days
        print(f"The last SQL table update was made {days_old} days ago on {formatted}.")
    else:
        print("No dates found in the [WEEK] field ofthe SQL table: [CBQ2].[cb].[Sellthrough_Amazon]")
        
def main():
    lastdate_display()

if __name__ == "__main__":
    main()