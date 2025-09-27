from functions import get_connection, fetch_data_from_db, save_to_pickle
from query_co import sql_co
from query_us import sql_us
import time
import pandas as pd

def run_query_and_save(query_func, filename, label):
    print(f'Loading "{label}" data from the SQL database...')
    print('Querying this data can take up to 4 minutes each ...')
    engine = get_connection()
    query = query_func()
    df = fetch_data_from_db(engine, query)
    print('Query completed ...')
    print(f'{label} shape: {df.shape}')
    save_to_pickle(df, filename)

def main():
    print("Starting the data extraction and saving process...")
    time.sleep(2)
    print("First, we are running the customer orders query ...")
    run_query_and_save(sql_co, "rr_customer_orders.pkl", "Customer Orders")
    time.sleep(2)
    print("Now, we are running the units shipped query ...")
    run_query_and_save(sql_us, "rr_units_shipped.pkl", "Units Shipped")

if __name__ == "__main__":
    main()