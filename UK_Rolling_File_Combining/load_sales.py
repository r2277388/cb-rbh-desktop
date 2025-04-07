# load_sales.py
import pandas as pd
import glob
import os
from paths import folder_path

def load_sales_data():
    """Loads and aggregates sales data from multiple text files."""
    files_sales = glob.glob(os.path.join(folder_path, "OPPSSALX*.txt"))
    df_sales_raw = pd.concat(
        (pd.read_csv(f, usecols=['ISBN-13', 'PUB-PRICE', 'DEL-QTY'], encoding='unicode_escape', dtype={'ISBN-13': object}, na_values=0) for f in files_sales),
        axis=0
    )
    df_sales = df_sales_raw[['ISBN-13', 'PUB-PRICE', 'DEL-QTY']].groupby('ISBN-13').agg({
        'PUB-PRICE': 'mean',
        'DEL-QTY': 'sum'
    }).reset_index().set_index('ISBN-13')
    df_sales.index.names = ['ISBN']
    return df_sales

if __name__ == "__main__":
    sales_data = load_sales_data()
    print(sales_data.head(20))
    
    print(sales_data.info())
    