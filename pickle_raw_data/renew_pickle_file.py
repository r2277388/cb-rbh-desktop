import pandas as pd

from queries import query_saldet
from functions import get_connection,save_pickle

# file paths
file_path = 'E:\\My Drive\\Colab Notebooks\\cb_forecasting\\df_pickle.pkl'
folder_path = 'E:\\My Drive\\Colab Notebooks\\cb_forecasting\\'
filename = 'df_pickle.pkl'

def query_data(period = '201501')-> pd.DataFrame:
    engine = get_connection()
    try:
        return pd.read_sql_query(query_saldet(period),engine)
    except Exception as e:
        print(f"An error occurred querying: {e}")
        return pd.DataFrame()
        
def main():
    df_additional = query_data()
    print(df_additional.info())
    print(df_additional.head())
    save_pickle(df_additional, folder_path, filename)   
        
if __name__ == "__main__":
    main()
    