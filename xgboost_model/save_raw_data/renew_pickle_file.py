import sys
from pathlib import Path

import pandas as pd
from functions import get_connection, save_pickle
from queries import query_saldet

# Add the parent directory (code_xgboost) to sys.path
sys.path.append(str(Path(__file__).resolve().parent.parent))
from paths import DATAWAREHOUSE_PICKLE_PATH, LOCAL_PICKLE_PATH


def query_data(period="201501") -> pd.DataFrame:
    engine = get_connection()
    try:
        return pd.read_sql_query(query_saldet(period), engine)
    except Exception as e:
        print(f"An error occurred querying: {e}")
        return pd.DataFrame()


def main():
    """Main function to query data and save it as a pickle file."""
    df_additional = query_data()
    print(df_additional.info())
    print(df_additional.head())
    # Save to both locations so downstream loaders stay in sync.
    save_pickle(df_additional, LOCAL_PICKLE_PATH.parent, LOCAL_PICKLE_PATH.name)
    save_pickle(
        df_additional, DATAWAREHOUSE_PICKLE_PATH.parent, DATAWAREHOUSE_PICKLE_PATH.name
    )


if __name__ == "__main__":
    main()
