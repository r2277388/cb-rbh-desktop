# load_reserve.py
import pandas as pd
import glob
import os
from paths import folder_path

def load_reserve_data():
    """Loads and aggregates reserve data from multiple text files."""
    files_reserve = glob.glob(os.path.join(folder_path, "SMPSTKRES*.txt"))
    df_reserve_raw = pd.concat(
        (pd.read_csv(f, usecols=['ISBN', 'RESERVED QTY'], encoding='unicode_escape', dtype={'ISBN': object}, na_values=0) for f in files_reserve),
        axis=0
    )
    df_reserve = df_reserve_raw[['ISBN', 'RESERVED QTY']].groupby('ISBN').sum().reset_index().set_index('ISBN')
    return df_reserve

if __name__ == "__main__":
    df_reserve = load_reserve_data()
    print(df_reserve.head())
    print(df_reserve.info())
    print(df_reserve.describe())