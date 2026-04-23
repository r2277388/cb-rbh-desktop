import os
import pandas as pd
from pathlib import Path


def load_pickled_data(file_path):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    return pd.read_pickle(file_path)

def save_pickle(df, folder_path, filename):
    folder_path = Path(folder_path)
    if not folder_path.exists():
        print(f"The folder '{folder_path}' does not exist. Please check the path.")
    else:
        full_path = folder_path / filename
        df.to_pickle(full_path)
        
def save_parquet(df, folder_path, filename):
    folder_path = Path(folder_path)
    if not folder_path.exists():
        print(f"The folder '{folder_path}' does not exist. Please check the path.")
    else:
        full_path = folder_path / filename
        df.to_parquet(full_path)