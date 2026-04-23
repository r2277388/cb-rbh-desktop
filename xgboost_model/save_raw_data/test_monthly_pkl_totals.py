import sys
from pathlib import Path
import pandas as pd

# Add the parent directory (code_xgboost) to sys.path
sys.path.append(str(Path(__file__).resolve().parent.parent))

from paths import DATAWAREHOUSE_PICKLE_PATH, LOCAL_PICKLE_PATH
from functions import load_pickled_data

path = 'df_pickle.pkl'

# Check if the file exists
if not Path(path).exists():
    raise FileNotFoundError(f"The file {path} does not exist.")

# unpickle saved off saldet
df_pickled = pd.read_pickle(path)

print(df_pickled.info())

# Add a year and month column for grouping
df_pickled['year'] = df_pickled['ds'].dt.year
df_pickled['month'] = df_pickled['ds'].dt.month

# Total y by year
total_by_year = df_pickled.groupby('year')['y'].sum()
print("Total y by year:")
print(total_by_year.apply(lambda x: f"{int(round(x)):,}"))


# YTD (Jan-Jun) by year
ytd = df_pickled[df_pickled['month'].between(1, 9)].groupby('year')['y'].sum()
print("\nYTD (Jan-Sep) y by year:")
print(ytd.apply(lambda x: f"{int(round(x)):,}"))
