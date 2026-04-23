import sys
from pathlib import Path

import pandas as pd

# Temporary path hack for local script execution
sys.path.append(str(Path(__file__).resolve().parent.parent))

from functions.sql import get_connection
from sql_script import run_sql

engine = get_connection()
sql_code = run_sql()

df = pd.read_sql_query(sql_code, engine)

out_path = Path(__file__).resolve().parent / "data.parquet"
df.to_parquet(out_path, engine="pyarrow", index=False)

print(df.info())
print(df.describe())
print(df.head())
