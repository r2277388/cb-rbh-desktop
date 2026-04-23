import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from load_all import load_all_data
from functions import to_saturday

# Load the data
data = load_all_data(force_refresh=False)
df_orders= data["orders"]
df_inventory = data["inventory"]
df_forecast = pd.read_parquet("future_forecasts.parquet")

target_isbn = '9781452179612'

## Forecasting for a single ISBN
# df_forecast_isbn = df_forecast[[target_isbn]].copy()
# df_forecast_isbn = df_forecast_isbn.sort_index()
# df_forecast_isbn.rename(columns={target_isbn: 'qty'}, inplace=True)
# print(df_forecast_isbn.head())

## CONVERT ORDER BY SATURDAY DATES
df_orders['EnteredDate'] = df_orders['EnteredDate'].apply(to_saturday)
df_orders['ReleaseDate'] = df_orders['ReleaseDate'].apply(to_saturday)
df_orders['OrderCancelDate'] = df_orders['OrderCancelDate'].apply(to_saturday)

df_orders['ProjectedShip'] = df_orders[df_orders['OrderTypeCode']=='REGULAR']

## Orders for a single ISBN
df_orders = df_orders.loc[df_orders['ISBN']==target_isbn].copy()


print(df_orders.head())