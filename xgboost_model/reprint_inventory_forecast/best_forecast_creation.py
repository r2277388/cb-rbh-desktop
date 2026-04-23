import pandas as pd
import numpy as np
import json

import xgboost as xgb
from sklearn.metrics import mean_squared_error

from load_all import load_all_data
from features import create_features, add_lags, add_holidays

# Obtain the best parameters for each ISBN from the JSON file
with open("best_params_all_isbns.json", "r") as json_file:
    best_params_dict = json.load(json_file)

# Get the ISBNs from the keys of the dictionary
isbn_list = list(best_params_dict.keys())

# Load the data
data = load_all_data(force_refresh=False)
df_saldet = data["saldet"]

#################################################################################
# Initialize future table
last_saturday = df_saldet['Date'].max()

# ✅ Step 2: Generate Next 52 Saturdays
future_saturdays = pd.date_range(start=last_saturday + pd.Timedelta(weeks=1), 
                                 periods=52,  # 1 year of Saturdays
                                 freq='W-SAT') # Only Saturdays

df_future = pd.DataFrame(future_saturdays, columns=['Date'])
df_future = df_future.assign(qty=np.nan) # Add a column for the forecasted quantity
df_future = df_future.sort_values(by="Date").reset_index(drop=True)
#################################################################################

future_predictions = {}

for isbn in isbn_list:
    df_isbn = df_saldet.loc[df_saldet['ISBN']==isbn].copy()
    df_isbn.drop(columns=['ISBN'],inplace=True)
    df_isbn = df_isbn.set_index('Date').sort_index()

    ## Feature Engineering for Past Data
    df_isbn = add_holidays(df_isbn)
    df_isbn = create_features(df_isbn)

    ## Feature Engineering for Future Data
    df_future_isbn = df_future.copy()
    df_future_isbn = df_future_isbn.set_index("Date")
    df_future_isbn = add_holidays(df_future_isbn)
    df_future_isbn = create_features(df_future_isbn)
    
    # Concatenate past and future data
    df_combined = pd.concat([df_isbn, df_future_isbn], axis = 0)
    df_combined = add_lags(df_combined)

    df_train = df_combined.loc[df_combined.index <= last_saturday]
    df_future_isbn = df_combined.loc[df_combined.index > last_saturday]

    TARGET = 'qty'
    FEATURES = [col for col in df_isbn.columns if col != TARGET]

    X_train = df_isbn[FEATURES]
    y_train = df_isbn[TARGET]

    # Get the best parameters for the current ISBN
    best_params = best_params_dict[isbn]

    # Create the model with the best parameters
    model = xgb.XGBRegressor(**best_params, objective='reg:squarederror')
    # Fit the model using all actuals
    model.fit(X_train, y_train)

    # Make predictions using newly trained data on future period
    X_future = df_future_isbn[FEATURES]
    future_pred = model.predict(X_future)
    
    future_predictions[isbn] = future_pred
    
# ✅ Convert Predictions to DataFrame
df_forecast = pd.DataFrame(future_predictions, index=df_future["Date"])

# ✅ Save as Parquet
df_forecast.to_parquet("future_forecasts.parquet")

print("📁 Future forecasts saved as 'future_forecasts.parquet' ✅")




