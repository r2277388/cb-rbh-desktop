# Replacing 'sparse_output' with 'sparse' for compatibility with older scikit-learn versions
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.metrics import (
    mean_absolute_error,
    mean_squared_error,
    r2_score
)
from xgboost import XGBRegressor
import pandas as pd
import numpy as np
import pickle

# Step 1: Load and clean data
df = pd.read_pickle("sales_data.pkl")
df.columns = df.columns.str.strip().str.lower()
df.rename(columns={'period': 'Period', 'isbn': 'ISBN', 'pgrp': 'PGRP', 'qty': 'Quantity'}, inplace=True)

# Convert period to datetime and extract month/year
df['Period'] = pd.to_datetime(df['Period'].astype(str), format='%Y%m')
df['month'] = df['Period'].dt.month
df['year'] = df['Period'].dt.year

# Sort for lag creation
df = df.sort_values(['ISBN', 'Period'])

# Step 2: Create lag features
lags = list(range(1, 13))
for lag in lags:
    df[f'lag_{lag}'] = df.groupby('ISBN')['Quantity'].shift(lag)

# Filter out ISBNs with insufficient history
valid_isbns = df.groupby('ISBN').size()[lambda x: x > 15].index
df = df[df['ISBN'].isin(valid_isbns)].dropna()

# Step 3: Encode categorical variables using older-compatible 'sparse' argument
encoder = OneHotEncoder(drop='first', sparse_output=False)
PGRP_encoded = encoder.fit_transform(df[['PGRP']])
PGRP_encoded_df = pd.DataFrame(PGRP_encoded, columns=encoder.get_feature_names_out(['PGRP']))

# Merge encoded columns
df = pd.concat([df.reset_index(drop=True), PGRP_encoded_df], axis=1)

# Step 4: Standardize quantity (target) and prepare features
scaler = StandardScaler()
df['Quantity_scaled'] = scaler.fit_transform(df[['Quantity']])

# Step 5: Train-test split (last row per ISBN is test)
test_df = df.groupby('ISBN').tail(1)
train_df = df.drop(test_df.index)

lag_cols = [f'lag_{l}' for l in lags]
feature_cols = lag_cols + ['month', 'year'] + list(PGRP_encoded_df.columns)

X_train = train_df[feature_cols]
y_train = train_df['Quantity_scaled']
X_test = test_df[feature_cols]
y_test = test_df['Quantity_scaled']

# Step 6: Train XGBoost model
model = XGBRegressor(n_estimators=100, learning_rate=0.1, random_state=42)
model.fit(X_train, y_train)
y_pred_scaled = model.predict(X_test)
y_pred = scaler.inverse_transform(y_pred_scaled.reshape(-1, 1)).ravel()
y_actual = scaler.inverse_transform(y_test.values.reshape(-1, 1)).ravel()

# Step 7: Evaluate
mae = mean_absolute_error(y_actual, y_pred)
mse = mean_squared_error(y_actual, y_pred)
rmse = np.sqrt(mse)
r2 = r2_score(y_actual, y_pred)

# Step 8: Feature importances
feature_importance = pd.DataFrame({
    'Feature': feature_cols,
    'Importance': model.feature_importances_
}).sort_values(by='Importance', ascending=False)

# Sample predictions
results = test_df[['ISBN', 'Period']].copy()
results['Actual'] = y_actual
results['Predicted'] = y_pred

# import ace_tools as tools; 
# tools.display_dataframe_to_user(name="XGBoost Forecast Results", dataframe=results)
# tools.display_dataframe_to_user(name="Feature Importances", dataframe=feature_importance)
print()
print(df.ISBN.nunique(), "unique ISBNs found")
print()
print(f"Mean Absolute Error (MAE): {mae:,.0f}")
print(f"Mean Squared Error (MSE): {mse:,.0f}")
print(f"Root Mean Squared Error (RMSE): {rmse:,.0f}")
print(f"R² Score: {r2:.4f}")

print(results.head())
print(feature_importance)


import calendar

def forecast_future(df, model, scaler, encoder, lags, months_ahead=12):
    future_results = []
    # Get unique ISBNs and their PGRP
    isbn_groups = df.groupby('ISBN')
    for isbn, group in isbn_groups:
        group = group.sort_values('Period')
        last_row = group.iloc[-1].copy()
        lags_history = list(group['Quantity'].values[-max(lags):])
        pgrp = last_row['PGRP']
        # One-hot encode PGRP for this ISBN
        pgrp_encoded = encoder.transform([[pgrp]])
        # Start forecasting from the month after the last available
        last_period = last_row['Period']
        for i in range(1, months_ahead + 1):
            next_period = last_period + pd.DateOffset(months=1)
            month = next_period.month
            year = next_period.year
            # Prepare lag features
            lag_features = lags_history[-max(lags):][::-1][:max(lags)]
            lag_features = lag_features[::-1]  # oldest to newest
            if len(lag_features) < max(lags):
                lag_features = [0] * (max(lags) - len(lag_features)) + lag_features
            # Build feature vector
            features = lag_features + [month, year] + list(pgrp_encoded[0])
            features = np.array(features).reshape(1, -1)
            # Predict (scaled), then inverse transform
            pred_scaled = model.predict(features)
            pred = scaler.inverse_transform(pred_scaled.reshape(-1, 1)).ravel()[0]
            # Store result
            future_results.append({
                'ISBN': isbn,
                'Period': next_period,
                'Predicted': pred
            })
            # Update lags_history for next prediction
            lags_history.append(pred)
            last_period = next_period
    return pd.DataFrame(future_results)

# Usage:
future_forecast = forecast_future(df, model, scaler, encoder, lags, months_ahead=12)
print(future_forecast.head(20))

# Save to pickle
future_forecast.to_pickle("future_forecast.pkl")
print("Forecast saved to future_forecast.pkl")