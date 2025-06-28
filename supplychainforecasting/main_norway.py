import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import OneHotEncoder
from sklearn.metrics import (
    mean_absolute_error,
    mean_squared_error,
    root_mean_squared_error,
    r2_score
)

# --- STEP 1: Load & format the dataset ---
data = pd.read_csv('norway_new_car_sales_by_make.csv')

# Construct proper datetime Period
data['Period'] = pd.to_datetime(data['Year'].astype(str) + '-' + data['Month'].astype(str).str.zfill(2))

# Keep only needed columns
data = data[['Make', 'Period', 'Quantity']].sort_values(['Make', 'Period'])

# --- STEP 2: Create lag features per Make ---
lags = list(range(1,13))  # Use past n months as predictors
for lag in lags:
    data[f'lag_{lag}'] = data.groupby('Make')['Quantity'].shift(lag)

# Drop rows with NaNs from lag creation
data = data.dropna()

# --- STEP 3: One-hot encode 'Make' ---
encoder = OneHotEncoder(sparse_output=False, drop='first')
make_encoded = encoder.fit_transform(data[['Make']])
make_encoded_df = pd.DataFrame(make_encoded, columns=encoder.get_feature_names_out(['Make']))

# Merge encoded columns into main df
data = pd.concat([data.reset_index(drop=True), make_encoded_df], axis=1)

# --- STEP 4: Train/test split (latest period per Make) ---
test_df = data.groupby('Make').tail(1)
train_df = data.drop(test_df.index)

# --- STEP 5: Prepare features ---
lag_cols = [f'lag_{l}' for l in lags]
feature_cols = lag_cols + list(make_encoded_df.columns)

X_train = train_df[feature_cols]
y_train = train_df['Quantity']

X_test = test_df[feature_cols]
y_test = test_df['Quantity']

# --- STEP 6: Train model ---
model = LinearRegression()
model.fit(X_train, y_train)

# --- STEP 7: Predict and evaluate ---
y_pred = model.predict(X_test)
mae = mean_absolute_error(y_test, y_pred)
mse = mean_squared_error(y_test, y_pred)
rmse = root_mean_squared_error(y_test, y_pred)
r2 = r2_score(y_test, y_pred)

# --- STEP 8: Output results ---
results = test_df[['Make', 'Period']].copy()
results['Actual'] = y_test.values
results['Predicted'] = y_pred

print(results)
print("\n🔍 Model Evaluation Metrics:")
print(f"MAE  (Mean Absolute Error):      {mae:,.0f}")
print(f"MSE  (Mean Squared Error):       {mse:,.0f}")
print(f"RMSE (Root Mean Squared Error):  {rmse:,.0f}")
print(f"R²   (R-squared):                {r2:,.4f}")
