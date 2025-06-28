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

# Step 1: Load and clean data
df = pd.read_csv("cb_5y_titles.csv")
df.columns = df.columns.str.strip().str.lower()
df.rename(columns={'period': 'Period', 'isbn': 'ISBN', 'pt': 'PT', 'qty': 'Quantity'}, inplace=True)

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
pt_encoded = encoder.fit_transform(df[['PT']])
pt_encoded_df = pd.DataFrame(pt_encoded, columns=encoder.get_feature_names_out(['PT']))

# Merge encoded columns
df = pd.concat([df.reset_index(drop=True), pt_encoded_df], axis=1)

# Step 4: Standardize quantity (target) and prepare features
scaler = StandardScaler()
df['Quantity_scaled'] = scaler.fit_transform(df[['Quantity']])

# Step 5: Train-test split (last row per ISBN is test)
test_df = df.groupby('ISBN').tail(1)
train_df = df.drop(test_df.index)

lag_cols = [f'lag_{l}' for l in lags]
feature_cols = lag_cols + ['month', 'year'] + list(pt_encoded_df.columns)

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

print(f"Mean Absolute Error (MAE): {mae:.2f}")
print(f"Mean Squared Error (MSE): {mse:.2f}")
print(f"Root Mean Squared Error (RMSE): {rmse:.2f}")
print(f"R² Score: {r2:.4f}")

print(results.head())
print(feature_importance)
    