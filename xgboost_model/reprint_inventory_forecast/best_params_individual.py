import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

import xgboost as xgb
from sklearn.metrics import mean_squared_error
from sklearn.model_selection import RandomizedSearchCV,TimeSeriesSplit

from load_all import load_all_data
from features import create_features, add_lags,add_holidays

color_pal = sns.color_palette()
plt.style.use('fivethirtyeight')

data = load_all_data(force_refresh=False)

df_saldet = data["saldet"]
df_orders = data["orders"]
df_inventory = data["inventory"]

df_saldet_wish = df_saldet.loc[df_saldet['ISBN']=='9781452126999'].copy()
df_saldet_wish.drop(columns=['ISBN'],inplace=True)
df_saldet_wish = df_saldet_wish.set_index('Date').sort_index()

# df_saldet_wish.plot(style='.',figsize=(15,5),title = 'Wish you More Sales')
# plt.show()

# train = df_saldet_wish.loc[df_saldet_wish.index < '2024-01-01']
# test = df_saldet_wish.loc[df_saldet_wish.index >= '2024-01-01']

# fig, ax = plt.subplots(figsize=(15, 5))
# train['qty'].plot(ax=ax, label = 'Train')
# test['qty'].plot(ax=ax, label = 'Test')

# ax.axvline('2024-01-01',color='black',linestyle='--',linewidth=1)
# ax.legend(['Train','Test'])
# plt.show()

## Feature Engineering
df_saldet_wish = add_holidays(df_saldet_wish)
df_saldet_wish = create_features(df_saldet_wish)
df_saldet_wish = add_lags(df_saldet_wish)

TARGET = 'qty'
FEATURES = [col for col in df_saldet_wish.columns if col != TARGET]

date_split = '2024-07-01'

X_train = df_saldet_wish.loc[df_saldet_wish.index < date_split, FEATURES]
y_train = df_saldet_wish.loc[df_saldet_wish.index < date_split, TARGET]

X_test = df_saldet_wish.loc[df_saldet_wish.index >= date_split, FEATURES]
y_test = df_saldet_wish.loc[df_saldet_wish.index >= date_split, TARGET]

param_grid = {
    'max_depth': [3, 5, 7],  # Increase depth
    'learning_rate': [0.01, 0.05, 0.1, 0.2],  # Test lower rates
    'n_estimators': [100, 300, 500],  # Increase number of trees
    'min_child_weight': [1, 3, 5],  # Regularization
    'subsample': [0.6, 0.8, 1.0],  # Control randomness
    'colsample_bytree': [0.6, 0.8, 1.0]  # Feature selection
}

# Use TimeSeriesSplit for cross-validation
tss = TimeSeriesSplit(n_splits=4, test_size=26)  # ~1 year of business days

# Create regressor
xgb_model = xgb.XGBRegressor(objective='reg:squarederror',
                             gamma=0)

# Initialize RandomizedSearchCV with TimeSeriesSplit
random_search = RandomizedSearchCV(
    estimator=xgb_model,
    param_distributions=param_grid,
    n_iter=20,  # Number of random parameter sets to try
    scoring='neg_mean_squared_error',
    cv=tss,  # ✅ Using TimeSeriesSplit here
    verbose=2,
    n_jobs=-1
    )

# Add weights
weights = np.where(y_train > y_train.quantile(0.90), 2, 1)  # Give 2x weight to top 10% spikes

# Fit RandomizedSearchCV
random_search.fit(X_train, y_train,sample_weight=weights)

# ✅ Step 1: Retrieve Best Parameters
best_params = random_search.best_params_

# ✅ Step 2: Train Final Model on Full Training Data
final_model = xgb.XGBRegressor(**best_params, objective='reg:squarederror')
final_model.fit(X_train, y_train)

# ✅ Step 3: Predict on Training & Test Set
y_train_pred = final_model.predict(X_train)  # Predictions for train
y_test_pred = final_model.predict(X_test)    # Predictions for test

# ✅ Step 4: Combine Data for Full Visualization
df_results = pd.DataFrame(index=df_saldet_wish.index)
df_results["Actual"] = df_saldet_wish[TARGET]
df_results["Train Predictions"] = np.nan
df_results["Test Predictions"] = np.nan

df_results.loc[X_train.index, "Train Predictions"] = y_train_pred
df_results.loc[X_test.index, "Test Predictions"] = y_test_pred

# ✅ Step 5: Plot the Entire Time Period
plt.figure(figsize=(15, 6))
plt.plot(df_results.index, df_results["Actual"], label="Actual", color="black", linewidth=2)
plt.plot(df_results.index, df_results["Train Predictions"], label="Train Predictions", color="blue", linestyle="dotted")
plt.plot(df_results.index, df_results["Test Predictions"], label="Test Predictions", color="red", linestyle="dashed")

# ✅ Step 6: Mark the Train-Test Split
plt.axvline(X_test.index.min(), color="black", linestyle="--", linewidth=1, label="Train-Test Split")

# ✅ Formatting
plt.title("XGBoost Predictions vs. Actual Sales (Full Time Period)")
plt.xlabel("Date")
plt.ylabel("Quantity Sold")
plt.legend()
plt.grid(True)
plt.show()

# ✅ Compute Forecast Error (Total Products Off)
test_error = (y_test - y_test_pred).sum()  # Total units off
test_abs_error = abs(y_test - y_test_pred).sum()  # Total absolute error (ignoring sign)
test_mae = np.mean(abs(y_test - y_test_pred))  # Mean Absolute Error (per week)

# ✅ Print Metrics
print(f"📉 Total Forecast Error (y_test - y_pred): {test_error:,.0f} units")
print(f"📊 Total Absolute Forecast Error: {test_abs_error:,.0f} units")
print(f"📈 Mean Absolute Error (MAE): {test_mae:,.0f} units per week")
