import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import seaborn as sns
import json

import xgboost as xgb
from sklearn.metrics import mean_squared_error
from sklearn.model_selection import RandomizedSearchCV,TimeSeriesSplit

from load_all import load_all_data
from features import create_features, add_lags,add_holidays

color_pal = sns.color_palette()
plt.style.use('fivethirtyeight')

data = load_all_data(force_refresh=False)

best_params_dict = {} 

df_saldet = data["saldet"]

#####################################
## Filter for the last two years
year_filter=2
x_years_ago = datetime.today() - timedelta(days=year_filter*365)
## this below is making a list of ISBN's that have the largest qty of sales.
# Looking at titles with high sales in the last x years
df_saldet_list = df_saldet.loc[df_saldet['Date'] >= x_years_ago].copy()
# Looking at how many sample are in each ISBN
isbn_counts = df_saldet_list.groupby('ISBN').size()
# ✅ Filter to keep only ISBNs with at least X samples - Need for xgboost splits and testing
valid_isbns = isbn_counts[isbn_counts >= 90].index
df_saldet_list = df_saldet_list[df_saldet_list['ISBN'].isin(valid_isbns)]
# Filtering now for only largest ISBN's
isbn_sales = df_saldet_list.groupby('ISBN')['qty'].sum()
isbn_list = isbn_sales.nlargest(100).index.tolist()
#####################################


for isbn in isbn_list:
    df_isbn = df_saldet.loc[df_saldet['ISBN']==isbn].copy()
    df_isbn.drop(columns=['ISBN'],inplace=True)
    df_isbn = df_isbn.set_index('Date').sort_index()

    ## Feature Engineering
    df_isbn = add_holidays(df_isbn)
    df_isbn = create_features(df_isbn)
    df_isbn = add_lags(df_isbn)

    TARGET = 'qty'
    FEATURES = [col for col in df_isbn.columns if col != TARGET]

    date_split = '2024-07-01'

    X_train = df_isbn.loc[df_isbn.index < date_split, FEATURES]
    y_train = df_isbn.loc[df_isbn.index < date_split, TARGET]

    X_test = df_isbn.loc[df_isbn.index >= date_split, FEATURES]
    y_test = df_isbn.loc[df_isbn.index >= date_split, TARGET]

    param_grid = {
        'max_depth': [3, 5, 7],  # Increase depth
        'learning_rate': [0.01, 0.05, 0.1, 0.2],  # Test lower rates
        'n_estimators': [100, 300, 500],  # Increase number of trees
        'min_child_weight': [1, 3, 5],  # Regularization
        'subsample': [0.6, 0.8, 1.0],  # Control randomness
        'colsample_bytree': [0.6, 0.8, 1.0]  # Feature selection
    }

    # Use TimeSeriesSplit for cross-validation
    tss = TimeSeriesSplit(n_splits=3, test_size=26)  # ~1 year of business days

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
    
    best_params_dict[isbn] = best_params
    
with open ('best_params_all_isbns.json','w') as f:
    json.dump(best_params_dict,f)
    
print('Best parameters for all ISBNs saved to best_params_all_isbns.json')