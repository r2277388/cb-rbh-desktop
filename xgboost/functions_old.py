import pandas as pd
from pandas.tseries.offsets import MonthBegin
import numpy as np
from datetime import datetime
from tabulate import tabulate
import xgboost as xgb
import os

# main file, location where sql data is saved.
file_path = fr"E:\My Drive\Colab Notebooks\cb_forecasting\df_pickle.pkl"

# Retrieve pickled saldet, see file: "E:\My Drive\code\forecast"
def load_dataframe_from_pickle(pickle_path=file_path):
    """
    Load a DataFrame from a pickle file.

    Parameters:
    pickle_path (str): The path to the pickled file. If not provided, a default path is used.

    Returns:
    pd.DataFrame: The loaded DataFrame.
    """
    if os.path.exists(file_path):
        print("File exists.")
    else:
        print("File does not exist.")

    if pickle_path is None:
        pickle_path = os.path.join("E:", "My Drive", "Colab Notebooks", "cb_forecasting", "df_pickle.pkl")
    
    try:
        if not os.path.exists(pickle_path):
            raise FileNotFoundError(f"The file at {pickle_path} does not exist.")
        
        df_raw = pd.read_pickle(pickle_path)
        print(f"DataFrame loaded successfully from {pickle_path}")
        return df_raw
    
    except FileNotFoundError as e:
        print(e)
    except Exception as e:
        print(f"An error occurred while loading the DataFrame: {e}")

# Filters raw file to parameters we want to test for
def clean_raw(df, date_filter=None, pgrp=None, channel=None, ssr=None, flbl=None):
    if pgrp:
        df = df.loc[df.pgrp == pgrp]
    if channel:
        df = df.loc[df.channel == channel]
    if ssr:
        df = df.loc[df.ssr == ssr]
    if flbl:
        df = df.loc[df.flbl == flbl]
    
    df = df.groupby('ds')['y'].sum().reset_index()
    df['ds'] = pd.to_datetime(df['ds'])
    df = df.set_index('ds')
    
    if date_filter:
        df = df[df.index <= pd.to_datetime(date_filter)]
    
    df = df.sort_index()
    return df

def filter_outliers(df, threshold=7):
    mean_y = df['y'].mean()
    std_dev_y = df['y'].std()
    upper_limit = mean_y + threshold * std_dev_y
    filtered_df = df.query('0 < y <= @mean_y + @upper_limit')
    return filtered_df

def create_features(df):
    df = df.copy()
    df['dayofweek'] = df.index.dayofweek
    df['quarter'] = df.index.quarter
    df['month'] = df.index.month
    df['year'] = df.index.year
    df['dayofyear'] = df.index.dayofyear
    df['dayofmonth'] = df.index.day
    df['weekofyear'] = df.index.isocalendar().week
    return df

def add_lags(df):
    target_map = df['y'].to_dict()
    df['lag_1y'] = (df.index - pd.Timedelta('364 days')).map(target_map)
    df['lag_2y'] = (df.index - pd.Timedelta('728 days')).map(target_map)
    df['lag_3y'] = (df.index - pd.Timedelta('1092 days')).map(target_map)
    return df

def target_features(df):
    target = 'y'
    features = df.columns.tolist()
    features.remove('y')
    return target, features

def create_future_with_features(df):
    """
    Takes the existing df, creates a future df thru the end of next year
    and concatenates it to the actual df. Then applies the create_features
    and add_lags functions. It also removes holidays for current and future.
    Returns
    --------
        A df with the actuals data, future dates, and features and lags
    """
    # creating variables
    current_year = datetime.now().year
    end_year = current_year + 1
    end_date = datetime(end_year, 12, 31)
    last_date = df.index.max()
    start_future_date = last_date + pd.Timedelta(days=1)

    # create future df, start date is the next day
    # after last actuals and end date is end of next year
    date_range = pd.date_range(start=start_future_date, end=end_date, freq='B')
    future_df = pd.DataFrame(index=date_range)
    future_df['isFuture'] = True

    # Add column to df for purposes of the concat.
    df['isFuture'] = False

    # combining current actuals data (df) and future df
    # and applies the create_feature() and then the add_lags() function
    df_and_future = pd.concat([df, future_df]).pipe(create_features).pipe(add_lags)
    df_and_future = df_and_future.loc[df_and_future.index.year >= current_year]
    df_and_future = df_and_future.drop(columns=['isFuture'])
    df_and_future = df_and_future.copy()
    dates_to_remove = pd.to_datetime([
        '2023-07-04', '2023-11-23', '2023-11-24', '2023-12-25', '2023-12-26',
        '2024-01-01', '2024-01-15', '2024-02-19', '2024-05-27', '2024-06-19',
        '2024-07-04', '2024-09-02', '2024-11-28', '2024-11-29', '2024-12-25',
        '2024-12-26', '2025-01-01', '2025-01-20', '2025-02-17', '2025-05-26',
        '2025-06-19', '2025-07-04', '2025-09-01', '2025-11-27', '2025-11-28',
        '2025-12-25', '2025-12-26'
        ])
    df_and_future = df_and_future.loc[~df_and_future.index.isin(dates_to_remove)]

    return df_and_future

def df_acts_fcts(df, future_df, year_starting=2024):
    df_acts = df['y']
    df_acts = df_acts.loc[df_acts.index.year >= year_starting]
    df_pred = future_df['pred']
    df_pred = df_pred.loc[df_pred.index.year >= year_starting]
    df_combo = pd.concat([df_acts, df_pred], axis=1)
    return df_combo

def reg_function(df):
    df = create_features(df)
    df = add_lags(df)
    df = filter_outliers(df, threshold=7)
    target, features = target_features(df)
    X_all = df[features]
    y_all = df[target]
    reg = xgb.XGBRegressor(
        base_score=0.2,
        booster='gbtree',
        n_estimators=1020,
        objective='reg:squarederror',
        max_depth=3,
        eval_metric='rmse',
        learning_rate=0.01,
        verbosity=1
    )
    reg.fit(X_all, y_all, eval_set=[[X_all, y_all]], verbose=200)
    return reg

def generate_future_predictions(df, outlier_threshold=7, prediction_year_starting=2024, date_filter=None, pgrp=None, channel=None, ssr=None, flbl=None):
    df = clean_raw(df, date_filter, pgrp, channel, ssr, flbl)
    df = create_features(df)
    df = add_lags(df)
    reg = reg_function(df)
    _, features = target_features(df)
    future_df = create_future_with_features(df, prediction_year_starting)
    future_df['pred'] = reg.predict(future_df[features])
    df_predictions = df_acts_fcts(df, future_df, prediction_year_starting)
    return df_predictions

def multiple_forecast(df, pgrp_list, prediction_year_starting=2019, ssr=None, flbl=None, channel=None):
    all_forecasts = pd.DataFrame()
    current_date = datetime.now().date()
    for pgrp in pgrp_list:
        pred_table = generate_future_predictions(df, prediction_year_starting, flbl=flbl, pgrp=pgrp, ssr=ssr, channel=channel)
        pred_table['est'] = pred_table.apply(lambda row: row['y'] if row.name.date() < current_date else row['pred'], axis=1)
        pred_table = pred_table.drop(columns=['y', 'pred'])
        pred_table['pgrp'] = pgrp
        all_forecasts = pd.concat([all_forecasts, pred_table])
    return all_forecasts

def multiple_forecast_combined(df):
    pgrp_list = df.pgrp.unique()
    fl_version = multiple_forecast(df, pgrp_list, flbl='F')
    bl_version = multiple_forecast(df, pgrp_list, flbl='B')
    fl_version['FL'] = True
    bl_version['FL'] = False
    combined_version = pd.concat([fl_version, bl_version])
    return combined_version

def reshape_combined_forecasts(df):
    if 'month' in df.columns:
        df = df.drop(['month'], axis=1)
    if 'year' in df.columns:
        df = df.drop(['year'], axis=1)
    if 'Date' in df.columns:
        df = df.drop(['Date'], axis=1)
    df['est'] = df['est'].round(2)
    df.reset_index(inplace=True)
    df.rename(columns={'index': 'date', 'est': 'fct'}, inplace=True)
    return df


def pgrp_function():
    df = load_dataframe_from_pickle()
    df['year_month'] = df['ds'].dt.to_period('M')
    pgrp_list = df.pgrp.unique()
    df = multiple_forecast(df, pgrp_list)
    return df

def check_year_totals(df):
    df['year'] = df.index.year
    grouped_df = df.groupby('year')['est'].sum()
    grouped_df = grouped_df.apply(lambda x: f"{x:,.0f}")
    print(grouped_df)

def pgrp_flbl_function():
    df = load_dataframe_from_pickle()
    pgrp_list = df.pgrp.unique()
    df = multiple_forecast_combined(df)
    return df

def check_year_totals_flbl(df_flbl):
    df_flbl['year'] = df_flbl.index.year
    grouped_df_flbl = df_flbl.groupby('year')['est'].sum()
    formatted_grouped_df_flbl = grouped_df_flbl.apply(lambda x: f"{x:,.0f}")
    print(formatted_grouped_df_flbl)