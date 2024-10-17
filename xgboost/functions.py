import os
import pandas as pd
from pandas.tseries.offsets import MonthBegin
import numpy as np
from datetime import datetime
import xgboost as xgb

# Default file path for the pickled DataFrame
DEFAULT_PICKLE_PATH = r"E:\My Drive\Colab Notebooks\cb_forecasting\df_pickle.pkl"

def load_dataframe_from_pickle(pickle_path=DEFAULT_PICKLE_PATH):
    """
    Load a DataFrame from a pickle file.

    Parameters:
    pickle_path (str): The path to the pickled file. If not provided, a default path is used.

    Returns:
    pd.DataFrame: The loaded DataFrame.
    """
    if not os.path.exists(pickle_path):
        raise FileNotFoundError(f"The file at {pickle_path} does not exist.")
    
    try:
        df_raw = pd.read_pickle(pickle_path)
        print(f"DataFrame loaded successfully from {pickle_path}")
        return df_raw
    except Exception as e:
        raise RuntimeError(f"An error occurred while loading the DataFrame: {e}")

def clean_raw(df, date_filter=None, pgrp=None, channel=None, ssr=None, flbl=None):
    """
    Clean and filter the raw DataFrame based on given parameters.
    """
    if pgrp:
        df = df[df.pgrp == pgrp]
    if channel:
        df = df[df.channel == channel]
    if ssr:
        df = df[df.ssr == ssr]
    if flbl:
        df = df[df.flbl == flbl]
    
    df = df.groupby('ds')['y'].sum().reset_index()
    df['ds'] = pd.to_datetime(df['ds'])
    df.set_index('ds', inplace=True)
    
    if date_filter:
        df = df[df.index <= pd.to_datetime(date_filter)]
    
    return df.sort_index()

def filter_outliers(df, threshold=7):
    """
    Filter outliers from the DataFrame based on a given threshold.
    """
    mean_y = df['y'].mean()
    std_dev_y = df['y'].std()
    upper_limit = mean_y + threshold * std_dev_y
    return df.query('0 < y <= @upper_limit')

def create_features(df):
    """
    Create time series features based on the DataFrame index.
    """
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
    """
    Add lag features to the DataFrame.
    """
    target_map = df['y'].to_dict()
    df['lag_1y'] = (df.index - pd.Timedelta(days=364)).map(target_map)
    df['lag_2y'] = (df.index - pd.Timedelta(days=728)).map(target_map)
    df['lag_3y'] = (df.index - pd.Timedelta(days=1092)).map(target_map)
    return df

def target_features(df):
    """
    Separate the target variable and features from the DataFrame.
    """
    features = df.columns.tolist()
    features.remove('y')
    return 'y', features

def create_future_with_features(df):
    """
    Create future DataFrame with features and lags, removing holidays.
    """
    current_year = datetime.now().year
    end_date = datetime(current_year + 1, 12, 31)
    date_range = pd.date_range(start=df.index.max() + pd.Timedelta(days=1), end=end_date, freq='B')
    
    future_df = pd.DataFrame(index=date_range)
    future_df['isFuture'] = True
    df['isFuture'] = False

    df_and_future = pd.concat([df, future_df]).pipe(create_features).pipe(add_lags)
    df_and_future = df_and_future[df_and_future.index.year >= current_year].drop(columns=['isFuture'])

    holidays = pd.to_datetime([
        '2023-07-04', '2023-11-23', '2023-11-24', '2023-12-25', '2023-12-26',
        '2024-01-01', '2024-01-15', '2024-02-19', '2024-05-27', '2024-06-19',
        '2024-07-04', '2024-09-02', '2024-11-28', '2024-11-29', '2024-12-25',
        '2024-12-26', '2025-01-01', '2025-01-20', '2025-02-17', '2025-05-26',
        '2025-06-19', '2025-07-04', '2025-09-01', '2025-11-27', '2025-11-28',
        '2025-12-25', '2025-12-26'])
    return df_and_future[~df_and_future.index.isin(holidays)]

def df_acts_fcts(df, future_df, year_starting=2024):
    """
    Combine actuals and predictions DataFrame.
    """
    df_acts = df['y'][df.index.year >= year_starting]
    df_pred = future_df['pred'][future_df.index.year >= year_starting]
    return pd.concat([df_acts, df_pred], axis=1)

def reg_function(df):
    """
    Train an XGBoost regressor on the DataFrame.
    """
    df = create_features(df).pipe(add_lags).pipe(filter_outliers)
    target, features = target_features(df)
    reg = xgb.XGBRegressor(
        base_score=0.2, booster='gbtree', n_estimators=1020,
        objective='reg:squarederror', max_depth=3, eval_metric='rmse',
        learning_rate=0.01, verbosity=1
    )
    reg.fit(df[features], df[target], eval_set=[[df[features], df[target]]], verbose=200)
    return reg

def generate_future_predictions(df, outlier_threshold=7, prediction_year_starting=2024, date_filter=None, pgrp=None, channel=None, ssr=None, flbl=None):
    """
    Generate future predictions for the DataFrame.
    """
    df = clean_raw(df, date_filter, pgrp, channel, ssr, flbl).pipe(create_features).pipe(add_lags)
    reg = reg_function(df)
    _, features = target_features(df)
    future_df = create_future_with_features(df)
    future_df['pred'] = reg.predict(future_df[features])
    return df_acts_fcts(df, future_df, prediction_year_starting)

def multiple_forecast(df, pgrp_list, prediction_year_starting=2019, ssr=None, flbl=None, channel=None):
    """
    Generate multiple forecasts for different pgrp values.
    """
    all_forecasts = pd.DataFrame()
    current_date = datetime.now().date()

    for pgrp in pgrp_list:
        pred_table = generate_future_predictions(df, prediction_year_starting, pgrp=pgrp, ssr=ssr, flbl=flbl, channel=channel)
        pred_table['est'] = pred_table.apply(lambda row: row['y'] if row.name.date() < current_date else row['pred'], axis=1)
        pred_table.drop(columns=['y', 'pred'], inplace=True)
        pred_table['pgrp'] = pgrp
        all_forecasts = pd.concat([all_forecasts, pred_table])

    return all_forecasts

def multiple_forecast_combined(df):
    """
    Generate combined multiple forecasts for 'F' and 'B' labels.
    """
    pgrp_list = df.pgrp.unique()
    fl_version = multiple_forecast(df, pgrp_list, flbl='F')
    bl_version = multiple_forecast(df, pgrp_list, flbl='B')
    fl_version['FL'] = True
    bl_version['FL'] = False
    return pd.concat([fl_version, bl_version])

def reshape_combined_forecasts(df):
    """
    Reshape combined forecasts DataFrame.
    """
    df.drop(columns=['month', 'year', 'Date'], errors='ignore', inplace=True)
    df['est'] = df['est'].round(2)
    df.reset_index(inplace=True)
    df.rename(columns={'index': 'date', 'est': 'fct'}, inplace=True)
    return df

def pgrp_function():
    """
    Load DataFrame and generate multiple forecasts for unique pgrp values.
    """
    df = load_dataframe_from_pickle()
    df['year_month'] = df['ds'].dt.to_period('M')
    pgrp_list = df.pgrp.unique()
    return multiple_forecast(df, pgrp_list)

def check_year_totals(df):
    """
    Check and print yearly totals from the DataFrame.
    """
    df['year'] = df.index.year
    grouped_df = df.groupby('year')['est'].sum()
    print(grouped_df.apply(lambda x: f"{x:,.0f}"))

def pgrp_flbl_function():
    """
    Load DataFrame and generate combined multiple forecasts.
    """
    df = load_dataframe_from_pickle()
    return multiple_forecast_combined(df)

def check_year_totals_flbl(df_flbl):
    """
    Check and print yearly totals for combined forecasts.
    """
    df_flbl['year'] = df_flbl.index.year
    grouped_df_flbl = df_flbl.groupby('year')['est'].sum()
    print(grouped_df_flbl.apply(lambda x: f"{x:,.0f}"))