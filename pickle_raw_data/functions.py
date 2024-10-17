# -*- coding: utf-8 -*-
"""
Created on Thu Aug  1 11:07:52 2024

@author: RBH
"""

# functions.py

from sqlalchemy import create_engine
import pandas as pd
import datetime as dt
import os
import getpass

def get_greeting():
    username = getpass.getuser()
    now=dt.datetime.now()
    current_hour = now.hour
    current_day = now.strftime('%A')

    if current_hour < 12:
        greeting = 'Good morning'
    elif 12 <= current_hour < 17:
        greeting = 'Good afternoon'
    else:
        greeting = 'Good evening'
    return f"{greeting} {username}, happy {current_day}"

def load_pickled_data(file_path):
    return pd.read_pickle(file_path)

def filter_last_two_weeks(df):
    df['ds'] = pd.to_datetime(df['ds'])
    latest_date = df['ds'].max()
    two_weeks_ago = latest_date - pd.Timedelta(weeks=2)
    return df[df['ds'] >= two_weeks_ago]

def get_daily_counts(df):
    return df.groupby('ds').size().reset_index(name='record_count')

def get_average_record_count(df):
    return df.record_count.mean()

def get_period_info():
    current_year = dt.datetime.now().year
    current_month = dt.datetime.now().month
    previous_year, previous_month = (current_year - int(current_month == 1), 12 if current_month == 1 else current_month - 1)
    current_period = f"{current_year}{current_month:02d}"
    previous_period = f"{previous_year}{previous_month:02d}"
    return current_period, previous_period

def remove_periods(df, periods):
    return df[~df['period'].isin(periods)]

def get_connection():
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def concatenate_data(df1, df2):
    return pd.concat([df1, df2], ignore_index=True)

def check_combination(df_combo, df1, df2):
    if df_combo.shape[0] == df1.shape[0] + df2.shape[0]:
        max_date = df_combo.ds.max()
        formatted_date = max_date.strftime('%Y-%m-%d')
        print(f'File updated Through: {formatted_date}')
    else:
        print(f'Check that the additional periods were correctly added')

def save_pickle(df, folder_path, filename):
    if not os.path.exists(folder_path):
        print(f"The folder '{folder_path}' does not exist. Please check the path.")
    else:
        full_path = os.path.join(folder_path, filename)
        df.to_pickle(full_path)

def read_sql_query(file_path):
    with open(file_path, 'r') as file:
        return file.read()