import pandas as pd
import numpy as np
import holidays

def create_features(df):
    """
    Create time series features based on time series index.
    """
    df = df.copy()
    # df['dayofweek_sin'] = np.sin(2 * np.pi * df.index.dayofweek / 7)
    # df['dayofweek_cos'] = np.cos(2 * np.pi * df.index.dayofweek / 7)
    
    df['weekofyear_sin'] = np.sin(2 * np.pi * df.index.isocalendar().week/53)
    df['weekofyear_cos'] = np.cos(2 * np.pi * df.index.isocalendar().week/53)
    
    df['quarter_sin'] = np.sin(2 * np.pi * df.index.quarter / 4)
    df['quarter_cos'] = np.cos(2 * np.pi * df.index.quarter / 4)
    
    df['month_sin'] = np.sin(2 * np.pi * df.index.month / 12)
    df['month_cos'] = np.cos(2 * np.pi * df.index.month / 12)   
    
    df['dayofyear_sin'] = np.sin(2 * np.pi *  df.index.dayofyear/365)
    df['dayofyear_cos'] = np.cos(2 * np.pi *  df.index.dayofyear/365)
    
    df['year'] = df.index.year
    return df

def add_lags(df):
    """
    Add lag features based on weekly aggregation.
    """
    target_map = df['qty'].to_dict()

    # ✅ Use 52 and 104 weeks instead of 364 and 728 days
    df['lag_1y'] = (df.index - pd.Timedelta(weeks=52)).map(target_map)
    df['lag_2y'] = (df.index - pd.Timedelta(weeks=104)).map(target_map)

    return df

def add_holidays(df):
    df = df.copy()

    # ✅ Define US holidays once
    us_holidays = holidays.US(years=df.index.year.unique())

    # ✅ Convert holiday names to datetime
    holiday_dates = {pd.to_datetime(date): name for date, name in us_holidays.items()}

    # ✅ Stores buy before holidays → Shift flags back by 4 weeks
    df['before_christmas'] = (df.index.month == 12) & (df.index.day >= 1)
    df['before_black_friday'] = (df.index.month == 10) & (df.index.isocalendar().week.astype(int) >= 3)
    df['before_cyber_monday'] = (df.index.month == 10) & (df.index.isocalendar().week.astype(int) >= 4)
    
    df['before_valentines'] = (df.index.month == 1) & (df.index.day >= 14)  # Mid-January
    df['before_halloween'] = (df.index.month == 9) & (df.index.isocalendar().week.astype(int) >= 3)
    df['before_july_4th'] = (df.index.month == 6) & (df.index.day >= 1)  # All of June

    # ✅ Use holiday list to get moving dates (shifted backward)
    thanksgiving = [date - pd.DateOffset(weeks=4) for date, name in holiday_dates.items() if "Thanksgiving" in name]
    labor_day = [date - pd.DateOffset(weeks=4) for date, name in holiday_dates.items() if "Labor Day" in name]
    easter = [date - pd.DateOffset(weeks=6) for date, name in holiday_dates.items() if "Easter" in name]

    df['before_thanksgiving'] = df.index.isin(thanksgiving)
    df['before_labor_day'] = df.index.isin(labor_day)
    df['before_easter'] = df.index.isin(easter)

    return df