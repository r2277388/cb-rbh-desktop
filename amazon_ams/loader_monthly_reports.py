import pandas as pd
import numpy as np
from UPDATE_ams_config import tab_dict, month_list

def load_monthly_data(tab_info, asin_mapping, period):
    df = pd.read_excel(
        tab_info['file'],
        sheet_name=tab_info['tab'],
        skiprows=tab_info['skiprows'],
        header=0,
        engine='openpyxl'
    )
    df.columns = df.columns.str.strip().str.lower()

    df = df.drop(columns=[
        'isbn', 'title', 'pub', 'pub group', 'osd', 'ctr', 'cvr',
        'acos', 'roas', 'product type description','campaign'
    ], errors='ignore')

    expected_columns = {'asin', 'clicks', 'impressions', 'units sold', 'spend', '14 day total sales','count of campaigns'}
    actual_columns = set(df.columns)
    missing = expected_columns - actual_columns
    extra = actual_columns - expected_columns

    if missing:
        raise ValueError(f"Missing critical columns in {period}: {missing}")
    if extra:
        print(f"ℹ️ Note: Extra columns found in {period}: {extra}")

    df.rename(columns={'asin': 'ASIN'}, inplace=True)
    df['ASIN'] = df['ASIN'].astype(str).str.strip().str.zfill(10)
    df = df.merge(asin_mapping, on='ASIN', how='left')
    df['ISBN'] = df['ISBN'].fillna(df['ISBN'].str.replace('-', '', regex=False)).infer_objects()

    df['CTR'] = df['clicks'].div(df['impressions']).replace([np.inf, -np.inf], np.nan)
    df['CRV'] = df['units sold'].div(df['clicks']).replace([np.inf, -np.inf], np.nan)
    df['ACOS'] = df['spend'].div(df['14 day total sales']).replace([np.inf, -np.inf], np.nan)
    df['ROAS'] = df['14 day total sales'].div(df['spend']).replace([np.inf, -np.inf], np.nan)

    df['period'] = period
    df['source_file'] = tab_info['file']
    return df
