import pandas as pd
import os
import numpy as np

pd.reset_option('display.max_columns')


file_asin_mapping = "G:\SALES\Amazon\RBH\DOWNLOADED_FILES\Chronicle-AsinMapping.xlsx"


folder_path = fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025"
file_jan = fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025- 01 - January- Performance by ASIN_ALL.xlsx"
tab_dict = {'2025-01': {'tab':'USE_main','skiprows':1},
            '2025-02': {'tab':'USE_main','skiprows':1}
            }

df_asin_mapping = pd.read_excel(
    file_asin_mapping,
    usecols = ['Asin', 'Isbn13'],  # specify the columns to read
    sheet_name='Sheet1',
    header=0,    # use the next row as header
    engine='openpyxl'
)

# Lowercase all column names
df_asin_mapping.columns = df_asin_mapping.columns.str.lower()
# Rename to standard names
df_asin_mapping.rename(columns={'asin': 'ASIN', 'isbn13': 'ISBN'}, inplace=True)

df = pd.read_excel(
    file_jan,
    sheet_name=tab_dict['2025-01']['tab'],
    skiprows=tab_dict['2025-01']['skiprows'],      # skip the first row
    header=0,        # use the next row as header
    engine='openpyxl'
)

df.columns = df.columns.str.strip().str.lower()
# Remove unwanted columns
df = df.drop(columns=['isbn', 'title', 'pub', 'pub group',\
    'osd','ctr','cvr','acos','roas', 'product type description'], errors='ignore')
df.rename(columns={'asin': 'ASIN'}, inplace=True)
df = df.merge(df_asin_mapping, on='ASIN', how='left')

df['ISBN'] = df['ISBN'].fillna(df['ISBN'].str.replace('-', '', regex=False))

df['CTR'] = df['clicks'].div(df['impressions']).replace([np.inf, -np.inf], np.nan)
df['CRV'] = df['units sold'].div(df['clicks']).replace([np.inf, -np.inf], np.nan)
df['ACOS'] = df['spend'].div(df['14 day total sales']).replace([np.inf, -np.inf], np.nan)
df['ROAS'] = df['14 day total sales'].div(df['spend']).replace([np.inf, -np.inf], np.nan)

# print(df.head(20))
print(df.columns)