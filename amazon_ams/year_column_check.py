import pandas as pd
import os
import numpy as np

pd.reset_option('display.max_columns')


file_asin_mapping = "G:\SALES\Amazon\RBH\DOWNLOADED_FILES\Chronicle-AsinMapping.xlsx"


folder_path = fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025"

tab_dict = {
    '2025-01': {'tab':'USE_main','skiprows':1,'file':fr"G:\SALES\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 01 - January- Performance by ASIN_ALL.xlsx"},
    '2025-02': {'tab':'USE_main','skiprows':1,'file':fr"G:\SALES\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 02 - February - Performance by ASIN_ALL.xlsx"},
    '2025-03': {'tab':'USE_main','skiprows':1,'file':fr"G:\SALES\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 03 - March - Performance by ASIN_ALL.xlsx"},
    '2025-04': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 04 - April - Performance by ASIN_ALL.xlsx"},
    '2025-05': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 05 - May - Performance by ASIN_ALL.xlsx"},
    '2025-06': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 06 - June - Performance by ASIN_ALL.xlsx"},
    '2025-07': {'tab':'USE_main','skiprows':1,'file':fr"G:\SALES\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 07 - July - Performance by ASIN_ALL.xlsx"}
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

month = '2025-07'
month_list = ['2025-01', '2025-02', '2025-03', '2025-04', '2025-05', '2025-06',\
    '2025-07']

df = pd.read_excel(
    tab_dict[month]['file'],
    sheet_name=tab_dict[month]['tab'],
    skiprows=tab_dict[month]['skiprows'],      # skip the first row
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
df['period'] = month

#######################

column_sets = {}

for month in month_list:
    try:
        df = pd.read_excel(
            tab_dict[month]['file'],
            sheet_name=tab_dict[month]['tab'],
            skiprows=tab_dict[month]['skiprows'],
            header=0,
            engine='openpyxl'
        )
        df.columns = df.columns.str.strip().str.lower()
        column_sets[month] = set(df.columns)
    except Exception as e:
        print(f"❌ Error for {month}: {e}")

# Print columns for each month
for month, cols in column_sets.items():
    print(f"{month}: {sorted(cols)}")

# Check if all months have the same columns
all_columns = list(column_sets.values())
if all(all_columns[0] == cols for cols in all_columns):
    print("✅ All months have the same columns.")
else:
    print("❌ Columns differ between months.")
    # Optionally, show which months are different
    from collections import Counter
    col_counter = Counter(tuple(sorted(cols)) for cols in all_columns)
    print("Unique column sets and their counts:", col_counter)
