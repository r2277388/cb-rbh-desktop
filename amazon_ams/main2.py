import pandas as pd
import os
import numpy as np
import time

def load_asin_mapping(file_path):
    df = pd.read_excel(
        file_path,
        usecols=['Asin', 'Isbn13'],
        sheet_name='Sheet1',
        header=0,
        engine='openpyxl'
    )
    df.columns = df.columns.str.lower()
    df.rename(columns={'asin': 'ASIN', 'isbn13': 'ISBN'}, inplace=True)
    return df

def load_monthly_data(tab_info, asin_mapping, period):
    df = pd.read_excel(
        tab_info['file'],
        sheet_name=tab_info['tab'],
        skiprows=tab_info['skiprows'],
        header=0,
        engine='openpyxl'
    )
    df.columns = df.columns.str.strip().str.lower()

    # Sanity check: expected columns
    expected_columns = {'asin', 'clicks', 'impressions', 'units sold', 'spend', '14 day total sales'}
    actual_columns = set(df.columns)

    missing = expected_columns - actual_columns
    extra = actual_columns - expected_columns

    if missing:
        raise ValueError(f"Missing critical columns in {period}: {missing}")
    if extra:
        print(f"ℹ️ Note: Extra columns found in {period}: {extra}")

    df = df.drop(columns=[
        'isbn', 'title', 'pub', 'pub group', 'osd', 'ctr', 'cvr',
        'acos', 'roas', 'product type description'
    ], errors='ignore')
    df.rename(columns={'asin': 'ASIN'}, inplace=True)
    df = df.merge(asin_mapping, on='ASIN', how='left')
    df['ISBN'] = df['ISBN'].fillna(df['ISBN'].str.replace('-', '', regex=False))

    # Derived metrics
    df['CTR'] = df['clicks'].div(df['impressions']).replace([np.inf, -np.inf], np.nan)
    df['CRV'] = df['units sold'].div(df['clicks']).replace([np.inf, -np.inf], np.nan)
    df['ACOS'] = df['spend'].div(df['14 day total sales']).replace([np.inf, -np.inf], np.nan)
    df['ROAS'] = df['14 day total sales'].div(df['spend']).replace([np.inf, -np.inf], np.nan)
    df['period'] = period
    df['source_file'] = tab_info['file']
    return df

def main():
    pd.reset_option('display.max_columns')
    start_time = time.time()

    file_asin_mapping = r"G:\SALES\Amazon\RBH\DOWNLOADED_FILES\Chronicle-AsinMapping.xlsx"

    tab_dict = {
        '2025-01': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025- 01 - January- Performance by ASIN_ALL.xlsx"},
        '2025-02': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025- 02 - February - Performance by ASIN_ALL.xlsx"},
        '2025-03': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025- 03 - March - Performance by ASIN_ALL.xlsx"},
        '2025-04': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 04 - April - Performance by ASIN_ALL.xlsx"},
        '2025-05': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 05 - May - Performance by ASIN_ALL.xlsx"},
    }

    month_list = ['2025-01', '2025-02', '2025-03', '2025-04', '2025-05']

    asin_mapping = load_asin_mapping(file_asin_mapping)
    combined_df = pd.DataFrame()
    errors = []

    for month in month_list:
        if month not in tab_dict:
            warning = f"⚠️ Skipping {month}: not found in tab_dict"
            print(warning)
            errors.append(warning)
            continue

        try:
            print(f"🔄 Processing {month}...")
            df_month = load_monthly_data(tab_dict[month], asin_mapping, month)
            combined_df = pd.concat([combined_df, df_month], ignore_index=True)
        except Exception as e:
            error_msg = f"❌ Failed for {month}: {str(e)}"
            print(error_msg)
            errors.append(error_msg)

    # Save results
    if not combined_df.empty:
        output_file = "combined_amazon_ads_by_asin.csv"
        combined_df.to_csv(output_file, index=False)
        print(f"✅ Combined data saved to: {output_file}")
    else:
        print("❗No data was successfully combined.")

    # Save error log
    if errors:
        with open("processing_errors.log", "w") as f:
            for line in errors:
                f.write(line + "\n")
        print(f"⚠️ Some issues occurred. See processing_errors.log.")

    end_time = time.time()
    print(f"⏱️ Finished in {end_time - start_time:.2f} seconds.")

if __name__ == "__main__":
    main()
