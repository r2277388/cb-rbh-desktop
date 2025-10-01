from paths import amz_catalog
from load_weekly_files import get_latest_catalog_csv, extract_week_info
import pandas as pd

def df_catalog():
    catalog_file = get_latest_catalog_csv()
    df = pd.read_csv(
        catalog_file,
        skiprows=1,
        usecols=['ASIN', 'Product Title', 'EAN', 'ISBN', 'Model Number']
    )

    df['ASIN'] = df['ASIN'].astype(str).str.strip().str.zfill(10)

    rename = {'ISBN':'ISBN_Amz'}
    df = df.rename(columns=rename)

    return df

def main():
    df = df_catalog()
    
    catalog_file = get_latest_catalog_csv()
    week_info = extract_week_info(catalog_file)
    print(f"Catalog Week Info: {week_info}")
    print(df.info())
    print(df.head())

if __name__ == "__main__":
    main()