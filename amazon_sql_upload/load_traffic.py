from load_weekly_files import get_latest_traffic_csv, extract_week_info
import pandas as pd

def df_traffic():
    traffic_file = get_latest_traffic_csv()
    df = pd.read_csv(traffic_file,
                             skiprows=1,
                             usecols=['ASIN', 'Product Title', 'Featured Offer Page Views'])

    df['ASIN'] = df['ASIN'].astype(str).str.strip().str.zfill(10)

    # Convert traffic columns to numeric
    traffic_cols = ['Featured Offer Page Views']
    for col in traffic_cols:
        df[col] = df[col].astype(str).str.replace(',', '', regex=False).str.replace('$', '', regex=False)
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    rename = {
        'Featured Offer Page Views': 'Glance Views'
    }
    df = df.rename(columns=rename)

    return df

def main():
    df = df_traffic()

    traffic_file = get_latest_traffic_csv()
    week_info = extract_week_info(traffic_file)
    print(f"Traffic Week Info: {week_info}")
    print()
    print(df.info())
    print()
    print(df.head())

if __name__ == "__main__":
    main()

 