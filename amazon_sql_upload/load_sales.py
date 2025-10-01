from load_weekly_files import get_latest_sales_csv, extract_week_info
import pandas as pd

def df_sales():
    sales_file = get_latest_sales_csv()
    df = pd.read_csv(
        sales_file,
        skiprows=1,
         usecols = ['ASIN', 'Product Title', 'Ordered Revenue'
                   ,'Ordered Units','Shipped Revenue','Shipped Units'
                    ,'Shipped COGS','Customer Returns']
    )

    df['ASIN'] = df['ASIN'].astype(str).str.strip().str.zfill(10)

    # Columns to clean
    money_cols = ['Ordered Revenue', 'Shipped Revenue', \
        'Shipped COGS','Ordered Units','Shipped Units','Customer Returns']
    for col in money_cols:
        df[col] = df[col].astype(str).str.replace('$', '', regex=False).str.replace(',', '', regex=False)
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Convert units columns to numeric
    unit_cols = ['Ordered Units', 'Shipped Units', 'Customer Returns']
    for col in unit_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    return df



def main():
    df = df_sales()

    sales_file = get_latest_sales_csv()
    week_info = extract_week_info(sales_file)
    print(f"Sales Week Info: {week_info}")
    print()
    print(df.info())
    print()
    print(df.head())

    # Show the row where ASIN == '1452111731'
    asin_row = df[df['ASIN'] == '1452111731']
    print("\nRow for ASIN 1452111731:")
    print(asin_row.T)

if __name__ == "__main__":
    main()