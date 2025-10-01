from load_weekly_files import get_latest_inventory_csv, extract_week_info
import pandas as pd

def df_inventory():
    inventory_file = get_latest_inventory_csv()
    df = pd.read_csv(inventory_file,
            skiprows=1,
            usecols=['ASIN', 'Product Title', 'Open Purchase Order Quantity'
                    ,'Unfilled Customer Ordered Units','Sellable On Hand Units'])

    df['ASIN'] = df['ASIN'].astype(str).str.strip().str.zfill(10)

    # Convert units columns to numeric
    unit_cols = ['Open Purchase Order Quantity', 'Unfilled Customer Ordered Units', 'Sellable On Hand Units']
    for col in unit_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    return df

def main():
    df = df_inventory()

    inventory_file = get_latest_inventory_csv()
    week_info = extract_week_info(inventory_file)
    print(f"Inventory Week Info: {week_info}")
    print()
    print(df.info())
    print()
    print(df.head())

if __name__ == "__main__":
    main()
