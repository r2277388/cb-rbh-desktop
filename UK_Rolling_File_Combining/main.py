# combine_files.py
import pandas as pd
import load_sales
import load_reserve
import load_midas
from paths import output_path

def combine_data():
    """Combines sales, reserve, and Midas data into a single DataFrame."""
    # Load individual dataframes
    df_sales = load_sales.load_sales_data()
    df_reserve = load_reserve.load_reserve_data()
    df_midas = load_midas.load_midas_data()
    
    # Merge dataframes
    df_combined = pd.merge(df_sales, df_reserve, on='ISBN', how='outer')
    df_combined = pd.merge(df_combined, df_midas, on='ISBN', how='outer')
    df_combined.fillna(0, inplace=True)

    # Rename columns for consistency
    column_dict = {
        'PUB-PRICE': 'Price',
        'DEL-QTY': 'Sales',
        'RESERVED QTY': 'Reserve',
        'Warehouse Stock': 'Avail',
        'Consignment Stock': 'Consignment',
        'All Due Quantity': 'BackOrder'
    }
    df_combined.rename(columns=column_dict, inplace=True)

    return df_combined

def save_to_excel(df_combined):
    """Saves the combined DataFrame to an Excel file."""
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df_combined.to_excel(writer, sheet_name='uk_weekly', index=True)

def main():
    df_combined = combine_data()
    save_to_excel(df_combined)
    print("Data successfully combined and saved to Excel.")

if __name__ == "__main__":
    main()
