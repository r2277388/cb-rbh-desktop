from paths import amazon_po_folder
import pandas as pd
import glob
import os
from datetime import datetime

def get_latest_file_date(folder, pattern="*.xlsx"):
    files = glob.glob(os.path.join(folder, pattern))
    if files:
        latest_file = max(files, key=os.path.getctime)
        latest_date = os.path.getctime(latest_file)
        # Convert to a readable datetime
        return datetime.fromtimestamp(latest_date)
    else:
        return None

def save_latest_amazon_po_as_pickle(folder, pkl_filename, usecols=['ISBN', 'Accepted Quantity']):
    xlsx_files = glob.glob(os.path.join(folder, "*.xlsx"))
    if xlsx_files:
        latest_file = max(xlsx_files, key=os.path.getctime)
        df = pd.read_excel(latest_file, usecols=usecols, dtype=str)
        df = df.rename(columns ={"Accepted Quantity": "PO_Qty"})
        df.to_pickle(pkl_filename)
        print(f"✅ Saved latest Amazon PO file as pickle: {pkl_filename}")
        return df
    else:
        print(f"No .xlsx files found in the folder: {folder}")
        return None

def main():
    # Example usage of the functions
    latest_date = get_latest_file_date(amazon_po_folder)
    if latest_date:
        print("Most recent PO file was updated date:", latest_date.strftime("%A, %Y-%m-%d %H:%M:%S"))
    else:
        print("No .xlsx files found in the folder:", amazon_po_folder)

    print()
    
    df = save_latest_amazon_po_as_pickle(amazon_po_folder, "latest_amazon_po.pkl")
    if df is not None:
        print(df.head())
        
if __name__ == "__main__":
    main()