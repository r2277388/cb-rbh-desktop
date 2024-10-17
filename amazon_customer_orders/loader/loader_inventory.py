import pandas as pd
from pathlib import Path
import os
import glob

folder_path = Path(r'G:\SALES\Amazon\RBH\DOWNLOADED_FILES')
file_glob_inventory = r'*inventory*csv'
files = glob.glob(str(folder_path / file_glob_inventory))
file_inventory = max(files, key=os.path.getctime)

# Inventory file to DataFrame
cols_invt = ['ASIN', 'Unfilled Customer Ordered Units', 'Open Purchase Order Quantity', 'Sellable On Hand Inventory', 'Sellable On Hand Units']

def upload_inventory():
    df = pd.read_csv(file_inventory,
                               skiprows=1,
                               na_values='—',
                               usecols=cols_invt
                               )
    
    df['Open Purchase Order Quantity'] = df['Open Purchase Order Quantity'].replace(',','',regex=True).fillna(0).astype(int)
    df['Unfilled Customer Ordered Units'] = df['Unfilled Customer Ordered Units'].replace(',','',regex=True).fillna(0).astype(int)
    df['Sellable On Hand Inventory'] = df['Sellable On Hand Inventory'].replace(r'[$,]','',regex=True).fillna(0).astype(float)
    df['Sellable On Hand Units'] = df['Sellable On Hand Units'].replace(',','',regex=True).fillna(0).astype(int)
    return df

def main():
    df = upload_inventory()
    print(df.info())
    print(df.head())  # Example of what you

if __name__ == '__main__':
        main()
