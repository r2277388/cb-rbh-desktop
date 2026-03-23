import glob
import os

import pandas as pd

from paths import DOWNLOADS_FOLDER


folder_path = DOWNLOADS_FOLDER
file_glob_inventory = r"*inventory*csv"
cols_invt = [
    "ASIN",
    "Unfilled Customer Ordered Units",
    "Open Purchase Order Quantity",
    "Sellable On Hand Inventory",
    "Sellable On Hand Units",
]


def upload_inventory():
    files = glob.glob(str(folder_path / file_glob_inventory))
    if not files:
        raise FileNotFoundError(f"No files found in {folder_path} matching {file_glob_inventory}")

    file_inventory = max(files, key=os.path.getctime)
    df = pd.read_csv(
        file_inventory,
        skiprows=1,
        na_values="â€”",
        usecols=cols_invt,
    )

    df["Open Purchase Order Quantity"] = (
        df["Open Purchase Order Quantity"].replace(",", "", regex=True).fillna(0).astype(int)
    )
    df["Unfilled Customer Ordered Units"] = (
        df["Unfilled Customer Ordered Units"].replace(",", "", regex=True).fillna(0).astype(int)
    )
    df["Sellable On Hand Inventory"] = (
        df["Sellable On Hand Inventory"].replace(r"[$,]", "", regex=True).fillna(0).astype(float)
    )
    df["Sellable On Hand Units"] = (
        df["Sellable On Hand Units"].replace(",", "", regex=True).fillna(0).astype(int)
    )
    return df


def main():
    df = upload_inventory()
    print(df.info())
    print(df.head())


if __name__ == "__main__":
    main()
