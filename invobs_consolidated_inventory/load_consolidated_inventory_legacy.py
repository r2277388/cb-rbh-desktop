import sys
from pathlib import Path
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import openpyxl
import pandas as pd


REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


def load_data(file_path, sheet_name):
    df = pd.read_excel(
        file_path,
        skiprows=3,
        na_values=[""],
        sheet_name=sheet_name,
        keep_default_na=False,
    )
    df = df.fillna(0)
    df.columns = [
        "Item",
        "pgrp",
        "bkft",
        "title",
        "price",
        "ship",
        "ans",
        "uc_cbc",
        "uc_hbg",
        "uc_cbp",
        "uc_all",
        "val_cbc",
        "units_cbc",
        "val_hbg",
        "units_hbg",
        "val_cbp",
        "units_cbp",
        "total_Units",
        "total_value",
    ]
    return df


def apply_filters(df, pgrp_filter, bkft_filter):
    df = df[df["pgrp"].isin(pgrp_filter)]
    df = df[df["bkft"].isin(bkft_filter)]
    df.reset_index(drop=True, inplace=True)
    return df


def choose_data_sheet(sheet_names):
    preferred_prefix = "all_consolidated_inventories_v_"
    preferred_matches = [
        sheet_name
        for sheet_name in sheet_names
        if sheet_name.lower().startswith(preferred_prefix)
    ]
    if preferred_matches:
        return preferred_matches[0]

    for sheet_name in sheet_names:
        if "consolidated" in sheet_name.lower():
            return sheet_name

    return sheet_names[0] if sheet_names else None


def consolidate_inventory(file_consolidated_inventory=None):
    print(">>> consolidate_inventory_legacy() started")

    if file_consolidated_inventory is None:
        Tk().withdraw()
        print(">>> File dialog opening...")
        file_consolidated_inventory = askopenfilename(
            title="Select the Consolidated Inventory File",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
        )
        print(f">>> File selected: {file_consolidated_inventory}")
    else:
        file_consolidated_inventory = str(Path(file_consolidated_inventory))
        print(f">>> Using provided file: {file_consolidated_inventory}")

    if not file_consolidated_inventory:
        print("No file selected. Exiting.")
        return None

    wb = openpyxl.load_workbook(file_consolidated_inventory, read_only=True)
    sheet_name = choose_data_sheet(wb.sheetnames)
    if sheet_name is None:
        print("No worksheets were found. Exiting.")
        return None

    print(f">>> Sheet selected: {sheet_name}")
    df = load_data(file_consolidated_inventory, sheet_name)

    pgrp_filter = [
        "ART",
        "LIF",
        "CHL",
        "ENT",
        "FWN",
        "PTC",
        "RID",
        "GAM",
        "BAR-LIF",
        "BAR-ENT",
        "BAR-ART",
        "CCB",
        "CPB",
        "CPA",
    ]
    bkft_filter = ["BK", "FT", "CP", "RP"]

    df_filtered = apply_filters(df, pgrp_filter, bkft_filter)
    df_filtered = df_filtered.rename(columns={"Item": "ISBN"})
    df_filtered["ISBN"] = df_filtered["ISBN"].astype(str).str.zfill(13)
    return df_filtered


def main():
    df = consolidate_inventory()
    if df is not None:
        print(df.shape)


if __name__ == "__main__":
    main()
