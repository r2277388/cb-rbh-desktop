import pandas as pd
from pathlib import Path
from datetime import datetime
import shutil
from tkinter import Tk, filedialog

from combined_file import get_merged_files
import excel_report
from paths import process_paths
from data_preprocessing.cleanup_catalog import CAT_FILE_PATH
from data_preprocessing.cleanup_inventorydetail import AVAIL_FILE_PATH
from data_preprocessing.cleanup_po import get_last_po_file_path


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    counter = 1
    while True:
        candidate = path.with_name(f"{stem}_{counter}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def archive_prior_dump_file(path: Path) -> None:
    if not path.exists():
        return None

    archive_dir = process_paths.AMAZON_PO_DATAWAREHOUSE_ARCHIVE_DIR
    archive_dir.mkdir(parents=True, exist_ok=True)
    date_str = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y%m%d")
    archive_path = archive_dir / f"amazon_order_py_dump_{date_str}{path.suffix}"
    archive_path = unique_path(archive_path)
    shutil.copy2(path, archive_path)
    return archive_path


def format_mtime(path: Path) -> str:
    return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %I:%M:%S %p")


def print_input_summary() -> None:
    po_file = get_last_po_file_path()
    print("Amazon (2) PO Report will use these inputs:")
    if po_file and po_file.exists():
        print(f"  PO file:         {po_file}")
        print(f"  Last modified:   {format_mtime(po_file)}")
    else:
        print("  PO file:         selected at runtime")
    print(f"  Catalog file:    {CAT_FILE_PATH}")
    print(f"  Last modified:   {format_mtime(CAT_FILE_PATH)}")
    print(f"  Inventory file:  {AVAIL_FILE_PATH}")
    print(f"  Last modified:   {format_mtime(AVAIL_FILE_PATH)}")
    print("  Item metadata:   sql-2-db / CBQ2 (ebs.Item)")
    print()


def print_output_summary(dated_filename: Path, dump_filename: Path, archive_path: Path | None) -> None:
    print("Amazon (2) PO Report saved outputs here:")
    print(f"  Dated report:    {dated_filename}")
    print(f"  Dump workbook:   {dump_filename}")
    if archive_path is not None:
        print(f"  Dump archive:    {archive_path}")
    else:
        print("  Dump archive:    no prior dump workbook was present")

def publisher_summary(df):
    df = df.groupby('Publisher').agg({'Requested quantity': 'sum',
                                        'Accepted Quantity': 'sum', 
                                        'Total accepted cost': 'sum',
                                        'Lost Sales': 'sum',
                                        'Requested Cost': 'sum',
                                       })
    df.sort_values(by='Total accepted cost', ascending=False, inplace=True)
    
    return df

def ordered_summary_cb(df):
    df_summary = df[df['Publisher'] == 'Chronicle'].copy()
    df_summary.sort_values(by='Requested quantity', ascending=False, inplace=True)
    return df_summary.head(20)

def ordered_summary_ga(df):
    df_summary = df[df['Publisher'] == 'Galison'].copy()
    df_summary.sort_values(by='Requested quantity', ascending=False, inplace=True)
    return df_summary.head(20)

def ordered_summary_dp(df):
    df_summary = df[~df['Publisher'].isin(['Chronicle', 'Galison'])].copy()
    df_summary.sort_values(by='Requested quantity', ascending=False, inplace=True)
    return df_summary.head(20)

def accepted_summary_cb(df):
    df_summary = df[df['Publisher'] == 'Chronicle'].copy()
    df_summary.sort_values(by='Total accepted cost', ascending=False, inplace=True)
    return df_summary.head(20)

def accepted_summary_ga(df):
    df_summary = df[df['Publisher'] == 'Galison'].copy()
    df_summary.sort_values(by='Total accepted cost', ascending=False, inplace=True)
    return df_summary.head(20)

def accepted_summary_dp(df):
    df_summary = df[~df['Publisher'].isin(['Chronicle', 'Galison'])].copy()
    df_summary.sort_values(by='Total accepted cost', ascending=False, inplace=True)
    return df_summary.head(20)

def lost_sales_summary(df):
    df_summary = df[~df['Publisher'].isin(['Chronicle', 'Galison'])].copy()
    df_summary.sort_values(by='Lost Sales', ascending=False, inplace=True)
    return df_summary.head(20)

def main():
    # Load/clean/merge once so the PO file picker only appears a single time.
    df = get_merged_files()
    print()
    print_input_summary()

    # Reuse the same merged data for the dated Excel report.
    agg_pub, agg_pgrp = excel_report.aggregate_data(df)
    po_year_folder = Path(r'G:\SALES\Amazon\PURCHASE ORDERS') / str(datetime.now().year)
    dated_filename = excel_report.generate_filename(po_year_folder)
    excel_report.save_to_excel(df, agg_pub, agg_pgrp, dated_filename)
    
    filename = process_paths.AMAZON_PO_DATAWAREHOUSE_DUMP_FILE
    filename.parent.mkdir(parents=True, exist_ok=True)
    archive_path = archive_prior_dump_file(filename)
    with pd.ExcelWriter(filename) as writer:
        publisher_summary(df).to_excel(writer, sheet_name='pub_summary')
        ordered_summary_cb(df).to_excel(writer, sheet_name='ordered_summary_cb',index=False)
        ordered_summary_ga(df).to_excel(writer, sheet_name='ordered_summary_ga',index=False)
        ordered_summary_dp(df).to_excel(writer, sheet_name='ordered_summary_dp',index=False)
        accepted_summary_cb(df).to_excel(writer, sheet_name='accepted_summary_cb',index=False)
        accepted_summary_ga(df).to_excel(writer, sheet_name='accepted_summary_ga',index=False)
        accepted_summary_dp(df).to_excel(writer, sheet_name='accepted_summary_dp',index=False)
        lost_sales_summary(df).to_excel(writer, sheet_name='lost_sales_summary',index=False)
    print()
    print_output_summary(dated_filename, filename, archive_path)

if __name__ == "__main__":
    main()
