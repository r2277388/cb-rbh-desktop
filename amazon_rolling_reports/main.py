import os
import time
import pandas as pd
from functions import build_column_totals, save_to_excel
from load_amazon_open_po import get_latest_file_date
from load_lastdate import lastdate_formats, lastdate_display,display_pickle_last_modified
from load_amazon_open_po import save_latest_amazon_po_as_pickle
from paths import amazon_po_folder,dp_folders,amazon_rolling_folder
from file_creation import create_rolling_report
from load_rolling_reports import run_query_and_save, sql_co, sql_us

#############
def save_reports_by_pub(df, report_type, week_number, full_year, date_formatted, dp_folders, summary=None, format_cols=None, decimal_cols=None):
    for pub, folder in dp_folders.items():
        df_pub = df[df['Pub'] == pub]
        if not df_pub.empty:
            filename = f"Week {week_number}-{full_year} Rolling Amazon ({date_formatted}) - {report_type}.xlsx"
            filepath = os.path.join(folder, filename)
            os.makedirs(folder, exist_ok=True)
            print(f"{'':#<40}")
            print(f"{pub.center(40)}")
            print(f"{'':#<40}")
            print(f"Saving {report_type} for {pub} to folder: {folder}")
            save_to_excel(df_pub, filepath, summary=summary, format_cols=format_cols, decimal_cols=decimal_cols)
            print(f"Saved {report_type} for {pub} to {filepath}")
            print()
        else:
            print(f"No data for {pub} in {report_type}")
#############

def prompt_update(filename, update_func, *args):
    display_pickle_last_modified(filename)
    choice = input(f"Do you want to update {filename}? (y/n or type 'exit' to quit): ").strip().lower()
    if choice in ['exit', 'quit', '^x']:
        print("Exiting as requested by user.")
        exit(0)
    elif choice in ['y', 'yes']:
        update_func(*args)
        print(f"{filename} updated.\n")
    else:
        print(f"Using existing {filename}.\n")


def main():
    start_time = time.time()

    # --- PO file ---
    pickle_po_file = "latest_amazon_po.pkl"
    print("PO file status:")
    prompt_update(pickle_po_file, save_latest_amazon_po_as_pickle, amazon_po_folder, pickle_po_file)

    # --- Customer Orders ---
    pickle_file1 = "rr_customer_orders.pkl"
    print("Customer Orders file status:")
    prompt_update(pickle_file1, run_query_and_save, sql_co, pickle_file1, "Customer Orders")

    # --- Units Shipped ---
    pickle_file2 = "rr_units_shipped.pkl"
    print("Units Shipped file status:")
    prompt_update(pickle_file2, run_query_and_save, sql_us, pickle_file2, "Units Shipped")

    date_formatted, week_number, full_year = lastdate_formats()

    # Get the latest Amazon PO file and save as pickle
    pickle_po_file = "latest_amazon_po.pkl"
    save_latest_amazon_po_as_pickle(amazon_po_folder, pickle_po_file)

    ###### CUSTOMER ORDERS ############################################

    # Create and save Customer Orders report
    pickle_file1 = "rr_customer_orders.pkl"
    name1 = "Customer Orders"
    df_customer = create_rolling_report(pickle_file1, pickle_po_file)
    sort_col = df_customer.columns[17]
    df_customer = df_customer.sort_values(by=sort_col, ascending=False)
    
    date_cols = [col for col in df_customer.columns if '-' in col and len(col) == 10]
    summary_cols = ['LTD', 'LY_FY', 'TYTD', 'LYTD', 'W52', 'OH', 'PO_Qty'] 
    decimal_cols = ['Price', 'OH_Avg']

    totals_co = build_column_totals(df_customer, date_cols + summary_cols)
    format_cols = date_cols + ['LTD', 'LY_FY', 'TYTD', 'LYTD', 'W52', 'OH', 'PO_Qty']
    
    # Saving to the main folder
    print(fr"Saving {name1} to the main Rolling Reports folder...")
    save_to_excel(
        df_customer,
        os.path.join(amazon_rolling_folder, f"Week {week_number}-{full_year} Rolling Amazon ({date_formatted}) - {name1}.xlsx"),
        summary=totals_co,
        format_cols=format_cols,
        decimal_cols=decimal_cols
    )
    # Saving to the dp folders
    print("Saving to the dp folders...")
    save_reports_by_pub(
    df_customer,
    "Customer Orders",
    week_number,
    full_year,
    date_formatted,
    dp_folders,
    summary=totals_co,
    format_cols=format_cols,
    decimal_cols=decimal_cols
)
    
    ######## UNITS SHIPPED #############################################
    print("#############################################")
    print("Now creating the Units Shipped report...")
    print("#############################################")
    # Create and save Units Shipped report
    pickle_file2 = "rr_units_shipped.pkl"
    name2 = "Units Shipped"
    df_units = create_rolling_report(pickle_file2, pickle_po_file)
    sort_col = df_units.columns[17]
    df_units = df_units.sort_values(by=sort_col, ascending=False)

    date_cols = [col for col in df_units.columns if '-' in col and len(col) == 10]
    summary_cols = ['LTD', 'LY_FY', 'TYTD', 'LYTD', 'W52', 'OH', 'PO_Qty']
    decimal_cols = ['Price', 'OH_Avg']

    totals_us = build_column_totals(df_units, date_cols + summary_cols)
    # Saving to the main folder
    print(fr"Saving {name2} to the main Rolling Reports folder...")
    save_to_excel(
        df_units,
        os.path.join(amazon_rolling_folder, f"Week {week_number}-{full_year} Rolling Amazon ({date_formatted}) - {name2}.xlsx"),
        summary=totals_us,
        format_cols=format_cols,
        decimal_cols=decimal_cols
    )
    # Saving to the dp folders
    print("Saving to the dp folders...")
    save_reports_by_pub(
    df_units,
    "Units Shipped",
    week_number,
    full_year,
    date_formatted,
    dp_folders,
    summary=totals_us,
    format_cols=format_cols,
    decimal_cols=decimal_cols
)

    #################################################################

    end_time = time.time()  # End timer
    elapsed = end_time - start_time
    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)
    print(f'All done! Total runtime: {minutes} minutes, {seconds} seconds.')

if __name__ == "__main__":
    main()