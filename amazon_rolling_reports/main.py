import os
import time
import pandas as pd
from functions import save_to_excel
from load_amazon_open_po import get_latest_file_date
from load_lastdate import lastdate_formats, lastdate_display
from load_amazon_open_po import save_latest_amazon_po_as_pickle
from paths import amazon_po_folder,dp_folders,amazon_rolling_folder
from file_creation import create_rolling_report

#############
def save_reports_by_pub(df, report_type, week_number, full_year, date_formatted, dp_folders):
    for pub, folder in dp_folders.items():
        df_pub = df[df['Pub'] == pub]
        if not df_pub.empty:
            filename = f"Week {week_number}-{full_year} Rolling Amazon ({date_formatted}) - {report_type}.xlsx"
            filepath = os.path.join(folder, filename)
            os.makedirs(folder, exist_ok=True)  # <-- Add this line here
            print(f"{'':#<40}")
            print(f"{pub.center(40)}")
            print(f"{'':#<40}")
            print(f"Saving {report_type} for {pub} to folder: {folder}")
            save_to_excel(df_pub, filepath)
            print(f"Saved {report_type} for {pub} to {filepath}")
            print()
        else:
            print(f"No data for {pub} in {report_type}") 
#############


def main():
    start_time = time.time()
    
    #Display last PO file date info
    latest_date = get_latest_file_date(amazon_po_folder)
    if latest_date:
        print("Most recent PO file was updated date:", latest_date.strftime("%A, %Y-%m-%d %H:%M:%S"))
    else:
        print("No .xlsx files found in the folder:", amazon_po_folder)
    
    # Display SQL date last date info
    lastdate_display()
    proceed = input("Do you want to proceed? (y/n): ").strip().lower()
    if proceed not in ['y', 'yes']:
        print("Operation cancelled by user.")
        return

    date_formatted, week_number, full_year = lastdate_formats()

    # Get the latest Amazon PO file and save as pickle
    pickle_po_file = "latest_amazon_po.pkl"
    save_latest_amazon_po_as_pickle(amazon_po_folder, pickle_po_file)

    ###### CUSTOMER ORDERS ############################################

    # Create and save Customer Orders report
    pickle_file1 = "rr_customer_orders.pkl"
    name1 = "Customer Orders"
    df_customer = create_rolling_report(pickle_file1, pickle_po_file)
    # Saving to the main folder
    print(fr"Saving {name1} to the main Rolling Reports folder...")
    save_to_excel(
        df_customer,
        os.path.join(amazon_rolling_folder, f"Week {week_number}-{full_year} Rolling Amazon ({date_formatted}) - {name1}.xlsx")
    )
    # Saving to the dp folders
    print("Saving to the dp folders...")
    save_reports_by_pub(df_customer, "Customer Orders", week_number, full_year, date_formatted, dp_folders)

    ######## UNITS SHIPPED #############################################
    print("#############################################")
    print("Now creating the Units Shipped report...")
    print("#############################################")
    # Create and save Units Shipped report
    pickle_file2 = "rr_units_shipped.pkl"
    name2 = "Units Shipped"
    df_units = create_rolling_report(pickle_file2, pickle_po_file)
    # Saving to the main folder
    print(fr"Saving {name2} to the main Rolling Reports folder...")
    save_to_excel(
        df_units,
        os.path.join(amazon_rolling_folder, f"Week {week_number}-{full_year} Rolling Amazon ({date_formatted}) - {name2}.xlsx")
    )
    # Saving to the dp folders
    print("Saving to the dp folders...")
    save_reports_by_pub(df_units, "Units Shipped", week_number, full_year, date_formatted, dp_folders)

    #################################################################

    end_time = time.time()  # End timer
    elapsed = end_time - start_time
    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)
    print(f'All done! Total runtime: {minutes} minutes, {seconds} seconds.')

if __name__ == "__main__":
    main()