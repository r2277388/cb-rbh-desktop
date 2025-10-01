import os
import glob
import re
from paths import amz_weekly_sales, amz_weekly_inventory, amz_weekly_traffic, amz_catalog

def get_latest_csv(folder):
    csv_files = glob.glob(os.path.join(folder, "*.csv"))
    if not csv_files:
        print(f"No CSV files found in {folder}")
        return None
    latest_file = max(csv_files, key=os.path.getmtime)
    return latest_file

def extract_week_info(filename):
    # Look for MM-DD-YYYY_MM-DD-YYYY pattern
    match = re.search(r'(\d{1,2}-\d{1,2}-\d{4}_\d{1,2}-\d{1,2}-\d{4})', filename)
    if match:
        return match.group(1)
    else:
        return "No week info found"

def get_latest_sales_csv():
    return get_latest_csv(amz_weekly_sales)

def get_latest_inventory_csv():
    return get_latest_csv(amz_weekly_inventory)

def get_latest_traffic_csv():
    return get_latest_csv(amz_weekly_traffic)

def get_latest_catalog_csv():
    return get_latest_csv(amz_catalog)

def main():
    sales_file = get_latest_sales_csv()
    inventory_file = get_latest_inventory_csv()
    traffic_file = get_latest_traffic_csv()
    catalog_file = get_latest_catalog_csv()

    print("Latest Sales CSV:", sales_file)
    print("  Week info:", extract_week_info(sales_file) if sales_file else "N/A")
    print()
    print("Latest Inventory CSV:", inventory_file)
    print("  Week info:", extract_week_info(inventory_file) if inventory_file else "N/A")
    print()
    print("Latest Traffic CSV:", traffic_file)
    print("  Week info:", extract_week_info(traffic_file) if traffic_file else "N/A")
    print()
    print("Latest Catalog CSV:", catalog_file)
    print("  Week info:", extract_week_info(catalog_file) if catalog_file else "N/A")
    print()
    
if __name__ == "__main__":
    main()