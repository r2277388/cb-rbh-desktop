import getpass
import importlib.util
import os
import subprocess
from datetime import datetime
from pathlib import Path

# call the PO archive manager directly
from paths import process_paths
import tools.po_archive_manager as po_archive_manager


def get_full_name():
    USER_NAMES = {
        "kbs": "Kate Breiting Schmitz",
        "mjk": "Marlena Kwasnik",
        "sdm": "Sam Mariucci",
        "RBH": "Barrett Hooper",
    }
    username = getpass.getuser()
    return USER_NAMES.get(username, username)


def greet_user():
    current_datetime = datetime.now()
    current_hour = current_datetime.hour
    current_day = current_datetime.strftime("%A")

    if current_hour < 12:
        greeting = "Good morning"
    elif 12 <= current_hour < 17:
        greeting = "Good afternoon"
    else:
        greeting = "Good evening"

    full_name = get_full_name()
    return f"\n{greeting}, {full_name}! Happy {current_day}!"


def get_farewell_message():
    current_hour = datetime.now().hour
    full_name = get_full_name()

    if 12 <= current_hour < 17:
        return f"\nHave a great afternoon, {full_name}!"
    elif 17 <= current_hour < 24:
        return f"\nGood evening, {full_name}!"
    else:
        return f"\nHave a great day, {full_name}!"


def display_options():
    options = [
        "01. Amazon PO Archive Manager",
        "02. Amazon PO Report",
        "03. Amazon PreOrders",
        "04. Amazon Customer Orders",
        "05. Amazon Sellthru SQL Upload",
        "06. Amazon Rolling Reports",
        "07. Amazon AMS Manager",
        "08. Barnes & Noble Rolling Reports",
        "09. Frontlist Supercharged Data",
        "10. SSR Daily Summary",
        "11. UK Rolling File Combining",
        "12. Hachette Orders - Shipping Estimates",
        "13. Consolidate Inventory for the INVOBS",
        "14. XGBoost Model",
        "15. Check Table Updates",
        "16. Desk Procedures",
        "17. Exit",
    ]
    print("\nWhat would you like to run?")
    print()
    for option in options:
        print(f"    {option}")


def display_info(choice):
    info = {
        "1": "PO Archive Manager: Launches the PO archive helper to archive prior current_amaz_preorders and copy the new file into po_analysis.",
        "2": f"""Amazon PO Report: Generates a detailed report based on Amazon Purchase Orders.
        Before running, save the Vendor Central PO File to:
        {process_paths.AMAZON_PO_ANALYSIS_INPUT_FILE}
        A PO Report is saved off to: {process_paths.AMAZON_PO_ROOT_FOLDER} folder""",
        "3": "Amazon NYP PreOrders: Generates a report for Amazon NYP PreOrders. Save the relevant data file to the appropriate location before running.",
        "4": "Amazon Customer Orders: Generates a report for Amazon Customer Orders. Save the relevant data file to the appropriate location before running.",
        "5": "amazon_sql_upload: Runs the amazon_sql_upload workflow (ASIN/ISBN conversion, uploads, etc.).",
        "6": "Amazon Rolling Reports: Runs a 10-week SQL freshness check first, then asks whether to continue with the full process.",
        "7": "Amazon AMS Manager: Manage/update AMS month configuration and run incremental or full AMS processing.",
        "8": "Barnes & Noble Rolling Reports: Builds weekly Barnes & Noble rolling-report source files, starting with the combined POS non-book extract.",
        "9": "Frontlist Supercharged Data: Builds the frontlist ISBN master file by merging Frontlist Tracking with cached Excel extracts and SQL source data.",
        "10": "SSR Daily Summary: Prepares the data for the SSR Daily Summary email.",
        "11": "UK Rolling File Combining: This combines the sales, reserve and midas files together.",
        "12": "Hachette Orders - Shipping Estimates: Generates a report for Hachette Orders.",
        "13": """Consolidate Inventory for the INVOBS: Runs the Consolidated Inventory program for INVOBS.
        This program takes the consolidated inventory data from Oracle, run by Ailing, and
        explodes out the CDU's into their components to give a component-only inventory file.""",
        "14": "XGBoost Model: Launches the xgboost_model workflow menu.",
        "15": "Check Table Updates: Runs SQL checks for table freshness and recent weeks for SSR/Amazon/Bookscan tables.",
        "16": "Desk Procedures: Opens a menu of desk procedures and run instructions.",
        "17": "Exit: Exits the program.",
    }
    return info.get(choice, "Invalid choice. No information available.")


def get_latest_matching_file(folder_path: str | Path, pattern: str) -> Path:
    folder = Path(folder_path)
    matches = list(folder.glob(pattern))
    if not matches:
        raise FileNotFoundError(f"No files found in {folder} with pattern {pattern}")
    return max(matches, key=os.path.getctime)


def confirm_amazon_preorders_files() -> bool:
    catalog_file = get_latest_matching_file(process_paths.DOWNLOADS_FOLDER, "*Catalog*csv")
    inventory_file = get_latest_matching_file(process_paths.DOWNLOADS_FOLDER, "*Inventory*csv")

    while True:
        print()
        print("Amazon PreOrders will use these files:")
        print(f"  Catalog:   {catalog_file}")
        print(f"  Inventory: {inventory_file}")
        print()
        print("    1. Continue")
        print("    2. Return to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def confirm_amazon_customer_orders_files() -> bool:
    script_path = process_paths.AMAZON_CUSTOMER_ORDERS_SCRIPT
    weekly_sales_file = get_latest_matching_file(process_paths.DOWNLOADS_FOLDER, "*Sales*Weekly*csv")
    catalog_file = get_latest_matching_file(process_paths.DOWNLOADS_FOLDER, "*Catalog*csv")
    traffic_file = get_latest_matching_file(process_paths.DOWNLOADS_FOLDER, "*Traffic*csv")
    ypticod_file = process_paths.ORACLE_YPTICOD_FILE
    output_file = process_paths.AMAZON_CUSTOMER_ORDERS_OUTPUT_FILE

    if not ypticod_file.exists():
        raise FileNotFoundError(f"Required file not found: {ypticod_file}")

    while True:
        print()
        print("Amazon Customer Orders will use these files:")
        print(f"  Script:            {script_path}")
        print(f"  Weekly Sales:      {weekly_sales_file}")
        print(f"  Catalog:           {catalog_file}")
        print(f"  Traffic:           {traffic_file}")
        print(f"  Oracle YPTICOD:    {ypticod_file}")
        print(f"  Output workbook:   {output_file}")
        print()
        print("    1. Continue")
        print("    2. Return to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def confirm_amazon_sql_upload_files() -> bool:
    script_path = process_paths.AMAZON_SQL_UPLOAD_SCRIPT
    manual_key_path = process_paths.AMAZON_SQL_UPLOAD_MANUAL_KEY_FILE
    removal_list_path = process_paths.AMAZON_SQL_UPLOAD_REMOVAL_LIST_FILE

    sales_file = get_latest_matching_file(
        process_paths.AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["sales"], "*.csv"
    )
    inventory_file = get_latest_matching_file(
        process_paths.AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["inventory"], "*.csv"
    )
    traffic_file = get_latest_matching_file(
        process_paths.AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["traffic"], "*.csv"
    )
    catalog_file = get_latest_matching_file(
        process_paths.AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["catalog"], "*.csv"
    )
    ypticod_file = process_paths.ORACLE_YPTICOD_FILE

    if not ypticod_file.exists():
        raise FileNotFoundError(f"Required file not found: {ypticod_file}")

    while True:
        print()
        print("amazon_sql_upload will use these files:")
        print(f"  Script:            {script_path}")
        print(f"  Sales CSV:         {sales_file}")
        print(f"  Inventory CSV:     {inventory_file}")
        print(f"  Traffic CSV:       {traffic_file}")
        print(f"  Catalog CSV:       {catalog_file}")
        print(f"  Oracle YPTICOD:    {ypticod_file}")
        print(f"  Manual key file:   {manual_key_path}")
        print(f"  Removal list file: {removal_list_path}")
        print("  SQL source:        sql-2-db / CBQ2 (EBS item ISBN key)")
        print("  Output workbook:   Chosen in save dialog after the process starts")
        print()
        print("    1. Continue")
        print("    2. Return to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def confirm_amazon_rolling_reports_check_files() -> bool:
    script_path = process_paths.AMAZON_ROLLING_CHECK_SCRIPT
    sql_file = process_paths.AMAZON_ROLLING_SQL_FILE

    while True:
        print()
        print("Amazon Rolling Reports check will use these files:")
        print(f"  Script:            {script_path}")
        print(f"  SQL file:          {sql_file}")
        print("  SQL source:        sql-2-db / CBQ2")
        print("  Table checked:     [CBQ2].[cb].[Sellthrough_Amazon]")
        print()
        print("    1. Continue")
        print("    2. Return to previous menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def confirm_amazon_rolling_reports_run_files() -> bool:
    script_path = process_paths.AMAZON_ROLLING_REPORTS_SCRIPT
    latest_po_file = get_latest_matching_file(process_paths.AMAZON_PO_FOLDER, "*.xlsx")
    customer_orders_pickle = process_paths.AMAZON_ROLLING_CUSTOMER_ORDERS_PICKLE
    units_shipped_pickle = process_paths.AMAZON_ROLLING_UNITS_SHIPPED_PICKLE
    po_pickle = process_paths.AMAZON_ROLLING_PO_PICKLE
    customer_orders_sql = process_paths.AMAZON_ROLLING_CUSTOMER_ORDERS_QUERY
    units_shipped_sql = process_paths.AMAZON_ROLLING_UNITS_SHIPPED_QUERY
    date_check_sql = process_paths.AMAZON_ROLLING_DATE_CHECK_QUERY
    output_folder = process_paths.AMAZON_ROLLING_OUTPUT_FOLDER

    while True:
        print()
        print("Amazon Rolling Reports will use these files:")
        print(f"  Script:                  {script_path}")
        print(f"  Latest Amazon PO file:   {latest_po_file}")
        print(f"  PO pickle:               {po_pickle}")
        print(f"  Customer Orders pickle:  {customer_orders_pickle}")
        print(f"  Units Shipped pickle:    {units_shipped_pickle}")
        print(f"  Customer Orders query:   {customer_orders_sql}")
        print(f"  Units Shipped query:     {units_shipped_sql}")
        print(f"  Date check query:        {date_check_sql}")
        print("  SQL source:              sql-2-db / CBQ2")
        print(f"  Main output folder:      {output_folder}")
        print("  DP output folders:       Configured in amazon_rolling_reports/paths.py")
        print()
        print("    1. Continue")
        print("    2. Return to previous menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def load_module_from_path(module_name: str, file_path: Path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load module from {file_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def confirm_amazon_ams_files() -> bool:
    manager_script = process_paths.AMAZON_AMS_MANAGER_SCRIPT
    process_script = process_paths.AMAZON_AMS_PROCESS_SCRIPT
    config_path = process_paths.AMAZON_AMS_CONFIG_FILE
    mapping_file = process_paths.CHRONICLE_ASIN_MAPPING_FILE
    output_pickle = process_paths.AMAZON_AMS_OUTPUT_PICKLE
    output_excel = process_paths.AMAZON_AMS_OUTPUT_EXCEL
    error_log = process_paths.AMAZON_AMS_ERROR_LOG

    config_module = load_module_from_path("amazon_ams_config_preview", config_path)
    tab_dict = getattr(config_module, "tab_dict", {})
    month_list = sorted(getattr(config_module, "month_list", []))

    latest_month = month_list[-1] if month_list else None
    latest_month_file = None
    latest_month_tab = None
    if latest_month and latest_month in tab_dict:
        latest_month_file = tab_dict[latest_month].get("file")
        latest_month_tab = tab_dict[latest_month].get("tab")

    while True:
        print()
        print("Amazon AMS Manager will use these files:")
        print(f"  Manager script:         {manager_script}")
        print(f"  Full-process script:    {process_script}")
        print(f"  Config file:            {config_path}")
        print(f"  ASIN mapping file:      {mapping_file}")
        if latest_month:
            print(f"  Latest configured month:{latest_month}")
            print(f"  Latest month file:      {latest_month_file}")
            print(f"  Latest month tab:       {latest_month_tab}")
        else:
            print("  Latest configured month: none found in config")
        print("  SQL source:             sql-2-db / CBQ2 (item metadata)")
        print(f"  Output pickle:          {output_pickle}")
        print(f"  Output workbook:        {output_excel}")
        print(f"  Error log:              {error_log}")
        print()
        print("    1. Continue")
        print("    2. Return to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def confirm_frontlist_supercharged_files() -> bool:
    script_path = process_paths.FRONTLIST_SUPERCHARGED_SCRIPT
    output_dir = process_paths.FRONTLIST_SUPERCHARGED_OUTPUT_DIR

    frontlist_file = get_latest_matching_file(process_paths.FRONTLIST_TRACKING_FOLDER, "*.xlsx")
    ingram_file = get_latest_matching_file(
        process_paths.INGRAM_DAILY_REPORT_FOLDER, "Daily Report*.xlsx"
    )
    barnes_noble_file = get_latest_matching_file(
        process_paths.BN_WEEKLY_REPORT_FOLDER, "Week *.xlsx"
    )
    inventory_file = next(
        process_paths.INVENTORY_DAILY_FINANCE_ONLY_FOLDER.glob("Inventory*.xlsx"), None
    )
    amazon_preorders_file = process_paths.CURRENT_AMAZON_PREORDERS_FILE
    amazon_sellthrough_sql = process_paths.FRONTLIST_AMAZON_SELLTHROUGH_SQL
    faire_qty_sql = process_paths.FRONTLIST_FAIRE_QTY_SQL
    faire_orders_sql = process_paths.FRONTLIST_FAIRE_ORDERS_SQL

    if inventory_file is None:
        raise FileNotFoundError(
            f"No files found in {process_paths.INVENTORY_DAILY_FINANCE_ONLY_FOLDER} "
            "with pattern Inventory*.xlsx"
        )
    if not amazon_preorders_file.exists():
        raise FileNotFoundError(f"Required file not found: {amazon_preorders_file}")

    while True:
        print()
        print("Frontlist Supercharged Data will use these files:")
        print(f"  Script:                  {script_path}")
        print(f"  Frontlist Tracking:      {frontlist_file}")
        print(f"  Inventory Detail:        {inventory_file}")
        print(f"  Amazon Preorders:        {amazon_preorders_file}")
        print(f"  Ingram Daily Report:     {ingram_file}")
        print(f"  Barnes & Noble Weekly:   {barnes_noble_file}")
        print(f"  Amazon Sellthrough SQL:  {amazon_sellthrough_sql}")
        print(f"  Faire Qty SQL:           {faire_qty_sql}")
        print(f"  Faire Orders SQL:        {faire_orders_sql}")
        print("  SQL source:              sql-2-db / CBQ2")
        print(f"  Cache/output folder:     {output_dir}")
        print("  Final workbook:          Chosen in save dialog, defaulting to FLTracking_Supercharged/output")
        print()
        print("    1. Continue")
        print("    2. Return to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def confirm_bn_rolling_reports_files() -> bool:
    script_path = process_paths.BN_ROLLING_REPORTS_SCRIPT
    raw_base_folder = process_paths.BN_RAW_BASE_FOLDER
    raw_folders = sorted(
        [
            path
            for path in raw_base_folder.iterdir()
            if path.is_dir() and path.name.lower().endswith("_raw_files")
        ],
        key=lambda path: path.name,
    )
    if not raw_folders:
        raise FileNotFoundError(
            f"No Barnes & Noble raw folders were found under {raw_base_folder}"
        )

    latest_raw_folder = raw_folders[-1]
    pos_csvs = sorted(
        [path for path in latest_raw_folder.iterdir() if path.is_file() and path.suffix.lower() == ".csv"],
        key=lambda path: path.name.lower(),
    )
    sales_files = sorted(
        [
            path
            for path in latest_raw_folder.iterdir()
            if path.is_file()
            and path.suffix.lower() == ".xlsx"
            and path.name.lower().startswith("sales")
            and not path.name.startswith("~$")
        ],
        key=lambda path: path.name.lower(),
    )
    inventory_files = sorted(
        [
            path
            for path in latest_raw_folder.iterdir()
            if path.is_file()
            and path.suffix.lower() == ".xlsx"
            and path.name.lower().startswith("inventory")
            and not path.name.startswith("~$")
        ],
        key=lambda path: path.name.lower(),
    )

    def format_modified(path: Path) -> str:
        return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %I:%M:%S %p")

    while True:
        print()
        print("Barnes & Noble Rolling Reports will use these files:")
        print(f"  Script:              {script_path}")
        print(f"  Raw base folder:     {raw_base_folder}")
        print(f"  Default raw folder:  {latest_raw_folder}")
        print("  POS CSV files:")
        for file_path in pos_csvs:
            print(f"    {file_path}")
        if sales_files:
            print("  Sales workbook(s):")
            for file_path in sales_files:
                print(f"    {file_path}")
                print(f"      Last modified: {format_modified(file_path)}")
        else:
            print("  Sales workbook(s):   None found in default raw folder")
        if inventory_files:
            print("  Inventory workbook(s):")
            for file_path in inventory_files:
                print(f"    {file_path}")
                print(f"      Last modified: {format_modified(file_path)}")
        else:
            print("  Inventory workbook(s): None found in default raw folder")
        print("  Output files:        Saved in the selected raw folder during submenu actions")
        print()
        print("    1. Continue")
        print("    2. Return to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def run_program(choice):
    reports = {
        "2": ("Amazon PO Report", "amazon_po/main.py"),
        "3": ("Amazon NYP PreOrders", "amazon_preorders/main.py"),
        "4": ("Amazon Customer Orders", "amazon_customer_orders/main.py"),
        "5": ("amazon_sql_upload", "amazon_sql_upload/main.py"),
        "7": ("Amazon AMS Manager", "amazon_ams/manage_ams.py"),
        "8": ("Barnes & Noble Rolling Reports", "bn_rolling_reports/main.py"),
        "9": ("Frontlist Supercharged Data", "FLTracking_Supercharged/main.py"),
        "10": ("SSR Daily Summary", "ssr_daily_summary/main.py"),
        "11": ("UK Rolling File Combining", "UK_Rolling_File_Combining/main.py"),
        "12": ("Hachette Orders - Shipping Estimates", "hachette_orders/main.py"),
        "13": (
            "Consolidate Inventory for the INVOBS",
            "invobs_consolidated_inventory/main.py",
        ),
        "14": ("XGBoost Model", "xgboost_model/main.py"),
        "16": ("Desk Procedures", "desk_procedures/main.py"),
    }

    if choice == "1":
        # run PO archive manager in-process (Tk GUI)
        try:
            po_archive_manager.main()
        except Exception as e:
            print(f"An error occurred while running PO Archive Manager: {e}")
        return

    if choice == "6":
        run_amazon_rolling_reports_menu()
        return

    if choice in reports:
        report_name, script_path = reports[choice]
        if choice == "3":
            try:
                if not confirm_amazon_preorders_files():
                    return
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon PreOrders source files: {e}")
                return
        if choice == "4":
            try:
                if not confirm_amazon_customer_orders_files():
                    return
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Customer Orders source files: {e}")
                return
        if choice == "5":
            try:
                if not confirm_amazon_sql_upload_files():
                    return
            except FileNotFoundError as e:
                print(f"Unable to locate the amazon_sql_upload source files: {e}")
                return
        if choice == "7":
            try:
                if not confirm_amazon_ams_files():
                    return
            except (FileNotFoundError, ImportError, AttributeError) as e:
                print(f"Unable to locate the Amazon AMS source files: {e}")
                return
        if choice == "8":
            try:
                if not confirm_bn_rolling_reports_files():
                    return
            except FileNotFoundError as e:
                print(f"Unable to locate the Barnes & Noble Rolling Reports files: {e}")
                return
        if choice == "9":
            try:
                if not confirm_frontlist_supercharged_files():
                    return
            except FileNotFoundError as e:
                print(f"Unable to locate the Frontlist Supercharged source files: {e}")
                return
        if choice != "8":
            print(f"Running the {report_name}... Please wait.")
        try:
            subprocess.run(["venv/Scripts/python", script_path], check=True)
            print(f"The {report_name} is now ready.")
        except subprocess.CalledProcessError:
            print(f"An error occurred while running {script_path}.")
        return

    if choice == "15":
        run_check_table_updates_menu()
        return

    if choice == "16":
        print("Opening Desk Procedures...")
        try:
            subprocess.run(["venv/Scripts/python", "desk_procedures/main.py"], check=True)
            print("Returned from Desk Procedures.")
        except subprocess.CalledProcessError:
            print("An error occurred while running desk_procedures/main.py.")
        return

    if choice == "17":
        print(get_farewell_message())
        return

    print("Invalid choice. Please select a valid option.")


def run_amazon_rolling_reports_menu():
    while True:
        print("\nAmazon Rolling Reports")
        print()
        print("    1. Check Amazon upload table (last 10 weeks)")
        print("    2. Run normal Amazon Rolling Reports process")
        print("    3. Back to main menu")
        print()
        try:
            subchoice = input("Choose an option: ").strip().lower()
        except KeyboardInterrupt:
            print("\nReturning to main menu.")
            return

        if subchoice == "1":
            try:
                if not confirm_amazon_rolling_reports_check_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Rolling Reports check files: {e}")
                continue
            print("Running SQL check for the latest 10 weeks... Please wait.")
            try:
                subprocess.run(
                    [
                        "venv/Scripts/python",
                        "amazon_rolling_reports/check_last_10_weeks.py",
                    ],
                    check=True,
                )
            except subprocess.CalledProcessError:
                print("The SQL check failed.")
            continue

        if subchoice == "2":
            try:
                if not confirm_amazon_rolling_reports_run_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Rolling Reports source files: {e}")
                continue
            print("Running the Amazon Rolling Reports... Please wait.")
            try:
                subprocess.run(
                    ["venv/Scripts/python", "amazon_rolling_reports/main.py"],
                    check=True,
                )
                print("The Amazon Rolling Reports is now ready.")
            except subprocess.CalledProcessError:
                print("An error occurred while running amazon_rolling_reports/main.py.")
            return

        if subchoice in ["3", "back", "b", "exit", "quit", "q"]:
            return

        print("Invalid choice. Please select a valid option.")


def run_check_table_updates_menu():
    while True:
        print("\nCheck Table Updates")
        print()
        print("    1. All Updates")
        print("    2. Tables for SSR Summary")
        print("    3. Ebs.Sales Prior 5 Days")
        print("    4. Amazon")
        print("    5. Bookscan")
        print("    6. Barnes & Noble")
        print("    7. Freight Costs")
        print("    8. Back to main menu")
        print()
        try:
            subchoice = input("Choose an option: ").strip().lower()
        except KeyboardInterrupt:
            print("\nReturning to main menu.")
            return

        if subchoice in ["1", "2", "3", "4", "5", "6", "7"]:
            print("Running table-update SQL check... Please wait.")
            try:
                subprocess.run(
                    [
                        "venv/Scripts/python",
                        "table_check/check_table_updates.py",
                        subchoice,
                    ],
                    check=True,
                )
            except subprocess.CalledProcessError:
                print("The table-update SQL check failed.")
            continue

        if subchoice in ["8", "back", "b", "exit", "quit", "q"]:
            return

        print("Invalid choice. Please select a valid option.")


def main():
    print(greet_user())

    while True:
        display_options()
        try:
            choice = input(
                "\nPlease enter the number of your choice (or type 'info' to learn more): "
            ).strip()
        except KeyboardInterrupt:
            print(get_farewell_message())
            break
        if choice.isdigit():
            choice = str(int(choice))

        if choice.lower() == "info":
            choice_info = input(
                "\nEnter the number of the option you want to learn more about: "
            ).strip()
            if choice_info.isdigit():
                choice_info = str(int(choice_info))
            print(display_info(choice_info))
            continue

        if choice.lower() in ["17", "exit", "quit"]:
            run_program("17")
            break

        run_program(choice)


if __name__ == "__main__":
    main()
