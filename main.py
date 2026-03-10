import getpass
import subprocess
from datetime import datetime

# call the PO archive manager directly
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
        "08. SSR Daily Summary",
        "09. UK Rolling File Combining",
        "10. Hachette Orders - Shipping Estimates",
        "11. Consolidate Inventory for the INVOBS",
        "12. XGBoost Model",
        "13. Check Table Updates",
        "14. Exit",
    ]
    print("\nWhat would you like to run?")
    print()
    for option in options:
        print(f"    {option}")


def display_info(choice):
    info = {
        "1": "PO Archive Manager: Launches the PO archive helper to archive prior current_amaz_preorders and copy the new file into po_analysis.",
        "2": """Amazon PO Report: Generates a detailed report based on Amazon Purchase Orders.
        Before running, save the Vendor Central PO File to:
        G:\\SALES\\Amazon\\PURCHASE ORDERS\\atelier\\po_analysis\\PurchaseOrderItems.csv
        A PO Report is saved off to: G:\\SALES\\Amazon\\PURCHASE ORDERS folder""",
        "3": "Amazon NYP PreOrders: Generates a report for Amazon NYP PreOrders. Save the relevant data file to the appropriate location before running.",
        "4": "Amazon Customer Orders: Generates a report for Amazon Customer Orders. Save the relevant data file to the appropriate location before running.",
        "5": "amazon_sql_upload: Runs the amazon_sql_upload workflow (ASIN/ISBN conversion, uploads, etc.).",
        "6": "Amazon Rolling Reports: Runs a 10-week SQL freshness check first, then asks whether to continue with the full process.",
        "7": "Amazon AMS Manager: Manage/update AMS month configuration and run incremental or full AMS processing.",
        "8": "SSR Daily Summary: Prepares the data for the SSR Daily Summary email.",
        "9": "UK Rolling File Combining: This combines the sales, reserve and midas files together.",
        "10": "Hachette Orders - Shipping Estimates: Generates a report for Hachette Orders.",
        "11": """Consolidate Inventory for the INVOBS: Runs the Consolidated Inventory program for INVOBS.
        This program takes the consolidated inventory data from Oracle, run by Ailing, and
        explodes out the CDU's into their components to give a component-only inventory file.""",
        "12": "XGBoost Model: Launches the xgboost_model workflow menu.",
        "13": "Check Table Updates: Runs SQL checks for table freshness and recent weeks for SSR/Amazon/Bookscan tables.",
        "14": "Exit: Exits the program.",
    }
    return info.get(choice, "Invalid choice. No information available.")


def run_program(choice):
    reports = {
        "2": ("Amazon PO Report", "amazon_po/main.py"),
        "3": ("Amazon NYP PreOrders", "amazon_preorders/main.py"),
        "4": ("Amazon Customer Orders", "amazon_customer_orders/main.py"),
        "5": ("amazon_sql_upload", "amazon_sql_upload/main.py"),
        "7": ("Amazon AMS Manager", "amazon_ams/manage_ams.py"),
        "8": ("SSR Daily Summary", "ssr_daily_summary/main.py"),
        "9": ("UK Rolling File Combining", "UK_Rolling_File_Combining/main.py"),
        "10": ("Hachette Orders - Shipping Estimates", "hachette_orders/main.py"),
        "11": (
            "Consolidate Inventory for the INVOBS",
            "invobs_consolidated_inventory/main.py",
        ),
        "12": ("XGBoost Model", "xgboost_model/main.py"),
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
        print(f"Running the {report_name}... Please wait.")
        try:
            subprocess.run(["venv/Scripts/python", script_path], check=True)
            print(f"The {report_name} is now ready.")
        except subprocess.CalledProcessError:
            print(f"An error occurred while running {script_path}.")
        return

    if choice == "13":
        run_check_table_updates_menu()
        return

    if choice == "14":
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
        print("    3. Amazon")
        print("    4. Bookscan")
        print("    5. Barnes & Noble")
        print("    6. Back to main menu")
        print()
        try:
            subchoice = input("Choose an option: ").strip().lower()
        except KeyboardInterrupt:
            print("\nReturning to main menu.")
            return

        if subchoice in ["1", "2", "3", "4", "5"]:
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

        if subchoice in ["6", "back", "b", "exit", "quit", "q"]:
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

        if choice.lower() in ["14", "exit", "quit"]:
            run_program("14")
            break

        run_program(choice)


if __name__ == "__main__":
    main()
