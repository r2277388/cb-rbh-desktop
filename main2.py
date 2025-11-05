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
        "1. PO Archive Manager",
        "2. Amazon PO Report",
        "3. Amazon PreOrders",
        "4. amazon_sql_upload",
        "5. Amazon Rolling Reports",
        "6. Amazon Customer Orders",
        "7. SSR Daily Summary",
        "8. UK Rolling File Combining",
        "9. Hachette Orders - Shipping Estimates",
        "10. Consolidate Inventory for the INVOBS",
        "11. Exit",
    ]
    print("\nWhat would you like to run?")
    for option in options:
        print(option)


def display_info(choice):
    info = {
        "1": "PO Archive Manager: Launches the PO archive helper to archive prior current_amaz_preorders and copy the new file into po_analysis.",
        "2": """Amazon PO Report: Generates a detailed report based on Amazon Purchase Orders.
        Before running, save the Vendor Central PO File to:
        G:\\SALES\\Amazon\\PURCHASE ORDERS\\atelier\\po_analysis\\PurchaseOrderItems.csv
        A PO Report is saved off to: G:\\SALES\\Amazon\\PURCHASE ORDERS folder""",
        "3": "Amazon NYP PreOrders: Generates a report for Amazon NYP PreOrders. Save the relevant data file to the appropriate location before running.",
        "4": "amazon_sql_upload: Runs the amazon_sql_upload workflow (ASIN/ISBN conversion, uploads, etc.).",
        "5": "Amazon Rolling Reports: Runs the full Amazon Rolling Reports workflow.",
        "6": "Amazon Customer Orders: Generates a report for Amazon Customer Orders. Save the relevant data file to the appropriate location before running.",
        "7": "SSR Daily Summary: Prepares the data for the SSR Daily Summary email.",
        "8": "UK Rolling File Combining: This combines the sales, reserve and midas files together.",
        "9": "Hachette Orders - Shipping Estimates: Generates a report for Hachette Orders.",
        "10": """Consolidate Inventory for the INVOBS: Runs the Consolidated Inventory program for INVOBS.
        This program takes the consolidated inventory data from Oracle, run by Ailing, and
        explodes out the CDU's into their components to give a component-only inventory file.""",
        "11": "Exit: Exits the program.",
    }
    return info.get(choice, "Invalid choice. No information available.")


def run_program(choice):
    reports = {
        "2": ("Amazon PO Report", "amazon_po/main.py"),
        "3": ("Amazon NYP PreOrders", "amazon_preorders/main.py"),
        "4": ("amazon_sql_upload", "amazon_sql_upload/main.py"),
        "5": ("Amazon Rolling Reports", "amazon_rolling_reports/main.py"),
        "6": ("Amazon Customer Orders", "amazon_customer_orders/main.py"),
        "7": ("SSR Daily Summary", "ssr_daily_summary/main.py"),
        "8": ("UK Rolling File Combining", "UK_Rolling_File_Combining/main.py"),
        "9": ("Hachette Orders - Shipping Estimates", "hachette_orders/main.py"),
        "10": (
            "Consolidate Inventory for the INVOBS",
            "invobs_consolidated_inventory/main.py",
        ),
    }

    if choice == "1":
        # run PO archive manager in-process (Tk GUI)
        try:
            po_archive_manager.main()
        except Exception as e:
            print(f"An error occurred while running PO Archive Manager: {e}")
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

    if choice == "11":
        print(get_farewell_message())
        return

    print("Invalid choice. Please select a valid option.")


def main():
    print(greet_user())

    while True:
        display_options()
        choice = input(
            "\nPlease enter the number of your choice (or type 'info' to learn more): "
        ).strip()

        if choice.lower() == "info":
            choice_info = input(
                "\nEnter the number of the option you want to learn more about: "
            ).strip()
            print(display_info(choice_info))
            continue

        if choice.lower() in ["11", "exit", "quit"]:
            run_program("11")
            break

        run_program(choice)


if __name__ == "__main__":
    main()
