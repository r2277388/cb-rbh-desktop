import getpass
import subprocess
from datetime import datetime


def get_full_name():
    # Dictionary to map usernames to real names
    USER_NAMES = {
        "kbs": "Kate Breiting Schmitz",
        "mjk": "Marlena Kwasnik",
        "sdm": "Sam Mariucci",
        "RBH": "Barrett Hooper",  # Add more as needed
        }

    username = getpass.getuser()
    full_name = USER_NAMES.get(username, username)
    return full_name


def greet_user():
    """Greet the user based on the time of day and day of the week."""
    current_datetime = datetime.now()
    current_hour = current_datetime.hour
    current_day = current_datetime.strftime('%A')

    if current_hour < 12:
        greeting = "Good morning"
    elif 12 <= current_hour < 17:
        greeting = "Good afternoon"
    else:
        greeting = "Good evening"

    full_name = get_full_name()
    
    return f"\n{greeting}, {full_name}! Happy {current_day}!"


def get_farewell_message():
    """Return a farewell message based on the time of day."""
    current_hour = datetime.now().hour
    
    full_name = get_full_name()

    if 12 <= current_hour < 17:
        return f"\nHave a great afternoon, {full_name}!"
    elif 17 <= current_hour < 24:
        return f"\nGood evening, {full_name}!"
    else:
        return f"\nHave a great day, {full_name}!"


def display_options():
    """Display the available program options."""
    options = [
        "1. Amazon PO Report", 
        "2. Amazon PreOrders",
        "3. Amazon Customer Orders",
        "4. SSR Daily Summary",
        "5. UK Rolling File Combining",
        "6. Hachette Orders - Shipping Estimates",
        "7. Consolidate Inventory for the INVOBS",
        "8. Amazon Rolling Reports",
        "9. Exit"
    ]
    print("\nWhat would you like to run?")
    for option in options:
        print(option)


def display_info(choice):
    """Display information about the selected option."""
    info = {
        '1': """Amazon PO Report: Generates a detailed report based on Amazon Purchase Orders.
        Before running, save the Vendor Central PO File to the following location:
        G:\\SALES\\Amazon\\PURCHASE ORDERS\\atelier\\po_analysis\\PurchaseOrderItems.csv
        A PO Report is saved off to: G:\\SALES\\Amazon\\PURCHASE ORDERS folder""",
        '2': """Amazon NYP PreOrders: Generates a report for Amazon NYP PreOrders.
        Save the relevant data file to the appropriate location before running.""",
        '3': """Amazon Customer Orders: Generates a report for Amazon Customer Orders.
        Save the relevant data file to the appropriate location before running.""",
        '4': "SSR Daily Summary: Prepares the data for the SSR Daily Summary email.",
        '5': "UK Rolling File Combining: This combines the sales, reserve and midas files together.",
        '6': "Hachette Orders - Shipping Estimates: Generates a report for Hachette Orders.",
        '7': """Consolidate Inventory for the INVOBS: Runs the Consolidated Inventory program for INVOBS.
        This program takes the consolidated inventory data from Oracle, run by Ailing and\
        explodes out the CDU's into their components to gives us a component-only inventory file.""",
        '8': "Amazon Rolling Reports: Runs the full Amazon Rolling Reports workflow.",
        '9': "Exit: Exits the program."
    }
    return info.get(choice, "Invalid choice. No information available.")


def run_program(choice):
    """Run the selected program with a loading indicator."""
    reports = {
        '1': ("Amazon PO Report", "amazon_po/main.py"),
        '2': ("Amazon NYP PreOrders", "amazon_preorders/main.py"),
        '3': ("Amazon Customer Orders", "amazon_customer_orders/main.py"),
        '4': ("SSR Daily Summary", "ssr_daily_summary/main.py"),
        '5': ("UK Rolling File Combining", "UK_Rolling_File_Combining/main.py"),
        '6': ("Hachette Orders - Shipping Estimates", "hachette_orders/main.py"),
        '7': ("Consolidate Inventory for the INVOBS", "invobs_consolidated_inventory/main.py"),
        '8': ("Amazon Rolling Reports", "amazon_rolling_reports/main.py")
    }

    if choice in reports:
        report_name, script_path = reports[choice]
        print(f"Running the {report_name}... Please wait.")
        try:
            subprocess.run(["venv/Scripts/python", script_path], check=True)
            print(f"The {report_name} is now ready.")
        except subprocess.CalledProcessError as e:
            print(f"An error occurred while running the {script_path}.")
    elif choice == '9':
        print(get_farewell_message())
    else:
        print("Invalid choice. Please select a valid option.")


def main():
    print(greet_user())

    while True:
        display_options()
        choice = input("\nPlease enter the number of your choice (or type 'info' to learn more): ").strip()

        if choice.lower() == 'info':
            choice_info = input("\nEnter the number of the option you want to learn more about: ").strip()
            print(display_info(choice_info))
            continue  # Return to the options list after displaying info

        if choice.lower() in ['9', 'exit', 'quit']:
            run_program('9')
            break

        run_program(choice)

if __name__ == "__main__":
    main()
