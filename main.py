import os
import getpass
import subprocess
from datetime import datetime

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

    username = getpass.getuser()
    return f"\n{greeting}, {username}! Happy {current_day}!"

def get_farewell_message():
    """Return a farewell message based on the time of day."""
    current_hour = datetime.now().hour

    if 12 <= current_hour < 17:
        return "Have a great afternoon!"
    elif 17 <= current_hour < 24:
        return "Good evening!"
    else:
        return "Have a great day!"

def display_options():
    """Display the available program options."""
    options = [
        "1. Amazon PO Report", 
        "2. Amazon PreOrders",
        "3. Amazon Customer Orders",
        "4. Pickle Raw Data",
        "5. XGBoost Forecast",
        "6. SSR Daily Summary",
        "7. Exit"
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
        '4': "Pickle Raw Data: Updates the ebs.sales datadump pickle file. Please run before running the XGBoost Forecast.",
        '5': "XGBoost Forecast: Runs a machine learning model to forecast future data trends using XGBoost.",
        '6': "SSR Daily Summary: Prepares the data for the SSR Daily Summary email.",
        '7': "Exit: Exits the program."
    }
    return info.get(choice, "Invalid choice. No information available.")

def run_program(choice):
    """Run the selected program with a loading indicator."""
    reports = {
        '1': ("Amazon PO Report", "amazon_po/main.py"),
        '2': ("Amazon NYP PreOrders", "amazon_preorders/main.py"),
        '3': ("Amazon Customer Orders", "amazon_customer_orders/main.py"),
        '4': ("Pickle Raw Data", "pickle_raw_data/main.py"),
        '5': ("XGBoost Forecast", "xgboost/main.py"),
        '6': ("SSR Daily Summary", "ssr_daily_summary/main.py")
    }

    if choice in reports:
        report_name, script_path = reports[choice]
        print(f"Running the {report_name}... Please wait.")
        subprocess.run(["python", script_path])
        print(f"The {report_name} is now ready.")
    elif choice == '7':
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

        if choice == '7':  # Option '7' to exit
            run_program(choice)
            break  # Exit the while loop to end the program

        run_program(choice)

if __name__ == "__main__":
    main()