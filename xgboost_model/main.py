import getpass
import subprocess
import sys
from pathlib import Path
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
        "1. Pickle Raw Data",
        "2. XGBoost Forecast",
        "3. Exit"
    ]
    print("\nWhat would you like to run?")
    for option in options:
        print(option)

def display_info(choice):
    """Display information about the selected option."""
    info = {
        '1': "Pickle Raw Data: Updates the ebs.sales datadump pickle file. Please run before running the XGBoost Forecast.",
        '2': "XGBoost Forecast: Runs a machine learning model to forecast future data trends using XGBoost.",
        '3': "Exit: Exits the program."
    }
    return info.get(choice, "Invalid choice. No information available.")

def run_program(choice):
    """Run the selected program with a loading indicator."""
    base_dir = Path(__file__).resolve().parent
    reports = {
        '1': ("Pickle Raw Data", "save_raw_data/main.py"),
        '2': ("XGBoost Forecast", "xgboost/main.py"),
    }

    if choice in reports:
        report_name, script_path = reports[choice]
        resolved_script_path = base_dir / script_path
        print(f"Running the {report_name}... Please wait.")

        # Reuse the currently active interpreter (e.g., C:\Users\rbh\code\venv).
        subprocess.run(
            [sys.executable, str(resolved_script_path)],
            check=True,
            cwd=base_dir,
        )
        
        print(f"The {report_name} is now ready.")
    elif choice == '3':
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

        if choice == '3':  # Option '3' to exit
            run_program(choice)
            break  # Exit the while loop to end the program

        run_program(choice)
    
if __name__ == "__main__":
    # Start timing
    start_time = datetime.now()
    main()
    end_time = datetime.now()
    elapsed_time = end_time - start_time
    print(f"Time taken: {elapsed_time}")
