from tqdm import tqdm
import pandas as pd
from variables import get_variables
from functions import get_connection
from queries import query_mtd, query_ytd, query_check_cbq_metrics

def get_user_input():
    """Get user input for the type of report they want to generate."""
    print("Options:")
    print("1 - See current month sales")
    print("2 - See YTD sales")
    print("3 - See the latest 25 updated tables in CBQ2")
    
    choices = input("Enter the numbers of the reports you want to generate, separated by commas (e.g., 1,2): ")
    return [choice.strip() for choice in choices.split(',')]

def run_query(query_func, format_numbers=False, *args):
    """Run a query function, optionally format the results, and display with a progress bar."""
    with tqdm(total=1, desc="Running query", leave=False, ncols=100):
        engine = get_connection()
        df = pd.read_sql_query(query_func(*args), engine)
    
    # Format numbers if requested (for options 1 and 2)
    if format_numbers:
        # Convert columns to integers to remove decimals
        for col in df.select_dtypes(include='number').columns:
            df[col] = df[col].astype(int)
        # Apply comma formatting and print as a formatted string
        print(df.to_string(index=False, formatters={col: "{:,}".format for col in df.select_dtypes(include='number').columns}))
    else:
        # Display DataFrame as-is without formatting
        print(df)

def main():
    # Get report options from user
    choices = get_user_input()
    
    # Get variables for queries
    variables = get_variables(use_current_date=True)
    if not variables:
        print("Failed to retrieve variables.")
        return
    prior_day, prior_day_ly, tp, typ1, tply, lyp1 = variables

    # Run selected queries based on user choices
    for choice in choices:
        if choice == '1':
            print("Current Month Sales:")
            run_query(query_mtd, True, tp)
        elif choice == '2':
            print("YTD Sales:")
            run_query(query_ytd, True, typ1)
        elif choice == '3':
            print("Latest 25 Updated Tables in CBQ2:")
            row = 25  # Set the number of rows to retrieve for the latest updates
            run_query(query_check_cbq_metrics, False,25)
        else:
            print(f"Invalid choice: {choice}. Please enter 1, 2, or 3.")

if __name__ == "__main__":
    main()