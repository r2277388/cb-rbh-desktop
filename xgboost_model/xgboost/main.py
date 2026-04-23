import time

from functions_pgrp import (
    pgrp_function,
    pgrp_flbl_function,
    check_year_totals,
    check_year_totals_flbl,
    calculate_monthly_forecast
)

def main():
    df = pgrp_function()
    df_flbl = pgrp_flbl_function()

    # Set a name for the index, if necessary
    df.index.name = 'Date'  # or use a more descriptive name
    df_flbl.index.name = 'Date'

    return df, df_flbl

if __name__ == "__main__":
    start_time = time.time()  # Start timing
    df, df_flbl = main()
    check_year_totals(df)
    check_year_totals_flbl(df_flbl)
    print()
    # Calculate and print monthly forecast totals for each DataFrame
    print("\nEstimates without FLBL:")
    monthly_totals_no_flbl = calculate_monthly_forecast(df)
    print(f"Current Month Total Estimate: {monthly_totals_no_flbl['current_month_total']}")
    print(f"Next Month Total Estimate: {monthly_totals_no_flbl['next_month_total']}")
    
    print("\nEstimates with FLBL:")
    monthly_totals_flbl = calculate_monthly_forecast(df_flbl)
    print(f"Current Month Total Estimate: {monthly_totals_flbl['current_month_total']}")
    print(f"Next Month Total Estimate: {monthly_totals_flbl['next_month_total']}")
    print()
    
    # Save the DataFrames to Excel files
    df.to_csv("pgrp_forecasts.csv", index=True)
    df_flbl.to_csv("pgrp_flbl_forecasts.csv", index=True)
    
    print("DataFrames have been saved to pgrp_forecasts.csv and pgrp_flbl_forecasts.csv")
    end_time = time.time()  # End timing
    print(f"Total execution time: {end_time - start_time:.2f} seconds")