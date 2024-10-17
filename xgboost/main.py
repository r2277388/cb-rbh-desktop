from functions_pgrp import (
    pgrp_function,
    pgrp_flbl_function,
    check_year_totals,
    check_year_totals_flbl
)

def main():
    df = pgrp_function()
    df_flbl = pgrp_flbl_function()

    # Set a name for the index, if necessary
    df.index.name = 'Date'  # or use a more descriptive name
    df_flbl.index.name = 'Date'

    return df, df_flbl

if __name__ == "__main__":
    df, df_flbl = main()
    check_year_totals(df)
    check_year_totals_flbl(df_flbl)

    # Save the DataFrames to Excel files
    df.to_csv("pgrp_forecasts.csv", index=True)
    df_flbl.to_csv("pgrp_flbl_forecasts.csv", index=True)
    
    print("DataFrames have been saved to pgrp_forecasts.csv and pgrp_flbl_forecasts.csv")