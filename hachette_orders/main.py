import pandas as pd
import os
from datetime import datetime

from ordertype_estimates_combined import create_estimate_dates
from function import sum_val_next_5_days_by_ssr_row

import os
from datetime import datetime

def create_pickle_file(df):
    # Ask the user if they want to save the file as a pickle
    should_pickle = input("Do you want to save the file as a pickle? (yes/no): ").strip().lower()
    
    if should_pickle == 'yes':
        # Generate filename with current date and time
        current_time = datetime.now().strftime('%Y%m%d_%H%M')
        filename = f'ho_{current_time}.pkl'
        
        # Save the combined DataFrame to a pickle file
        df.to_pickle(filename)
        
        # Get the absolute path of the saved file
        file_path = os.path.abspath(filename)
        print(f"File saved at {file_path}")
    else:
        print("File not saved as a pickle.")
    
def sum_val_next_5_days(df):
    todays_date = pd.Timestamp.today().normalize()
    next_5_days = todays_date + pd.DateOffset(days=5)
    
    # Filter the DataFrame for the next 5 days
    df_next_5_days = df[(df.EstimateDate >= todays_date) & (df.EstimateDate <= next_5_days)]
    
    # Group by Publisher and EstimateDate, and sum the val field
    summary = df_next_5_days.groupby(['Publisher', 'EstimateDate'])['val'].sum().reset_index()
    
    # Print the summary
    print("Sum of 'val' for the next 5 days for each Publisher:")
    for index, row in summary.iterrows():
        print(f"Publisher: {row['Publisher']}, Date: {row['EstimateDate'].date()}, Value: {row['val']:,}")
    
    return summary
    
def main():
    df = create_estimate_dates()
    
    # This saves off the df to a pickle file
    create_pickle_file(df)  
    print()
    sum_val_next_5_days(df)
    print('Next 5 Day Estimate')
    sum_val_next_5_days_by_ssr_row(df)

if __name__ == "__main__":
    main()