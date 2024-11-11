import pandas as pd
from datetime import datetime
from ordertype_estimates_combined import create_estimate_dates
from function import sum_val_next_5_days_by_ssr_row

def create_pickle_file(df):
    # Generate filename with current date and time
    current_time = datetime.now().strftime('%Y%m%d_%H%M')
    filename = f'ho_{current_time}.pkl'

    # Save the combined DataFrame to a pickle file
    df.to_pickle(filename)
    print(f"File saved as {filename}")
    
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
    print('Next 5 Day Estimate')
    sum_val_next_5_days_by_ssr_row(df)

if __name__ == "__main__":
    main()