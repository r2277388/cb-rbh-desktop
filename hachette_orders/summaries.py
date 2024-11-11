import pandas as pd

file_name = r'C:\Users\rbh\code\hachette_orders\ho_20241111_1318.pkl'
df = pd.read_pickle(file_name)

def sum_val_next_5_days_by_ssr_row(df):
    # Get today's date and the date 5 days from now
    todays_date = pd.Timestamp.today().normalize()
    next_5_days = todays_date + pd.DateOffset(days=5)
    
    # Filter the DataFrame for the next 5 days
    df_next_5_days = df[(df['EstimateDate'] >= todays_date) & (df['EstimateDate'] < next_5_days)]
    
    # Group by SSR_Row and EstimateDate, and sum the 'val' column
    summary = df_next_5_days.groupby(['SSR_Row', 'EstimateDate'])['val'].sum().unstack(fill_value=0)
    
    # Print the summary
    print("Sum of 'val' for the next 5 days by SSR_Row:")
    print(summary)
    
    return summary