import pandas as pd
from datetime import datetime
from ordertype_estimates_combined import create_estimate_dates

df = create_estimate_dates()

def create_pickle_file(df):
    # Generate filename with current date and time
    current_time = datetime.now().strftime('%Y%m%d_%H%M')
    filename = f'ho_{current_time}.pkl'

    # Save the combined DataFrame to a pickle file
    df.to_pickle(filename)
    print(f"File saved as {filename}")
    
    
def main():
    #create_pickle_file(df)
    file_name = r'C:\Users\rbh\code\hachette_orders\ho_20241104_1558.pkl'
    df = pd.read_pickle(file_name)
    
    df_reg_issues = df.loc[
        (df.OrderTypeCode == 'REGULAR') &
        (df.EstimateDate == pd.Timestamp.today().normalize())
        ]
    print(df_reg_issues.head())

if __name__ == "__main__":
    main()