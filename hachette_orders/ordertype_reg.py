import pandas as pd
from loader.load_ho import upload_ho
from list_provider import bn_and_ps_shipping_days, readerlink_shipping_days,bn_list
from function import adjust_to_weekday, next_three_days,summarize_by_estimate_date

def calculate_est_ship_date_regular(df):
    todays_date = pd.Timestamp('today').normalize()
    min_estimated_date = adjust_to_weekday(todays_date + pd.DateOffset(days=3))
    # READERLINK REGULAR

    if (df.SSR_Row == 'Readerlink') and pd.notnull(df.WMSDoNotDeliverAfter):
        if df.STATE in readerlink_shipping_days:
            days_offset = readerlink_shipping_days[df.STATE]
            estimated_date = df.WMSDoNotDeliverAfter + pd.DateOffset(days=days_offset)
            return adjust_to_weekday(estimated_date)
    
    # Barnes & Noble and Paper Source REGULAR

    if df.SSR_Row in bn_list:
        if pd.notnull(df.WMSDoNotDeliverAfter):
            if df.STATE in bn_and_ps_shipping_days:
                days_offset = bn_and_ps_shipping_days[df.STATE]
                estimated_date = df.WMSDoNotDeliverAfter - pd.DateOffset(days=days_offset)
                return adjust_to_weekday(estimated_date)
            else:
                estimated_date = df.WMSDoNotDeliverAfter - pd.DateOffset(days=2)
                return adjust_to_weekday(estimated_date)
        elif pd.notnull(df.ReleaseDate) and (df.ReleaseDate > todays_date):
            return adjust_to_weekday(df.ReleaseDate)
        elif pd.notnull(df.ReleaseDate) and (df.ReleaseDate <= todays_date):
            estimated_date = next_three_days()+pd.DateOffset(days=3) # regular orders need at least 3 days to ship
            return adjust_to_weekday(estimated_date)
        
    # Orders with ReleaseDate and OrderCancelDate
    if pd.notnull(df.OrderCancelDate):
        estimated_date = df.OrderCancelDate - pd.DateOffset(days=7)        
        if estimated_date < min_estimated_date:
            estimated_date = min_estimated_date
        
        return adjust_to_weekday(estimated_date)
    
    # No ReleaseDate and no OrderCancelDate
    if pd.isnull(df.ReleaseDate) and pd.isnull(df.OrderCancelDate) and pd.notnull(df.EnteredDate):
        # Check if the difference between today's date and EnteredDate is more than 15 days
        if (todays_date - df.EnteredDate).days > 15:
            return pd.NaT  # Return NaT if EnteredDate is more than 15 days old
        else:
            estimated_date = df.EnteredDate + pd.DateOffset(days=10)
            if estimated_date < min_estimated_date:
                estimated_date = min_estimated_date
            return adjust_to_weekday(estimated_date)  # Return 10 days after EnteredDate
    
    # General case for ReleaseDate
    if pd.notnull(df.ReleaseDate) and (df.ReleaseDate > todays_date):
        if df.ReleaseDate < min_estimated_date:
            return min_estimated_date
        return adjust_to_weekday(df.ReleaseDate)
    
    elif pd.notnull(df.ReleaseDate) and (df.ReleaseDate <= todays_date):
        estimated_date = next_three_days()+pd.DateOffset(days=3) # regular orders need at least 3 days to ship
        return adjust_to_weekday(estimated_date)
    
    else:
        return pd.NaT
   
def get_regular(df):
        df = df.loc[df.OrderTypeCode == 'REGULAR']
        df['EstimateDate'] = df.apply(calculate_est_ship_date_regular, axis=1)
        return df

def issues_search(df):
    # Push out the EstimateDate by 3 days only for REGULAR orders scheduled to ship today
    today = pd.Timestamp.today().normalize()
    
    # Apply the 3-day offset only to rows that match both conditions
    df.loc[
        (df['OrderTypeCode'] == 'REGULAR') & (df['EstimateDate'] == today),
        'EstimateDate'
    ] += pd.DateOffset(days=3)
    
    return df
def main():
    df = upload_ho()
    df = get_regular(df)
    print(df.info())
    print(df.head())
    print()

    # Summarize the "val" field by EstimateDate
    daily_summary, summary = summarize_by_estimate_date(df)
    
    # Print daily summary
    print("Daily Summary for the Next 5 Days:")
    for index, row in daily_summary.iterrows():
        print(f"Date: {row['EstimateDate'].date()}, Value: {row['val']}")
    
    # Print overall summary
    print("\nOverall Summary:")
    for key, value in summary.items():
        print(f"{key}: {value}")
        
    # print('Rows with missing EstimateDate:')
    # nat_rows = df[df['EstimateDate'].isna()]
    # print(nat_rows)    
        
if __name__ == '__main__':
    main()