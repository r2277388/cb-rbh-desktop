import pandas as pd
from function import adjust_to_weekday,summarize_by_estimate_date
from loader.load_ho import upload_ho

def calculate_est_ship_date_backordered(df):
    # If there is no ReprintDate, the book is not being reprinted and won't have an estimated shipping date.
    if pd.isnull(df.ReprintDate):
        return pd.NaT
    
    # CASES with both an OrdereCanceldDate and a ReprintDate
    elif pd.notnull(df.OrderCancelDate) and pd.notnull(df.ReprintDate):
        # if the OrderCancelDate is more than 3 days from the ReprintDate, it'll be shipped 3 days after the ReprintDate
        if df.ReprintDate + pd.DateOffset(days=4) <= df.OrderCancelDate:
            estimated_date = df.ReprintDate + pd.DateOffset(days=3)
            return adjust_to_weekday(estimated_date)
        else:
            # if the canceldate is less than 3 days from the reprint date, it'll be canceled.
            return pd.NaT
        
    elif pd.isnull(df.OrderCancelDate) and pd.notnull(df.ReprintDate):
        estimated_date = df.ReprintDate + pd.DateOffset(days=3)
        return adjust_to_weekday(estimated_date)

    else:
        return pd.NaT
    
def get_backordered(df):
        df = df.loc[df.OrderTypeCode.isin(['BACKORDERED'])]
        df['EstimateDate'] = df.apply(calculate_est_ship_date_backordered, axis=1)
        return df

def main():
    df = upload_ho()
    df = get_backordered(df)
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