import pandas as pd
from datetime import datetime
from loader.load_ho import upload_ho
from function import summarize_by_estimate_date
from ordertype_reg import calculate_est_ship_date_regular
from ordertype_bo import calculate_est_ship_date_backordered
from ordertype_rel_sal_rush import calculate_est_ship_date_released
from ordertype_hold import calculate_est_ship_date_hold

pd.set_option('display.max_columns', None)

def convert_faire_credit_hold(df):
    # Replace 'CREDIT HOLD' with 'REGULAR' in the 'OrderTypeCode' column for rows where SSR_Row is 'FAIRE WHOLESALE INC'
    # Jeff said the credit hold for Faire Warehouse only happens for a day or two and changes to regular.
    df.loc[(df['OrderTypeCode'] == 'CREDIT HOLD') & (df['SSR_Row'] == 'FAIRE WHOLESALE INC'), 'OrderTypeCode'] = 'REGULAR'
    return df 

def create_estimate_dates():
    # Load and copy the original DataFrame
    df_raw = upload_ho()    
    df = df_raw.copy()
    
    # Modify the DataFrame as needed
    convert_faire_credit_hold(df)
    
    # Generate the individual DataFrames
    df_reg = df.loc[df.OrderTypeCode == 'REGULAR'].copy()
    df_reg['EstimateDate'] = df_reg.apply(calculate_est_ship_date_regular, axis=1)
  
    df_rel = df.loc[df.OrderTypeCode.isin(['RELEASED', 'SOFT ALLOCATED','RUSH EDI AM'])].copy()
    df_rel['EstimateDate'] = df_rel.apply(calculate_est_ship_date_released, axis=1)
    
    df_hol = df.loc[df.OrderTypeCode =='HOLD'].copy()
    df_hol['EstimateDate'] = df_hol.apply(calculate_est_ship_date_hold, axis=1)
    
    df_bo = df.loc[df.OrderTypeCode == 'BACKORDERED'].copy()
    df_bo['EstimateDate'] = df_bo.apply(calculate_est_ship_date_backordered, axis=1)
    
    df_ch = df.loc[df['OrderTypeCode'] == 'CREDIT HOLD'].copy()
    df_ch['EstimateDate'] = pd.NaT
    
    # Concatenate all DataFrames into one
    df_combined = pd.concat([df_reg, df_rel, df_hol, df_bo,df_ch], ignore_index=True)
    
    issues_search(df_combined)
    
    return df_combined

def main(): 
    
    df_combined = create_estimate_dates()
    
    # Summarize the "val" field by EstimateDate
    daily_summary, summary = summarize_by_estimate_date(df_combined)
    
    # Print daily summary
    print("Daily Summary for the Next 5 Days:")
    for index, row in daily_summary.iterrows():
        print(f"Date: {row['EstimateDate'].date()}, Value: {row['val']}")
    
    # Print overall summary
    print("\nOverall Summary:")
    for key, value in summary.items():
        print(f"{key}: {value}")

if __name__ == "__main__":
    main()