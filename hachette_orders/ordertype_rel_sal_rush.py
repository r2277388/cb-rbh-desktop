import pandas as pd
from function import adjust_to_weekday, next_three_days,summarize_by_estimate_date
from list_provider import hbg_tier1, hbg_tier2, hbg_tier3, hbg_tier4, hbg_tier5, big_box_stores
from loader.load_ho import upload_ho
# Set pandas display option to show all columns
pd.set_option('display.max_columns', None)


def calculate_est_ship_date_released(df):
    todays_date = pd.Timestamp('today').normalize()
    
    if (df.OrderTypeCode == 'RUSH EDI AM'):
        return adjust_to_weekday(todays_date)
    ######################
    # PRE-OSD Rules
    ######################
    elif pd.notnull(df.osd) and (df.osd > todays_date) and (df.SSR_Row in big_box_stores):
        if (df.STATE in hbg_tier1):
            estimated_date = df.osd - pd.DateOffset(days=10)
            return adjust_to_weekday(estimated_date)
        elif (df.STATE in hbg_tier2):
            estimated_date = df.osd - pd.DateOffset(days=12)
            return adjust_to_weekday(estimated_date)
        elif (df.STATE in hbg_tier3):
            estimated_date = df.osd - pd.DateOffset(days=13)
            return adjust_to_weekday(estimated_date)
        elif (df.STATE in hbg_tier4):
            estimated_date = df.osd - pd.DateOffset(days=13)
            return adjust_to_weekday(estimated_date)
        elif (df.STATE in hbg_tier5):
            estimated_date = df.osd - pd.DateOffset(days=14)
            return adjust_to_weekday(estimated_date)
        
        # For Not Big Box Stores
    if pd.notnull(df.osd) and (df.osd > todays_date) and (df.SSR_Row not in big_box_stores):
        if (df.STATE in hbg_tier1):
            estimated_date = df.osd - pd.DateOffset(days=3)
            return adjust_to_weekday(estimated_date)
        elif (df.STATE in hbg_tier2):
            estimated_date = df.osd - pd.DateOffset(days=4)
            return adjust_to_weekday(estimated_date)
        elif (df.STATE in hbg_tier3):
            estimated_date = df.osd - pd.DateOffset(days=6)
            return adjust_to_weekday(estimated_date)
        elif (df.STATE in hbg_tier4):
            estimated_date = df.osd - pd.DateOffset(days=8)
            return adjust_to_weekday(estimated_date)
        elif (df.STATE in hbg_tier5):
            estimated_date = df.osd - pd.DateOffset(days=10)
            return adjust_to_weekday(estimated_date)
    
    # ######################
    # # POST OSD Rules
    # ######################
    # # READERLINK RELEASED
    readerlink_shipping_days = {
        'IL': 6,   # 1 day shipping
        'VA': 5,   # 2 day shipping
        'TX': 4,   # 3 day shipping
        'GA': 4,   # 3 day shipping
        'UT': 3    # 4 day shipping
    }
    if (df.SSR_Row == 'Readerlink') and (df.WMSDoNotDeliverAfter is not None):
        if df.STATE in readerlink_shipping_days:
            days_offset = readerlink_shipping_days[df.STATE]
            estimated_date = df.WMSDoNotDeliverAfter + pd.DateOffset(days=days_offset)
            return adjust_to_weekday(estimated_date)
    
    # BARNES & NOBLE OR PAPER SOURCE RELEASED 
    bn_and_ps_shipping_days = {
        'IL': 1,   # 1 day shipping
        'NJ': 2,   # 2 day shipping
        'NV': 4    # 4 day shipping
    }
    if (df.SSR_Row in ['Barnes & Noble', 'Barnes & Noble College','Paper Source']):
        if pd.notnull(df.WMSDoNotDeliverAfter):
            if df.STATE in bn_and_ps_shipping_days:
                if df.STATE in bn_and_ps_shipping_days:
                    days_offset = bn_and_ps_shipping_days[df.STATE]
                    estimated_date = df.WMSDoNotDeliverAfter - pd.DateOffset(days=days_offset)
                    return adjust_to_weekday(estimated_date)
        elif pd.notnull(df.ReleaseDate) and (df.ReleaseDate > todays_date):
            return adjust_to_weekday(df.ReleaseDate)
        elif pd.notnull(df.ReleaseDate) and (df.ReleaseDate <= todays_date):
            return next_three_days()
    
    # AMAZON RELEASED
    amazon_shipping_days = {
        'PA': 1,   # 1 day shipping
        'TN': 2,   # 2 day shipping
        'CA': 4    # 4 day shipping
    }
    if (df.SSR_Row == 'Amazon.com') and (df.WMSDoNotShipAfter is not None):
        if df.STATE in amazon_shipping_days:
            days_offset = amazon_shipping_days[df.STATE]
            estimated_date = df.WMSDoNotShipAfter - pd.DateOffset(days=days_offset)
            return adjust_to_weekday(estimated_date)
        
    elif (df.SSR_Row == 'Amazon.com') and (df.OrderCancelDate is not None) and (df.OrderCancelDate <= todays_date):
        return adjust_to_weekday(df.OrderCancelDate)
    
    # Target Direct RELEASED
    elif (df.SSR_Row == 'Target Direct') and (df.OrderCancelDate is not None):
        estimated_date = df.OrderCancelDate - pd.DateOffset(days=1)
        return adjust_to_weekday(estimated_date)
    
    # Ingram RELEASED
    elif (df.SSR_Row == 'Ingram') and (df.OrderCancelDate is not None):
        estimated_date = df.OrderCancelDate + pd.DateOffset(days=7)
        return adjust_to_weekday(estimated_date)
    
    # General case for release date
    if (df.ReleaseDate is not None) and (df.ReleaseDate > todays_date):
        return adjust_to_weekday(df.ReleaseDate)
    
    elif (df.ReleaseDate is not None) and (df.ReleaseDate <= todays_date):
        return next_three_days()
    
    else:
        return pd.NaT

def get_released_softallocated(df):
    '''
    FILTER for the following OrderTypeCodes: 'RELEASED', 'SOFT ALLOCATED','RUSH EDI AM'
    '''
    df = df.loc[df.OrderTypeCode.isin(['RELEASED', 'SOFT ALLOCATED','RUSH EDI AM'])]
    df['EstimateDate'] = df.apply(calculate_est_ship_date_released, axis=1)
    return df

def main():
    df = upload_ho()
    df = get_released_softallocated(df)
    print(df.info())
    print(df.head())
    print()
    print('Rows with missing EstimateDate:')
    nat_rows = df[df['EstimateDate'].isna()]
    print(nat_rows)
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
        
if __name__ == '__main__':
    main()