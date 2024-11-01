import pandas as pd
from function import adjust_to_weekday, next_three_days,summarize_by_estimate_date

def calculate_est_ship_date_regular(df):
    todays_date = pd.Timestamp('today').normalize()
    
    ## READERLINK RELEASED
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
    
    # Barnes & Noble holds
    # BARNES & NOBLE OR PAPER SOURCE RELEASED 
    bn_and_ps_shipping_days = {
        'IL': 1,   # 1 day shipping
        'NJ': 2,   # 2 day shipping
        'NV': 4    # 4 day shipping
    }
    if (df.SSR_Row in ['Barnes & Noble','Barnes & Noble College', 'Paper Source']):
        if pd.notnull(df.WMSDoNotDeliverAfter):
            if df.STATE in bn_and_ps_shipping_days:
                if df.STATE in bn_and_ps_shipping_days:
                    days_offset = bn_and_ps_shipping_days[df.STATE]
                    estimated_date = df.WMSDoNotDeliverAfter - pd.DateOffset(days=days_offset)
                    return adjust_to_weekday(estimated_date)
            if df.STATE not in bn_and_ps_shipping_days:
                return adjust_to_weekday(df.WMSDoNotDeliverAfter) - pd.Dateoffset(day = 2)
        elif pd.notnull(df.ReleaseDate) and (df.ReleaseDate > todays_date):
            return adjust_to_weekday(df.ReleaseDate)
        elif pd.notnull(df.ReleaseDate) and (df.ReleaseDate <= todays_date):
            return next_three_days()
    
    # General case for release date
    if (df.ReleaseDate is not None) and (df.ReleaseDate > todays_date):
        return adjust_to_weekday(df.ReleaseDate)
    
    elif (df.ReleaseDate is not None) and (df.ReleaseDate <= todays_date):
        return next_three_days()
    
    else:
        return pd.NaT