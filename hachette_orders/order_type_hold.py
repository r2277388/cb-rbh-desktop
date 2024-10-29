import pandas as pd
from function import adjust_to_weekday

def calculate_est_ship_date_hold(df):
    # READERLINK regular
    if (df.SSR_Row == 'Readerlink') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE == 'IL'): 
        estimated_date =  df.WMSDoNotDeliverAfter + pd.DateOffset(days=6) # 1 day shipping
        return adjust_to_weekday(estimated_date)
    elif (df.SSR_Row == 'Readerlink') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE == 'VA'):  
        estimated_date =  df.WMSDoNotDeliverAfter + pd.DateOffset(days=5) # 2 day shipping
        return adjust_to_weekday(estimated_date)
    elif (df.SSR_Row == 'Readerlink') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE in ['TX','GA']):  
        estimated_date =  df.WMSDoNotDeliverAfter + pd.DateOffset(days=4) # 3 day shipping
        return adjust_to_weekday(estimated_date)
    elif (df.SSR_Row == 'Readerlink') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE == 'UT'):  
        estimated_date =  df.WMSDoNotDeliverAfter + pd.DateOffset(days=3)
        return adjust_to_weekday(estimated_date)
    
    # Barnes & Noble holds
    elif (df.SSR_Row == 'Barnes & Noble') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE == 'NJ'):
        estimated_date = df.WMSDoNotDeliverAfter - pd.DateOffset(days=2) # 2 day shipping
        return adjust_to_weekday(estimated_date)
    elif (df.SSR_Row == 'Barnes & Noble') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE == 'NV'):
        estimated_date = df.WMSDoNotDeliverAfter - pd.DateOffset(days=4) # 4 day shipping to Nevada warehouse
        return adjust_to_weekday(estimated_date)
    
    # Everything else
    else:
        return adjust_to_weekday(df.release_date)