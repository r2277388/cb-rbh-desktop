import pandas as pd
from function import adjust_to_weekday

def calculate_est_ship_date_released(df):
    # READERLINK RELEASED
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
        estimated_date =  df.WMSDoNotDeliverAfter + pd.DateOffset(days=3) # 4 day shipping
        return adjust_to_weekday(estimated_date)
    
    # PAPER SOURCE RELEASED 
    elif (df.SSR_Row == 'Barnes & Noble') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE == 'NJ'):
        return df.WMSDoNotDeliverAfter + pd.DateOffset(days=2) # 2 day shipping
    
    elif (df.SSR_Row == 'Barnes & Noble') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE == 'NV'):
        return df.WMSDoNotDeliverAfter + pd.DateOffset(days=2) # 4 day shipping to Nevada warehouse
    
    
    
    else:
        return df.release_date - pd.DateOffset(days=5)