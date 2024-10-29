import pandas as pd
from function import adjust_to_weekday

def calculate_est_ship_date_backordered(df):
    # READERLINK regular
    if (df.SSR_Row == 'Readerlink') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE == 'IL'): 
        estimated_date =  df.WMSDoNotDeliverAfter + pd.DateOffset(days=6) # 1 day shipping
        return adjust_to_weekday(estimated_date)
    
    


    else:
        return adjust_to_weekday(df.release_date)