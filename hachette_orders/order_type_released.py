import pandas as pd
from function import adjust_to_weekday,next_three_days

def calculate_est_ship_date_released(df):
    todays_date = pd.Timestamp('today').normalize()
    # Future OSD titles
    # Big Box Stores
    big_box_stores = list('Readerlink','Barnes & Noble','Amazon.com','Target Direct','Ingram')
    
    # Rules for Big Box Stores
    if (df.osd > todays_date) and (df.STATE in ['CA','WA','OR','NV']) and (df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=14) 
        return adjust_to_weekday(estimated_date)
    
    elif (df.osd > todays_date) and (df.STATE in ['CA','WA']) and (df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=6) 
        return adjust_to_weekday(estimated_date)
    
    # Rules for Not Big Box Stores
    elif (df.osd > todays_date) and (df.STATE in ['CA','WA']) and (~df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=6) # 5 day shipping
        return adjust_to_weekday(estimated_date)
    
    elif (df.osd > todays_date) and (df.STATE in ['AL','NY','NJ','PA','WI']) and (~df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=3) # 5 day shipping
        return adjust_to_weekday(estimated_date) 
    
    # READERLINK RELEASED
    elif (df.SSR_Row == 'Readerlink') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE == 'IL'):
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
    
    # BARNES & NOBLE RELEASED 
    elif (df.SSR_Row == 'Barnes & Noble') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE == 'NJ'):
        estimated_date = df.WMSDoNotDeliverAfter - pd.DateOffset(days=2) # 2 day shipping
        return adjust_to_weekday(estimated_date)
    elif (df.SSR_Row == 'Barnes & Noble') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE == 'NV'):
        estimated_date = df.WMSDoNotDeliverAfter - pd.DateOffset(days=4) # 4 day shipping to Nevada warehouse
        return adjust_to_weekday(estimated_date)
    
    elif (df.SSR_Row == 'Paper Source') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE == 'IL'):
        estimated_date = df.WMSDoNotDeliverAfter - pd.DateOffset(days=1) # 1 day shipping to Nevada warehouse
        return adjust_to_weekday(estimated_date)
    
    # Amazon RELEASED
    elif (df.SSR_Row == 'Amazon.com') and (df.WMSDoNotDeliverAfter is not None) and (df.STATE == 'IN'):
        estimated_date = df.WMSDoNotDeliverAfter
        return adjust_to_weekday(estimated_date)
    
    elif (df.SSR_Row == 'Amazon.com') and (df.WMSDoNotShipAfter is not None) and (df.STATE in ['IN', 'PA']):
        estimated_date = df.WMSDoNotShipAfter - pd.DateOffset(days=1)
        return adjust_to_weekday(estimated_date)
    
    elif (df.SSR_Row == 'Amazon.com') and (df.WMSDoNotShipAfter is not None) and (df.STATE in ['TN']):
        estimated_date = df.WMSDoNotShipAfter - pd.DateOffset(days=2)
        return adjust_to_weekday(estimated_date)
    
    elif (df.SSR_Row == 'Amazon.com') and (df.WMSDoNotShipAfter is not None) and (df.STATE == 'CA'):
        estimated_date = df.WMSDoNotShipAfter - pd.DateOffset(days=4)
        return adjust_to_weekday(estimated_date)
    
    elif (df.SSR_Row == 'Amazon.com') and (df.OrderCancelDate <= todays_date):
        return adjust_to_weekday(df.OrderCancelDate)
    
    # Target Direct RELEASED
    elif (df.SSR_Row == 'Target Direct') and (df.OrderCancelDate is not None):
        estimate_date = df.OrderCancelDate - pd.DateOffset(days=1)
        return adjust_to_weekday(estimate_date)
    
    # Ingram RELEASED
    elif (df.SSR_Row == 'Ingram') and (df.OrderCancelDate is not None):
        # Convo with Jeff, order to ship when enough quantity is collected so shipment is large.
        estimated_date = df.OrderCancelDate + pd.DateOffset(days=7)
        return adjust_to_weekday(estimated_date)
    
    
    # Everything else
    else:
        return adjust_to_weekday(df.release_date)