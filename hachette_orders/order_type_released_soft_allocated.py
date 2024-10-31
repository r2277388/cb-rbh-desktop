import pandas as pd
from function import adjust_to_weekday,next_three_days
from list_provider import hbg_tier1,hbg_tier2,hbg_tier3,hbg_tier4,hbg_tier5,big_box_stores

def calculate_est_ship_date_released(df):
    todays_date = pd.Timestamp('today').normalize()
    
    ######################
    # PRE-OSD Rules
    ######################
    # PRE-OSD Rules for Big Box Stores
    if (df.osd > todays_date) and (df.STATE in hbg_tier1) and (df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=10) 
        return adjust_to_weekday(estimated_date)
    
    elif (df.osd > todays_date) and (df.STATE in hbg_tier2) and (df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=12) 
        return adjust_to_weekday(estimated_date)
    
    elif (df.osd > todays_date) and (df.STATE in hbg_tier3) and (df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=13) 
        return adjust_to_weekday(estimated_date)
    
    elif (df.osd > todays_date) and (df.STATE in hbg_tier4) and (df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=13) 
        return adjust_to_weekday(estimated_date)

    elif (df.osd > todays_date) and (df.STATE in hbg_tier5) and (df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=14) 
        return adjust_to_weekday(estimated_date)
    
    # PRE-OSD Rules for Not Big Box Stores
    if (df.osd > todays_date) and (df.STATE in hbg_tier1) and (~df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=10) 
        return adjust_to_weekday(estimated_date)
    
    elif (df.osd > todays_date) and (df.STATE in hbg_tier2) and (~df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=12) 
        return adjust_to_weekday(estimated_date)
    
    elif (df.osd > todays_date) and (df.STATE in hbg_tier3) and (~df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=13) 
        return adjust_to_weekday(estimated_date)
    
    elif (df.osd > todays_date) and (df.STATE in hbg_tier4) and (~df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=13) 
        return adjust_to_weekday(estimated_date)

    elif (df.osd > todays_date) and (df.STATE in hbg_tier5) and (~df.SSR_Row in big_box_stores):
        estimated_date =  df.osd - pd.DateOffset(days=14) 
        return adjust_to_weekday(estimated_date) 

######################
# POST OSD Rules
######################

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
    
    elif (df.SSR_Row == 'Amazon.com') and (df.WMSDoNotShipAfter is not None) and (df.STATE =='PA'):
        estimated_date = df.WMSDoNotShipAfter - pd.DateOffset(days=1)
        return adjust_to_weekday(estimated_date)
    
    elif (df.SSR_Row == 'Amazon.com') and (df.WMSDoNotShipAfter is not None) and (df.STATE == 'TN'):
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