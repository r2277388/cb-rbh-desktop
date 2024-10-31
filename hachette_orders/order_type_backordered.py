import pandas as pd
from function import adjust_to_weekday
from list_provider import hbg_tier1,hbg_tier2,hbg_tier3,hbg_tier4,hbg_tier5,big_box_stores

def calculate_est_ship_date_backordered(df):
    todays_date = pd.Timestamp('today').normalize()
    # Future OSD titles
    # Big Box Stores    
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
    
    


    else:
        return adjust_to_weekday(df.release_date)