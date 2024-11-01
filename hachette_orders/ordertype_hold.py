import pandas as pd
import random
from function import adjust_to_weekday
from list_provider import bn_and_ps_shipping_days, readerlink_shipping_days, bn_list

def calculate_est_ship_date_hold(df):
    todays_date = pd.Timestamp('today').normalize()
    
    if (df.HoldReason == 'SPECIAL PACK') and pd.notnull(df.ReleaseDate):
        # Jeff said the SPECIAL PACK orders will be shipped 2-5 days after the ReleaseDate
        numbers = [2, 3, 4, 5]
        selected_number = random.choices(numbers, weights=[1, 1, 1, 1], k=1)[0]
        estimated_date = df.ReleaseDate + pd.DateOffset(days=selected_number)
        return adjust_to_weekday(estimated_date)

    # Readerlink Hold
    
    ###### FIXE THIS ##### the case where Reprints are not null
    elif (df.HoldReason == 'HOT TITLE') and (df.SSR_Row == 'Readerlink') \
        and pd.notnull(df.WMSDoNotDeliverAfter) and pd.notnull(df.ReprintDate):
        if (df.WMSDoNotDeliverAfter <= todays_date):
            numbers = [1, 2, 3]
            selected_number = random.choices(numbers, weights=[1, 1, 1], k=1)[0]
            return adjust_to_weekday(todays_date + pd.DateOffset(days=selected_number))
        elif (df.WMSDoNotDeliverAfter > todays_date) and (df.STATE in readerlink_shipping_days):
            days_offset = readerlink_shipping_days[df.STATE]
            estimated_date = df.WMSDoNotDeliverAfter - pd.DateOffset(days=days_offset)
            return adjust_to_weekday(estimated_date)
        else:
            estimated_date = df.WMSDoNotDeliverAfter - pd.DateOffset(days=2)
            return adjust_to_weekday(estimated_date)

    # B&N and Paper Source Hold
    elif (df.HoldReason == 'ACCOUNT') and (df.SSR_Row.isin(bn_list)) and pd.notnull(df.WMSDoNotDeliverAfter):
        if df.WMSDoNotDeliverAfter <= todays_date:
            numbers = [1, 2, 3]
            selected_number = random.choices(numbers, weights=[1, 1, 1], k=1)[0]
            return adjust_to_weekday(todays_date + pd.DateOffset(days=selected_number))
        elif (df.WMSDoNotDeliverAfter > todays_date) and (df.STATE in bn_and_ps_shipping_days):
            days_offset = bn_and_ps_shipping_days[df.STATE]
            estimated_date = df.WMSDoNotDeliverAfter - pd.DateOffset(days=days_offset)
            return adjust_to_weekday(estimated_date)
        else:
            estimated_date = df.WMSDoNotDeliverAfter - pd.DateOffset(days=2)
            return adjust_to_weekday(estimated_date)
        
    else:
        return adjust_to_weekday(df.release_date)