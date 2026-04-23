import pandas as pd
import numpy as np
from load_all import load_all_data
from functions import to_saturday

# Load the data
data = load_all_data(force_refresh=False)
df_orders= data["orders"]

target_isbn = '9781452179612'

## CONVERT ORDER BY SATURDAY DATES
df_orders['EnteredDateSaturday'] = df_orders['EnteredDate'].apply(to_saturday)
df_orders['ReleaseDateSaturday'] = df_orders['ReleaseDate'].apply(to_saturday)
df_orders['OrderCancelDateSaturday'] = df_orders['OrderCancelDate'].apply(to_saturday)

############################################
# ✅ Define conditions
conditions_release = [
    # Release Orders
    (df_orders["OrderTypeCode"] == "RELEASED") & (df_orders["OrderCancelDate"].isna())
    (df_orders["OrderTypeCode"] == "RELEASED") & (df_orders["OrderCancelDate"] <= pd.Timestamp.today())
    (df_orders["OrderTypeCode"] == "RELEASED") & (df_orders["OrderCancelDate"] > pd.Timestamp.today())
]

values_release = [
    # Released Order Values
    to_saturday(pd.Timestamp.today()),
    to_saturday(pd.Timestamp.today()),
    df_orders["OrderCancelDateSaturday"]
    ]
############################################
conditions_regular = [
    # Regular Orders
    # ORDERTYPECODE and RELEASEDATES are NULL
    (df_orders["OrderTypeCode"] == "REGULAR") #1
        & (df_orders["OrderCancelDate"].isna()) 
        & (df_orders["ReleaseDate"].isna())
        & (df_orders["EnteredDate"] < pd.Timestamp.today()-pd.DateOffset(months=1)),
    
    (df_orders["OrderTypeCode"] == "REGULAR") #2
        & (df_orders["OrderCancelDate"].isna()) 
        & (df_orders["ReleaseDate"].isna())
        & (df_orders["EnteredDate"] >= pd.Timestamp.today()-pd.DateOffset(months=1)),    
    
    # ORDERTYPECODE IS NULL and RELEASEDATES NOT NULL
    (df_orders["OrderTypeCode"] == "REGULAR") #3
        & (df_orders["OrderCancelDate"].isna()) 
        & (df_orders["ReleaseDate"]< pd.Timestamp.today()),
        
    (df_orders["OrderTypeCode"] == "REGULAR") #4
        & (df_orders["OrderCancelDate"].isna()) 
        & (df_orders["ReleaseDate"]>= pd.Timestamp.today()),
    
    
    (df_orders["OrderTypeCode"] == "REGULAR") #5
        & (df_orders["OrderCancelDate"].isna()) 
        & (df_orders["ReleaseDate"]> pd.Timestamp.today()),

    (df_orders["OrderTypeCode"] == "REGULAR") #6
        & (df_orders["ReleaseDate"].isna()) 
        & (df_orders["OrderCancelDate"]< pd.Timestamp.today()),

    ]

values_regular = [
    # Regular Order Values
    np.nat, #1
    to_saturday(pd.Timestamp.today()+pd.DateOffset(day=7)), #2
    np.nat, #3  
    df_orders["ReleaseDateSaturday"], #4
    to_saturday(pd.Timestamp.today()+pd.DateOffset(day=7)), #5
    df_orders["ReleaseDateSaturday"], #6
    ]
############################################

# # ✅ Apply conditions to create 'ProjectedShip'
# df_orders["ProjectedShip"] = np.select(conditions, values, default=np.nan)  # Default to NaN if no condition is met

# df_orders['ProjectedShip'] = df_orders[df_orders['OrderTypeCode']=='REGULAR']

## Orders for a single ISBN
df_orders = df_orders.loc[df_orders['ISBN']==target_isbn].copy()

print(df_orders.head())