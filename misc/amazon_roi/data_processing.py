import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from loader_ams import upload_ams
from loader_sellthru import sellthru_preparation
from loader_ebs_item import upload_item

df_item = upload_item()
df_ams=upload_ams()
df_sellthru=sellthru_preparation(use_cache = True)

def data_prep():
    df_merged = pd.merge(df_sellthru,df_ams,on=['ISBN','Period'],how='left')

    # This how = "inner" will filter out all non-core. I'm using the df_item
    # to filter publishing_codes or anything else. So focus is on the ebs.item SQL 
    # to do the filtering as needed.
    df_merged = pd.merge(df_merged,df_item,on='ISBN',how='inner')

    df_merged['AMS Spend'] = df_merged['AMS Spend'].fillna(0)

    # Create a binary column indicating whether a title was advertised by AMS
    df_merged['AMS Advertised'] = df_merged['AMS Spend'].apply(lambda x: True if x > 0 else False)

    # Convert Period to datetime format
    df_merged['Period'] = pd.to_datetime(df_merged['Period'],format='%Y%m')
    
    # Extract the month from the Period and apply cyclical encoding
    df_merged['Month'] = df_merged['Period'].dt.month
    df_merged['Month_sin'] = np.sin(2 * np.pi * df_merged['Month'] / 12)
    df_merged['Month_cos'] = np.cos(2 * np.pi * df_merged['Month'] / 12)

    # Calculate the age of the title in months
    df_merged['Title Age (Months)'] = (df_merged['Period'].dt.year - df_merged['pub_date'].dt.year) * 12 + (df_merged['Period'].dt.month - df_merged['pub_date'].dt.month)

   # Last step, there are many nulls in the Returns area. So I'll make them zero I guess. :(
    df_merged.fillna(0,inplace=True)

    return df_merged
    
    
def main():
    
    df_merged = data_prep()
    print('Shape of Combined Data: ',df_merged.shape)
    print('-'*40)
    print('Searching for NaNs ',df_merged.isna().sum()) 
    print('-'*40)
    print('Dataframe Detail ',df_merged.info())
    print('-'*40)
    print('First 5 Rows',df_merged.head())
    

if __name__=='__main__':
    main()