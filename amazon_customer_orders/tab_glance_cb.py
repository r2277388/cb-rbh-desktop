import pandas as pd
import numpy as np

from loader.loader_traffic import upload_traffic
from loader.loader_item import upload_item
from asin_isbn_converter import asin_isbn_conversion

def calculate_rev_units():
    df = upload_traffic()
    
    df['gv_pp'] = np.where(
        df['Glance Views - Prior Period'] == 0,
        df['Glance Views'],  # If 'Ordered Revenue - Prior Period' is 0
        np.where(
            df['Glance Views - Prior Period'].isnull(),
            0,  # If 'Ordered Revenue - Prior Period' is NaN
            np.divide(df['Glance Views'], (1 + df['Glance Views - Prior Period']))  # Else
        )
    )
    
    # Fix for 'Ordered Revenue - Same Period Last Year'
    df['gv_ly'] = np.where(
        df['Glance Views - Same Period Last Year'] == 0,
        df['Glance Views'],  # If 'Ordered Revenue - Same Period Last Year' is 0
        np.where(
            df['Glance Views - Same Period Last Year'].isnull(),
            0,  # If 'Ordered Revenue - Same Period Last Year' is NaN
            np.divide(df['Glance Views'], (1 + df['Glance Views - Same Period Last Year']))  # Else
        )
    )
    
    df_converter = asin_isbn_conversion()
    df_converter = df_converter[['ASIN','ISBN']]
    
    df = df.merge(df_converter,on='ASIN',how='inner')
    
    df_item = upload_item()
    df_item = df_item[['ISBN','publisher']]
    df = df.merge(df_item,on='ISBN',how='inner')
    
    df['div'] = np.where(df['publisher']=='Chronicle','CB','DP')
    
    df.drop_duplicates(subset=['ASIN','ISBN'],inplace=True)
    
    drop_list = ['Glance Views - Prior Period','Glance Views - Same Period Last Year'\
        ,'publisher','ASIN','ISBN']
    df.drop(columns=drop_list,inplace=True)
    df.dropna(inplace=True)
    
    return df

def format_percentage(numerator, denominator):
    return str(round((numerator-denominator) / denominator * 100, 2)) + '%'
 
def summarize_data(df):
    summary = df.groupby('div').agg({
        'Glance Views': 'sum',
        'gv_pp': 'sum',
        'gv_ly': 'sum'
    }).reset_index()

    summary['gv_pp_pct'] = summary.apply(lambda row: format_percentage(row['gv_pp'], row['Glance Views']), axis=1)
    summary['gv_lp_pct'] = summary.apply(lambda row: format_percentage(row['gv_ly'], row['Glance Views']), axis=1)

    total = summary[['Glance Views', 'gv_pp', 'gv_ly']].sum()
    total['div'] = 'Total'
    total['gv_pp_pct'] = format_percentage(total['gv_pp'], total['Glance Views'])
    total['gv_lp_pct'] = format_percentage(total['gv_ly'], total['Glance Views'])

    total = pd.DataFrame([total])
    summary = pd.concat([summary, total], ignore_index=True)
    summary.drop(columns=['gv_pp', 'gv_ly'], inplace=True)

    # Reorder columns if necessary
    summary = summary[['div', 'Glance Views', 'gv_pp_pct', 'gv_lp_pct']]

    return summary


def glance_views():
    df = calculate_rev_units()
    df = summarize_data(df)
    return df

def main():
    df = glance_views()
    print(df.info())
    print(df.head())
    
if __name__ == "__main__":
    main()