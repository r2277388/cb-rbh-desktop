import pandas as pd
import numpy as np

from loader.loader_weeklysales import uploader_weeklysales
from loader.loader_item import upload_item
from asin_isbn_converter import asin_isbn_conversion

def calculate_rev_units():
    df = uploader_weeklysales()
    
    df['or_pp'] = np.where(
        df['Ordered Revenue - Prior Period'] == 0,
        df['Ordered Revenue'],  # If 'Ordered Revenue - Prior Period' is 0
        np.where(
            df['Ordered Revenue - Prior Period'].isnull(),
            0,  # If 'Ordered Revenue - Prior Period' is NaN
            np.divide(df['Ordered Revenue'], (1 + df['Ordered Revenue - Prior Period']))  # Else
        )
    )
    
    # Fix for 'Ordered Revenue - Same Period Last Year'
    df['or_ly'] = np.where(
        df['Ordered Revenue - Same Period Last Year'] == 0,
        df['Ordered Revenue'],  # If 'Ordered Revenue - Same Period Last Year' is 0
        np.where(
            df['Ordered Revenue - Same Period Last Year'].isnull(),
            0,  # If 'Ordered Revenue - Same Period Last Year' is NaN
            np.divide(df['Ordered Revenue'], (1 + df['Ordered Revenue - Same Period Last Year']))  # Else
        )
    )
    
    # Fix for 'Ordered Units - Prior Period'
    df['ou_pp'] = np.where(
        df['Ordered Units - Prior Period'] == 0,
        df['Ordered Units'],  # If 'Ordered Units - Prior Period' is 0
        np.where(
            df['Ordered Units - Prior Period'].isnull(),
            0,  # If 'Ordered Units - Prior Period' is NaN
            np.divide(df['Ordered Units'], (1 + df['Ordered Units - Prior Period']))  # Else
        )
    )
    
    # Fix for 'Ordered Units - Same Period Last Year'
    df['ou_ly'] = np.where(
        df['Ordered Units - Same Period Last Year'] == 0,
        df['Ordered Units'],  # If 'Ordered Units - Same Period Last Year' is 0
        np.where(
            df['Ordered Units - Same Period Last Year'].isnull(),
            0,  # If 'Ordered Units - Same Period Last Year' is NaN
            np.divide(df['Ordered Units'], (1 + df['Ordered Units - Same Period Last Year']))  # Else
        )
    )
    
    drop_list = ['Ordered Revenue - Prior Period','Ordered Revenue - Same Period Last Year'
                 ,'Ordered Units - Prior Period','Ordered Units - Same Period Last Year'
                 ]
    df.drop(columns=drop_list,inplace=True)
    df.dropna(inplace=True)
    return df

def div():
    df_converter = asin_isbn_conversion()
    df_rev_units = calculate_rev_units()
    df_converter = df_converter[['ASIN', 'ISBN']]
    
    df = df_rev_units.merge(df_converter,on='ASIN',how='inner')

    df_item = upload_item()
    df_item = df_item[['ISBN','publisher']]
    df = df.merge(df_item,on='ISBN',how='inner')
    
    df['div'] = np.where(df['publisher']=='Chronicle', "CB", "DP")
    
    df.drop_duplicates(subset=['ASIN'],inplace=True)
    
    df.drop(columns=['publisher','ISBN','ASIN'],inplace=True)
    
    return df

def summarize_by_div(df):
    summary = df.groupby('div').agg({
        'Ordered Revenue': 'sum',
        'Ordered Units': 'sum',
        'ou_pp': 'sum',
        'ou_ly': 'sum',
        'or_pp': 'sum',
        'or_ly': 'sum'
    }).reset_index()

    summary['OR_PP_PCT'] = ((summary['or_pp']-summary['Ordered Revenue']) / summary['Ordered Revenue'] * 100).round(2).astype(str) + '%'
    summary['OR_LY_PCT'] = ((summary['or_ly']-summary['Ordered Revenue']) / summary['Ordered Revenue'] * 100).round(2).astype(str) + '%'
    summary['OU_PP_PCT'] = ((summary['ou_pp']-summary['Ordered Units']) / summary['Ordered Units'] * 100).round(2).astype(str) + '%'
    summary['OU_LY_PCT'] = ((summary['ou_ly']-summary['Ordered Units']) / summary['Ordered Units'] * 100).round(2).astype(str) + '%'

    total = summary[['Ordered Revenue', 'Ordered Units', 'ou_pp', 'ou_ly', 'or_pp', 'or_ly']].sum()
    total['div'] = 'Total'
    total['OR_PP_PCT'] = str(round(total['or_pp'] / total['Ordered Revenue'] * 100, 2)) + '%'
    total['OR_LY_PCT'] = str(round(total['or_ly'] / total['Ordered Revenue'] * 100, 2)) + '%'
    total['OU_PP_PCT'] = str(round(total['ou_pp'] / total['Ordered Units'] * 100, 2)) + '%'
    total['OU_LY_PCT'] = str(round(total['ou_ly'] / total['Ordered Units'] * 100, 2)) + '%'

    total = pd.DataFrame(total).T
    summary = pd.concat([summary, total], ignore_index=True)
    summary.drop(columns=['ou_pp', 'ou_ly', 'or_pp', 'or_ly'], inplace=True)

    return summary  

def stats_cb():
    df = div()
    summary = summarize_by_div(df)
    return summary

def main():
    df = stats_cb()
    print(df.head())

if __name__ == "__main__":
    main()