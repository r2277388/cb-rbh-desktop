import pandas as pd
import numpy as np

from loader.loader_weeklysales import uploader_weeklysales
from loader.loader_item import upload_item
from asin_isbn_converter import asin_isbn_conversion

def calculate_rev_units():
    df = uploader_weeklysales()
        
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
    # Grouping by 'div' and aggregating relevant columns
    summary = df.groupby('div').agg({
        'Ordered Revenue': 'sum',
        'Ordered Units': 'sum',
        'ou_pp': 'sum',
        'ou_ly': 'sum',
        'or_pp': 'sum',
        'or_ly': 'sum'
    }).reset_index()

    # Calculating percentage changes with division by zero handling
    summary['OR_PP_PCT'] = np.where(
        summary['Ordered Revenue'] == 0,
        '0%',
        ((summary['Ordered Revenue'] - summary['or_pp']) / summary['Ordered Revenue'] * 100).round(2).astype(str) + '%'
    )

    summary['OR_LY_PCT'] = np.where(
        summary['Ordered Revenue'] == 0,
        '0%',
        ((summary['Ordered Revenue'] - summary['or_ly']) / summary['Ordered Revenue'] * 100).round(2).astype(str) + '%'
    )

    summary['OU_PP_PCT'] = np.where(
        summary['Ordered Units'] == 0,
        '0%',
        ((summary['Ordered Units'] - summary['ou_pp']) / summary['Ordered Units'] * 100).round(2).astype(str) + '%'
    )

    summary['OU_LY_PCT'] = np.where(
        summary['Ordered Units'] == 0,
        '0%',
        ((summary['Ordered Units'] - summary['ou_ly']) / summary['Ordered Units'] * 100).round(2).astype(str) + '%'
    )

    # Calculating total row
    total = summary[['Ordered Revenue', 'Ordered Units', 'ou_pp', 'ou_ly', 'or_pp', 'or_ly']].sum()
    total['div'] = 'Total'

    # Handling division by zero in the total row
    total['OR_PP_PCT'] = (
        '0%' if total['Ordered Revenue'] == 0 
        else str(round((total['Ordered Revenue'] - total['or_pp']) / total['Ordered Revenue'] * 100, 2)) + '%'
        )

    total['OR_LY_PCT'] = (
        '0%' if total['Ordered Revenue'] == 0 
        else str(round((total['Ordered Revenue'] - total['or_ly']) / total['Ordered Revenue'] * 100, 2)) + '%'
        )

    total['OU_PP_PCT'] = (
        '0%' if total['Ordered Units'] == 0 
        else str(round((total['Ordered Units'] - total['ou_pp'] ) / total['Ordered Units'] * 100, 2)) + '%'
        )

    total['OU_LY_PCT'] = (
        '0%' if total['Ordered Units'] == 0 
        else str(round((total['Ordered Units'] - total['ou_ly']) / total['Ordered Units'] * 100, 2)) + '%'
        )

    # Convert total row to a DataFrame and concatenate with the summary
    total = pd.DataFrame([total])
    summary = pd.concat([summary, total], ignore_index=True)

    # Dropping intermediate columns no longer needed
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