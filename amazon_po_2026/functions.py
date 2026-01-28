# Aggregate the po doc
import pandas as pd
import numpy as np
from datetime import datetime as dt

def merge_cat_po(df_po, df_cat)-> pd.DataFrame:
    return pd.merge(df_po, df_cat, how='left', on='ISBN').fillna(0)


def create_dp_list(col)-> list:
    dp_list = list(col)
    if 0 in dp_list:
        dp_list.remove(0)
    if 'Chronicle' in dp_list:
        dp_list.remove('Chronicle')
    return dp_list

def filter_dp(df, dp)-> pd.DataFrame:
    return df.loc[(df['Publisher'] == dp)]

def dp_by_pgrp(df)-> pd.DataFrame:
    return df.groupby(['pgrp']).agg(Sum_of_Total_Cost=pd.NamedAgg(column='Total accepted cost', aggfunc=sum))

def dp_top20(df, col)-> pd.DataFrame:
    df_dp = df.groupby(
        ['ASIN', 'ISBN', 'title', 'pub', 'Reprint Due Date', 'Reprint Quantity', 'pgrp']).agg(
        Quantity_Requested=pd.NamedAgg(column='Requested quantity', aggfunc=sum),
        Quantity_Accepted=pd.NamedAgg(column='Accepted Quantity', aggfunc=sum),
        Total_Cost=pd.NamedAgg(column='Total accepted cost', aggfunc=sum)
    )
    return df_dp.sort_values(by=col, ascending=False).head(20)