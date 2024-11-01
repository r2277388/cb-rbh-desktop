import pandas as pd
import numpy as np
from sqlalchemy import create_engine

# SQL query to retrieve item data
def ho_sql():
    return '''                                       
    WITH osd as (Select
        tt.ean13 ISBN
        ,tt.active_datevalue osd
    from tmm.cb_Import_Title_Tasks tt
    Where
        tt.date_desc = 'On Sale Date'
        AND tt.active_datevalue is not null
        AND tt.printingnumber = 1)
                                        
    SELECT                                   
        ho.PONumber
        ,chan.Description channel
        ,subchan.Description SubChan
        ,ssr_row.Description SSR_Row
        ,shipto.STATE
        ,ho.ISBN
        ,i.SHORT_TITLE Title
        ,i.PRICE_AMOUNT Price
        ,case
                when i.PUBLISHING_GROUP in('FWN','LIF') then 'FLS'
                else i.PUBLISHING_GROUP
        end pgrp
        ,i.SEASON
        ,i.AMORTIZATION_DATE pubdate
        ,osd.osd
        ,case
                when ho.Discount is null then 0.54
                when ho.Discount = 0 then 0.54
                else Discount
        end discount
        ,ho.EnteredDate
        ,ho.ReleaseDate
        ,ho.OrderCancelDate
        ,ho.OrderTypeCode
        ,ho.PONumber
        ,ho.WMSDoNotDeliverAfter
        ,ho.WMSDoNotDeliverBefore
        ,ho.WMSDoNotShipAfter
        ,ho.WMSDoNotShipBefore
        ,sum(ho.Quantity) qty
    FROM                                           
        hachette.HachetteOrders ho
        inner join ebs.item i on i.ITEM_TITLE = ho.isbn     
        inner join ssr.SSRRow ssr_row on ssr_row.SSRRowID = ho.SSRRowID
        inner join ssr.SubChannel subchan on subchan.SubChannelID = ssr_row.SubChannelID
        inner join ssr.Channel chan on chan.ChannelID = subchan.ChannelID        
        left join ebs.Customer shipto on shipto.PARTYSITENUMBER = ho.StoreNumber
        left join osd on ho.ISBN = osd.ISBN
    WHERE                        
            i.PUBLISHER_CODE = 'Chronicle'
            AND i.PUBLISHING_GROUP not IN('MKT')                 
            and ho.EnteredDate > (GETDATE() -180)
            and i.PRICE_AMOUNT > 0
    GROUP BY                                             
        ho.PONumber
        ,chan.Description
        ,subchan.Description
        ,ssr_row.Description
        ,shipto.STATE
        ,ho.ISBN
        ,i.SHORT_TITLE
        ,i.PRICE_AMOUNT
        ,case
                when i.PUBLISHING_GROUP in('FWN','LIF') then 'FLS'
                else i.PUBLISHING_GROUP
        end
        ,i.SEASON
        ,i.AMORTIZATION_DATE
        ,osd.osd
        ,case
                when ho.Discount is null then .54
                when ho.Discount = 0 then 0.54
                else Discount
        end
        ,ho.EnteredDate
        ,ho.ReleaseDate
        ,ho.OrderCancelDate
        ,ho.OrderTypeCode
        ,ho.WMSDoNotDeliverAfter
        ,ho.WMSDoNotDeliverBefore
        ,ho.WMSDoNotShipAfter
        ,ho.WMSDoNotShipBefore
    ORDER BY
        osd.osd,ssr_row.Description  asc
    '''

def get_connection():
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def remove_consignment(df):
    countries_to_exclude = ['Australia', 'Canada']
    df = df[~df['SSR_Row'].isin(countries_to_exclude)]
    return df

def remove_rows(df):
    rows_to_exclude = ['Author/Individual']
    df = df[~df['SSR_Row'].isin(rows_to_exclude)]
    return df

def strip_spaces(df):
    """
    Strips leading and trailing spaces from specified columns in a DataFrame.
    
    Parameters:
    df (pd.DataFrame): The DataFrame to process.
    columns (list): List of column names to strip spaces from.
    
    Returns:
    pd.DataFrame: The modified DataFrame with stripped spaces in specified columns.
    """
    columns = ['OrderTypeCode','SSR_Row','STATE']
    for col in columns:
        if col in df.columns:
            df[col] = df[col].str.strip()
    return df

def create_val_column(df):
    df['val'] = round((1-df['discount']) * df['Price'] * df['qty'],2)
    return df

def upload_ho() -> pd.DataFrame:
    engine = get_connection()
    with engine.connect() as connection:
        df = pd.read_sql_query(ho_sql(), connection)
        df = remove_consignment(df)
        df = remove_rows(df)
        df = strip_spaces(df)
        df = create_val_column(df)
    return df

# Set pandas display option to show all columns
pd.set_option('display.max_columns', None)

def main():
    df = upload_ho()
    print(df.info())
    print(df.head())
    
if __name__ == '__main__':
    main()