import pandas as pd
import numpy as np
from sqlalchemy import create_engine

# SQL query to retrieve item data
def ho_sql():
    return '''                                       
    SELECT                                           
        chan.Description channel
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
        ,case
                when ho.Discount is null then 0.54
                when ho.Discount = 0 then 0.54
                else Discount
        end discount
        ,ho.EnteredDate
        ,ho.ReleaseDate
        ,ho.OrderCancelDate
        ,ho.OrderTypeCode
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
    WHERE                        
        i.PUBLISHER_CODE = 'Chronicle'
        AND i.PUBLISHING_GROUP not IN('MKT')                 
        and ho.EnteredDate > (GETDATE() -180)
        and i.PRICE_AMOUNT > 0
        and ho.OrderTypeCode = 'RELEASED'
        AND ssr_row.Description = 'Readerlink'
    GROUP BY                                             
        chan.Description
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
        ho.ReleaseDate asc
    '''

def get_connection():
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def upload_ho() -> pd.DataFrame:
    engine = get_connection()
    with engine.connect() as connection:
        df = pd.read_sql_query(ho_sql(), connection)
    return df

def main():
    df = upload_ho()
    print(df.info())
    print(df.head())
    
if __name__ == '__main__':
    main()