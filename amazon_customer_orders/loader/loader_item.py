import pandas as pd
import numpy as np
from sqlalchemy import create_engine

# SQL query to retrieve item data
def item_sql():
    return '''
    SELECT
        i.ISBN
        ,i.SHORT_TITLE title
        ,CASE
            WHEN left(i.PUBLISHING_GROUP,2) = 'HG' THEN 'Hardie Grant'
            WHEN left(i.PUBLISHING_GROUP,3) = 'QDP' THEN 'Quadrille'
            ELSE i.PUBLISHER_CODE
	    END publisher
        ,case                             
            when i.PUBLISHER_CODE = 'Tourbillon' then 'TW'                      
            when i.PUBLISHER_CODE = 'Sierra Club' then 'SC'                     
            when i.PUBLISHER_CODE IN('Glam Media','Benefit','PQ Blackwell','San Francisco Art Institute','AFO LLC','FareArts','Sager') then 'CD'                 
            when i.PUBLISHER_CODE = 'Creative Company' then 'CC'   
            when i.PUBLISHER_CODE = 'Do Books' then 'DO'
            when i.PUBLISHER_CODE = 'Levine Querido' then 'LQ'
            when i.PUBLISHER_CODE = 'AMMO Books' then 'AM'                                           
            when i.PUBLISHING_GROUP = 'GAL' then 'GA'                                                      
            when i.PUBLISHING_GROUP = 'GAL-CL' then 'CL'                        
            when i.PUBLISHING_GROUP = 'MUD' then 'MP'
            when i.PUBLISHING_GROUP = 'GAL-BM' then 'BM'             
            when i.PUBLISHING_GROUP in('LAU-BIS') then 'LKBS'                          
            when i.PUBLISHER_CODE = 'Laurence King' and i.PRODUCT_TYPE = 'FT' then 'LKGI'                      
            when i.PUBLISHER_CODE = 'Laurence King' and i.PRODUCT_TYPE <> 'FT' then 'LKBK'         
            when left(i.PUBLISHING_GROUP,2) = 'HG' then 'HG' 
            when left(i.PUBLISHING_GROUP,3) = 'QDP' then 'QD'
            when i.PUBLISHING_GROUP in('BAR-ART','BAR-ENT','BAR-LIF') then 'BAR'                  
            else i.PUBLISHING_GROUP                 
        end pgrp
    FROM                
        ebs.Item i
    WHERE
        i.SHORT_TITLE IS NOT NULL
        AND i.ISBN IS NOT NULL
        AND i.PUBLISHING_GROUP <> '???'
        AND i.PRODUCT_TYPE in('BK','FT','DI')
        --AND i.AVAILABILITY_STATUS not in('OP','WIT','OPR','NOP','OSI','PC','DIS','CS','POS')
        AND i.AVAILABILITY_STATUS is not null
    '''

def get_connection():
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def upload_item() -> pd.DataFrame:
    engine = get_connection()
    with engine.connect() as connection:
        df = pd.read_sql_query(item_sql(), connection)
    return df

def main():
    df = upload_item()
    print(df.info())
    print(df.head())
    
if __name__ == '__main__':
    main()