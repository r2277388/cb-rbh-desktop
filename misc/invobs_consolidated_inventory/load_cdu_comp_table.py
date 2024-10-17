import pandas as pd
import pandas as pd
import numpy as np
from sqlalchemy import create_engine

# SQL query to retrieve item data
def cdu_comp_table():
    return '''
    SELECT               	
        p.Assembly_item cdu       
        ,p.COMPONENT component            
        ,p.Unit qty                
    FROM                 	
        ebs.Packs p          
        INNER JOIN ebs.item i on i.ISBN = p.Assembly_item                
    WHERE                	
        i.PUBLISHER_CODE = 'Chronicle'
        AND i.PUBLISHING_GROUP NOT IN('ZZZ','MKT')
    ORDER BY                   	
        p.Assembly_item

    '''

def get_connection():
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def upload_cdu_table() -> pd.DataFrame:
    engine = get_connection()
    with engine.connect() as connection:
        df = pd.read_sql_query(cdu_comp_table(), connection)
    return df


def main():
    df = upload_cdu_table()
    print(df.info())
    print(df.head())
    
if __name__ == "__main__":
    main()
