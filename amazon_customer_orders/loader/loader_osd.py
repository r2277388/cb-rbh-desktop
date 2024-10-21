import pandas as pd
import numpy as np
from sqlalchemy import create_engine

# SQL query to retrieve item data
def osd_sql():
    return '''
    Select
        tt.ean13 ISBN
        ,tt.active_datevalue osd
    from tmm.cb_Import_Title_Tasks tt
    Where
        tt.date_desc = 'On Sale Date'
        AND tt.active_datevalue is not null
        AND tt.printingnumber = 1
    '''

def get_connection():
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def upload_osd() -> pd.DataFrame:
    engine = get_connection()
    with engine.connect() as connection:
        df = pd.read_sql_query(osd_sql(), connection)
    return df

def main():
    df = upload_osd()
    print(df.info())
    print(df.head())
    
if __name__ == '__main__':
    main()