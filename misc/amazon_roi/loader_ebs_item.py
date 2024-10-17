import pandas as pd
from sqlalchemy import create_engine
from queries import db_item

# Create SQL Connection and Upload Item
def upload_item_raw() -> pd.DataFrame:
    # Create the engine inside the function
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    
    # Read doc
    with engine.connect() as connection:
        df = pd.read_sql_query(db_item(), connection)
        df['ISBN'] = df['ISBN'].astype(str)
        
    return df

def upload_item() -> pd.DataFrame:
    # Upload ebs.item
    df = upload_item_raw()
    
    # Filter for Chroniclebooks only
    df = df.loc[df['pub']=='Chronicle']
    
    # Droping some columns to make it less busy
    columns_to_drop = ['pub','pgrp','FT','RF','Title']
    df.drop(columns = columns_to_drop,inplace=True)

    return df


def main():
    # Call upload_item and store the result in df
    df = upload_item_raw()
    print(df.info())
    # # Print the head and info of the dataframe
    # print(df.head())
    # print("Item Data ", df.info())
    # print("-" * 60)
    
if __name__ == '__main__':
    main()