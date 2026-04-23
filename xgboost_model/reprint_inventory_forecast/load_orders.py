from sqlalchemy import create_engine
import pandas as pd

def get_connection():
    """
    Establishes and returns a database connection using SQLAlchemy.
    """
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def fetch_hachette_orders(engine):
    """
    Runs the SQL query to fetch Hachette orders and returns the results as a Pandas DataFrame.
    """
    query = """                                             
    SELECT                                               
           ho.ISBN,
           ho.EnteredDate,
           ho.ReleaseDate,
           ho.OrderCancelDate,
           ho.OrderTypeCode,
           SUM(ho.Quantity) AS Orders
    FROM                                           
         hachette.HachetteOrders ho
         INNER JOIN ebs.item i ON i.ITEM_TITLE = ho.ISBN          
    WHERE                        
          i.PUBLISHER_CODE = 'Chronicle'
          AND i.PUBLISHING_GROUP NOT IN ('MKT')                 
          AND ho.EnteredDate > (GETDATE() - 180)
          AND i.PRICE_AMOUNT > 0
    GROUP BY                                             
            ho.ISBN,
            ho.EnteredDate,
            ho.ReleaseDate,
            ho.OrderCancelDate,
            ho.OrderTypeCode
    """
    
    return pd.read_sql_query(query, engine)

# ✅ Example usage in another script:
if __name__ == "__main__":
    engine = get_connection()
    df = fetch_hachette_orders(engine)
    
    print(df.info())  # Display DataFrame structure
    print(df.head())  # Display first few rows
