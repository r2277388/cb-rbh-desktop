from sqlalchemy import create_engine
import pandas as pd

# CDU Component Version Code

def get_connection():
    """
    Establishes and returns a database connection using SQLAlchemy.
    """
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def fetch_cdu_comps(engine):
    """
    Runs the SQL query to fetch Hachette orders and returns the results as a Pandas DataFrame.
    """
    query = """                                             
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
        AND p.Unit > 0
    ORDER BY
        cdu
    """
    
    return pd.read_sql_query(query, engine)

# ✅ Example usage in another script:
if __name__ == "__main__":
    engine = get_connection()
    df = fetch_cdu_comps(engine)
    
    print(df.info())  # Display DataFrame structure
    print(df.head())  # Display first few rows