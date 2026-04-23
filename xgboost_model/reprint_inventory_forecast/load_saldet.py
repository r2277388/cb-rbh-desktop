import pandas as pd
from sqlalchemy import create_engine

def get_connection():
    """
    Establishes and returns a database connection using SQLAlchemy.
    """
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def fetch_sales_detail(engine):
    """
    Runs the SQL query to fetch Hachette orders and returns the results as a Pandas DataFrame.
    """
    query = """                                             
    WITH SalesLast2Years AS (
        SELECT  
            i.ISBN,
            SUM(sd.QUANTITY_INVOICED) AS total_qty
        FROM ebs.Sales sd
        INNER JOIN ssr.SalesSSRRow stie 
            ON stie.CUSTOMER_TRX_LINE_ID = sd.CUSTOMER_TRX_LINE_ID
        INNER JOIN ssr.SSRRow ssr_row 
            ON ssr_row.SSRRowID = stie.SSRRowID  
        INNER JOIN ebs.Item i 
            ON sd.ITEM_ID = i.ITEM_ID
        WHERE 
            sd.TRX_DATE >= DATEADD(YEAR, -2, GETDATE()) -- Last 2 years
            AND CBQ2.dbo.fnSaleTypeCode(sd.AR_TRX_TYPE_ID) = 'N' 
            AND sd.INVOICE_LINE_TYPE = 'SALE'
            AND i.PRODUCT_TYPE IN ('BK', 'FT')
            AND ssr_row.SSRRowID NOT IN ('100', '101') -- Excludes off-price purchases
            AND i.PUBLISHER_CODE = 'Chronicle'
            AND i.AMORTIZATION_DATE <= DATEADD(MONTH, -25, GETDATE())
        GROUP BY i.ISBN
        HAVING SUM(sd.QUANTITY_INVOICED) > 500 -- Filter out titles below 500
    )

    SELECT  
        -- Compute the next Saturday for each transaction date
        DATEADD(DAY, (7 - DATEPART(WEEKDAY, sd.TRX_DATE)), sd.TRX_DATE) AS [Date], 
        i.ISBN,
        SUM(sd.QUANTITY_INVOICED) AS qty
    FROM ebs.Sales sd
    INNER JOIN ssr.SalesSSRRow stie 
        ON stie.CUSTOMER_TRX_LINE_ID = sd.CUSTOMER_TRX_LINE_ID
    INNER JOIN ssr.SSRRow ssr_row 
        ON ssr_row.SSRRowID = stie.SSRRowID  
    INNER JOIN ebs.Item i 
        ON sd.ITEM_ID = i.ITEM_ID
    INNER JOIN SalesLast2Years s2y -- ✅ Join to only include qualified ISBNS
        ON i.ISBN = s2y.ISBN
    WHERE 
        LEFT(sd.PERIOD, 4) >= '2020'
        AND CBQ2.dbo.fnSaleTypeCode(sd.AR_TRX_TYPE_ID) = 'N' 
        AND sd.INVOICE_LINE_TYPE = 'SALE'
        AND i.PRODUCT_TYPE IN ('BK', 'FT')
        AND ssr_row.SSRRowID NOT IN ('100', '101') -- Excludes off-price purchases
        AND i.PUBLISHER_CODE = 'Chronicle'
        AND i.AMORTIZATION_DATE <= DATEADD(MONTH, -25, GETDATE())

    GROUP BY  
        DATEADD(DAY, (7 - DATEPART(WEEKDAY, sd.TRX_DATE)), sd.TRX_DATE), 
        i.ISBN;

    """
    return pd.read_sql_query(query
                             ,engine
                             ,parse_dates = ['Date']
                            ,dtype={'ISBN':'str','qty':'int64'})

if __name__ == '__main__':
    engine = get_connection()
    df = fetch_sales_detail(engine)
    print(df.info())
    print(df.head())


