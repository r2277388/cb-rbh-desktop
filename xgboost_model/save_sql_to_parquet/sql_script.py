def run_sql():
    sql_query = """
    SELECT 			
        sd.TRX_DATE [Date]
        ,ssr_row.Description ssr_row
        ,ssr_row.SSRRowID		
        ,i.ISBN		
        ,i.SHORT_TITLE Title
        ,SUM(sd.REVENUE_AMOUNT) AS val		
        ,sum(sd.QUANTITY_INVOICED) AS qty		        
    FROM 			
        ebs.Sales sd		
        inner join ssr.SalesSSRRow stie on stie.CUSTOMER_TRX_LINE_ID = sd.CUSTOMER_TRX_LINE_ID		
        inner join ssr.SSRRow ssr_row on ssr_row.SSRRowID= stie.SSRRowID  		
        INNER JOIN ebs.Item i ON sd.ITEM_ID = i.ITEM_ID		
    WHERE 			
        left(sd.PERIOD,4) >= '2020'
        AND CBQ2.dbo.fnSaleTypeCode(SD.AR_TRX_TYPE_ID) = 'N'		
        AND sd.INVOICE_LINE_TYPE = 'SALE'		
        AND i.PUBLISHER_CODE = 'Chronicle'
    GROUP BY 			
        sd.TRX_DATE
        ,ssr_row.Description
        ,ssr_row.SSRRowID		
        ,i.ISBN		
        ,i.SHORT_TITLE
    """
    return sql_query
