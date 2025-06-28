def sql_5y_sales() -> str:
    """
    Obtaining data from sql to get up-to-date meta data titles that are at least 5 years old
    and have had at least 100 sales during the current month.
    
    """
    
    return """
        
    WITH title_list as (
        SELECT 
        i.ITEM_TITLE ISBN
    FROM
        ebs.sales as sd 
        inner join ebs.Item as i on sd.ITEM_ID=i.ITEM_ID 
    WHERE 
        sd.PERIOD = '202506' 
        AND cbq2.dbo.fnSaleTypeCode(SD.AR_TRX_TYPE_ID) = 'N'
        and i.PRODUCT_TYPE in ('bk', 'ft') 
        AND sd.INVOICE_LINE_TYPE IN('SALE') 
        and i.PUBLISHER_CODE ='Chronicle'
        AND i.AMORTIZATION_DATE < '2020-06-01'
    GROUP BY 
        i.ITEM_TITLE
    HAVING
        sum(sd.QUANTITY_INVOICED) > 100
        )

    SELECT 
        sd.PERIOD
        ,i.ITEM_TITLE ISBN 
        ,i.PRODUCT_TYPE PT
        ,sum(sd.QUANTITY_INVOICED) qty 
    FROM
        ebs.sales as sd 
        inner join ebs.Item as i on sd.ITEM_ID=i.ITEM_ID 
        INNER JOIN title_list as tl on tl.ISBN = i.ITEM_TITLE
    WHERE 
        sd.PERIOD between '202001' and '202506' 
        AND cbq2.dbo.fnSaleTypeCode(SD.AR_TRX_TYPE_ID) = 'N'
        and i.PRODUCT_TYPE in ('bk', 'ft') 
        AND sd.INVOICE_LINE_TYPE IN('SALE') 
        and i.PUBLISHER_CODE = 'Chronicle'
    GROUP BY 
        sd.PERIOD
        ,i.ITEM_TITLE 
        ,i.PRODUCT_TYPE
        """