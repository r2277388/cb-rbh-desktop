def sql_5y_sales() -> str:
    """
    Obtaining data from sql to get up-to-date meta data titles that are at least 5 years old
    and have had at least 100 sales during the current month.
    
    """
    
    return """
        
    DECLARE @last_full_month_date DATE = DATEADD(MONTH, -1, GETDATE())
    DECLARE @tp CHAR(6) = FORMAT(@last_full_month_date, 'yyyyMM');
    DECLARE @pp CHAR(6) = FORMAT(DATEADD(MONTH, -1, @last_full_month_date), 'yyyyMM');
    DECLARE @start_period CHAR(6) = FORMAT(DATEADD(MONTH, -60, @last_full_month_date), 'yyyyMM');

    WITH title_list AS (
        SELECT 
            i.ITEM_TITLE AS ISBN
        FROM
            ebs.sales AS sd
            INNER JOIN ebs.Item AS i ON sd.ITEM_ID = i.ITEM_ID
        WHERE 
            sd.PERIOD IN (@tp, @pp)
            AND cbq2.dbo.fnSaleTypeCode(sd.AR_TRX_TYPE_ID) = 'N'
            AND i.PRODUCT_TYPE IN ('bk', 'ft')
            AND sd.INVOICE_LINE_TYPE = 'SALE'
            AND i.PUBLISHER_CODE = 'Chronicle'
            AND i.AMORTIZATION_DATE < DATEADD(MONTH, -12, @last_full_month_date)  -- Optional logic tweak
        GROUP BY 
            i.ITEM_TITLE
        HAVING 
            SUM(sd.QUANTITY_INVOICED) > 100
        )

    SELECT 
        sd.PERIOD
        ,i.ITEM_TITLE AS ISBN
        ,CASE
            WHEN LEFT(i.PUBLISHING_GROUP, 3) = 'BAR' THEN 'BAR'
            ELSE i.PUBLISHING_GROUP
        END AS PGRP
        ,i.PRODUCT_TYPE PT
        ,i.AMORTIZATION_DATE PUB_DATE
        ,i.PRICE_AMOUNT PRICE
        ,SUM(sd.QUANTITY_INVOICED) AS QTY
    FROM
        ebs.sales AS sd
        INNER JOIN ebs.Item AS i ON sd.ITEM_ID = i.ITEM_ID
        INNER JOIN title_list AS tl ON tl.ISBN = i.ITEM_TITLE
    WHERE 
        sd.PERIOD BETWEEN @start_period AND @tp
        AND cbq2.dbo.fnSaleTypeCode(sd.AR_TRX_TYPE_ID) = 'N'
        AND i.PRODUCT_TYPE IN ('bk', 'ft')
        AND sd.INVOICE_LINE_TYPE = 'SALE'
        AND i.PUBLISHER_CODE = 'Chronicle'
        AND i.PRICE_AMOUNT IS NOT NULL
        AND i.AMORTIZATION_DATE IS NOT NULL
    GROUP BY 
        sd.PERIOD,
        i.ITEM_TITLE,
        CASE
            WHEN LEFT(i.PUBLISHING_GROUP, 3) = 'BAR' THEN 'BAR'
            ELSE i.PUBLISHING_GROUP
        END
        ,i.PRODUCT_TYPE
        ,i.AMORTIZATION_DATE
        ,i.PRICE_AMOUNT;
        """