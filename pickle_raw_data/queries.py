def query_saldet(period):
    '''
    Query obtains daily sales data for Core Chronicle.
    The sales that hit on the Saturday or Sunday reassigned to Friday so 
    that all sales are contained on Monday thru Friday.
    '''
    
    return f'''
    SELECT
        sd.period
        ,CASE
            WHEN DATEPART(WEEKDAY, sd.TRX_DATE) = 1 THEN DATEADD(DAY, -2, sd.TRX_DATE) -- Adjusting Sunday to Friday
            WHEN DATEPART(WEEKDAY, sd.TRX_DATE) = 7 THEN DATEADD(DAY, -1, sd.TRX_DATE) -- Adjusting Saturday to Friday
            ELSE sd.TRX_DATE
        END AS [ds]
        ,CASE
            WHEN LEFT(i.PUBLISHING_GROUP,3) = 'BAR' THEN 'BAR'
            WHEN i.PUBLISHER_CODE = 'Princeton' THEN 'CPA'
            ELSE i.PUBLISHING_GROUP
        END pgrp
        ,ssr_row.Description ssr
        ,CASE
            WHEN ssr_row.SSRRowID IN('32','146') then 'Consignment'
            WHEN ssr_row.SSRRowID IN('6') then 'Amazon'
            ELSE chan.Description
        END channel
        ,CASE
            WHEN [dbo].[fnFrontBackListCode](i.AMORTIZATION_DATE,sd.TRX_DATE) IN('A','R') THEN 'F'
            ELSE 'B'
        END flbl
        ,SUM(CASE 
                WHEN i.PUBLISHER_CODE = 'Princeton' AND YEAR(sd.TRX_DATE) > 2022 THEN 0 
                ELSE sd.REVENUE_AMOUNT 
            END) AS [y]
    FROM
        ebs.Sales sd
        INNER JOIN ssr.SalesSSRRow stie on stie.CUSTOMER_TRX_LINE_ID = sd.CUSTOMER_TRX_LINE_ID
        INNER JOIN ssr.SSRRow ssr_row on ssr_row.SSRRowID= stie.SSRRowID
        INNER JOIN ssr.SubChannel sub on sub.SubChannelID = ssr_row.SubChannelID
        INNER JOIN ssr.Channel chan on chan.ChannelID = sub.ChannelID
        INNER JOIN ebs.Item i ON i.ITEM_ID = sd.ITEM_ID
    WHERE
        sd.PERIOD >= '{period}'
        AND sd.INVOICE_LINE_TYPE = 'SALE'
        AND cbq2.dbo.fnSaleTypeCode(SD.AR_TRX_TYPE_ID) = 'N'
        AND i.PRICE_AMOUNT <> 0
        AND i.PUBLISHER_CODE IN('Chronicle','Princeton')
        AND i.PRODUCT_TYPE IN ('BK', 'FT')
    GROUP BY
        sd.period
        ,CASE 
            WHEN DATEPART(WEEKDAY, sd.TRX_DATE) = 1 THEN DATEADD(DAY, -2, sd.TRX_DATE) -- Adjusting Sunday to Friday
            WHEN DATEPART(WEEKDAY, sd.TRX_DATE) = 7 THEN DATEADD(DAY, -1, sd.TRX_DATE) -- Adjusting Saturday to Friday
            ELSE sd.TRX_DATE
        END
        ,CASE
            WHEN LEFT(i.PUBLISHING_GROUP,3) = 'BAR' THEN 'BAR'
            WHEN i.PUBLISHER_CODE = 'Princeton' THEN 'CPA'
            ELSE i.PUBLISHING_GROUP
        END
        ,ssr_row.Description
        ,CASE
            WHEN ssr_row.SSRRowID IN('32','146') then 'Consignment'
            WHEN ssr_row.SSRRowID IN('6') then 'Amazon'
            ELSE chan.Description
        END
        ,CASE
            WHEN [dbo].[fnFrontBackListCode](i.AMORTIZATION_DATE,sd.TRX_DATE) IN('A','R') THEN 'F'
            ELSE 'B'
    END
    '''