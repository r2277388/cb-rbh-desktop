def query_saldet(start_period, end_period):
    return f'''
    SELECT
        i.ISBN,
        DATEADD(DAY, -DATEDIFF(DAY, 0, sd.TRX_DATE) % 7, sd.TRX_DATE) AS WeekStartDate,
        SUM(sd.QUANTITY_INVOICED) AS qty
    FROM
        ebs.Sales sd
        INNER JOIN ebs.Item i ON i.ITEM_ID = sd.ITEM_ID
    WHERE 
        sd.PERIOD BETWEEN '{start_period}' AND '{end_period}'
        AND i.PRODUCT_TYPE IN('BK','FT','RP','CP')
        AND sd.INVOICE_LINE_TYPE = 'SALE'
        AND cbq2.dbo.fnSaleTypeCode(sd.AR_TRX_TYPE_ID) = 'N'
        AND i.PUBLISHER_CODE NOT IN('ZZZ','MKT')
        AND i.isbn IS NOT NULL
    GROUP BY
        i.ISBN,
        DATEADD(DAY, -DATEDIFF(DAY, 0, sd.TRX_DATE) % 7, sd.TRX_DATE)
    ORDER BY
        i.ISBN,
        WeekStartDate;

    '''