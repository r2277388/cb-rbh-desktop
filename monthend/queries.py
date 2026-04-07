def barnes_noble_monthly_coop_sql(period: str) -> str:
    return f"""
SELECT
    sd.TRX_DATE [Invoice Date]
    ,sd.TRX_NUMBER [Invoice Number]
    ,sd.CUSTOMER_PO_NUMBER [PO Number]
    ,i.PUBLISHER_CODE [Publisher]
    ,i.PUBLISHING_GROUP [pgrp]
    ,i.ISBN
    ,i.SHORT_TITLE [Title]
    ,i.PRODUCT_TYPE [PT]
    ,i.FORMAT_DESCRIPTION [Format]
    ,i.PRICE_AMOUNT [Retail Price]
    ,sum(CASE WHEN sd.INVOICE_LINE_TYPE = 'SALE' THEN sd.QUANTITY_INVOICED ELSE 0 END) [Gross Units]
    ,sum(CASE WHEN sd.INVOICE_LINE_TYPE = 'RETURN' THEN sd.QUANTITY_INVOICED ELSE 0 END) [Return Units]
    ,sum(sd.QUANTITY_INVOICED) [Net Units]
    ,sum(CASE WHEN sd.INVOICE_LINE_TYPE = 'SALE' THEN sd.REVENUE_AMOUNT ELSE 0 END) [Gross Sales]
    ,sum(CASE WHEN sd.INVOICE_LINE_TYPE = 'RETURN' THEN sd.REVENUE_AMOUNT ELSE 0 END) [Return Sales]
    ,sum(sd.REVENUE_AMOUNT) [Net Sales]
    ,sum(sd.QUANTITY_INVOICED * i.PRICE_AMOUNT) [Net Retail Sales]
FROM
    ebs.sales as sd
    inner join ebs.Customer shipto on shipto.SITE_USE_ID = sd.SHIP_TO_SITE_USE_ID
    left join ebs.customer billto on billto.SITE_USE_ID = shipto.BILL_TO_SITE_USE_ID
    inner join ebs.Item as i on sd.ITEM_ID = i.ITEM_ID
WHERE
    sd.PERIOD = '{period}'
    AND cbq2.dbo.fnSaleTypeCode(SD.AR_TRX_TYPE_ID) = 'N'
    AND billto.MAIN_CUSTOMER = '03'
    AND i.PRODUCT_TYPE IN('BK', 'FT')
    AND sd.INVOICE_LINE_TYPE IN('SALE', 'RETURN')
GROUP BY
    sd.TRX_DATE
    ,sd.TRX_NUMBER
    ,sd.CUSTOMER_PO_NUMBER
    ,i.PUBLISHER_CODE
    ,i.PUBLISHING_GROUP
    ,i.ISBN
    ,i.SHORT_TITLE
    ,i.PRODUCT_TYPE
    ,i.FORMAT_DESCRIPTION
    ,i.PRICE_AMOUNT
"""
