SELECT
    i.ISBN,
    SUM(QUANTITY_INVOICED) AS FaireQty
FROM ebs.Sales sd
    INNER JOIN ebs.Item i ON sd.ITEM_ID = i.ITEM_ID
    INNER JOIN ebs.Customer shipto ON shipto.SITE_USE_ID = sd.SHIP_TO_SITE_USE_ID
    LEFT JOIN ebs.Customer billto ON billto.SITE_USE_ID = shipto.BILL_TO_SITE_USE_ID
WHERE
    sd.PERIOD >= '202510'
    AND CBQ2.dbo.fnSaleTypeCode(sd.AR_TRX_TYPE_ID) = 'N'
    AND i.PRODUCT_TYPE IN ('BK', 'FT')
    AND LEFT(i.SEASON, 4) = '2026'
    AND LEFT(billto.PARTYSITENUMBER, 8) = '10460194'
GROUP BY
    i.ISBN;
