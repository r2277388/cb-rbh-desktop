SELECT
    ho.ISBN,
    SUM(ho.Quantity) AS Faire_OO_qty
FROM hachette.HachetteOrders ho
    INNER JOIN ebs.Item i ON i.ITEM_TITLE = ho.ISBN
    INNER JOIN ssr.SSRRow ssr_row ON ssr_row.SSRRowID = ho.SSRRowID
WHERE
    i.PUBLISHER_CODE = 'Chronicle'
    AND i.PUBLISHING_GROUP NOT IN ('MKT')
    AND ho.EnteredDate > (GETDATE() - 180)
    AND ssr_row.SSRRowID = '307'
    AND i.PRICE_AMOUNT > 0
    AND LEFT(i.SEASON, 4) = '2026'
GROUP BY
    ho.ISBN;
