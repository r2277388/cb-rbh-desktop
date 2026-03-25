SELECT TOP 15
    sbn.[WEEK]
    ,count(sbn.week) Row_Cnt
    ,sum(sbn.BarnesAndNoble) AS BarnesAndNoble
    ,sum(sbn.BarnesAndNobleTotal) AS BarnesAndNobleTotal
    ,sum(sbn.BNcom) AS BarnesAndNobelCOM
    ,sbn.FILENAME
FROM
    [CBQ2].[cb].[Sellthrough_Barnes_and_Noble] sbn
    INNER JOIN ebs.item i on i.ITEM_TITLE = sbn.ISBN13
WHERE
    i.PUBLISHER_CODE NOT IN ('Benefit', 'AFO LLC', 'Glam Media', 'PQ Blackwell','PRINCETON','AMMO Books'
    ,'San Francisco Art Institute','FareArts','Sager','In Active','Driscolls','Impossible Foods','Moleskine')
GROUP BY
    sbn.[WEEK]
    ,sbn.FILENAME
ORDER BY [WEEK] DESC;
