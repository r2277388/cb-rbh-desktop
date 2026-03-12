SELECT DISTINCT TOP 10
    sbn.[WEEK]
    ,count(sbn.week) Row_Cnt
    ,sum(sbn.BarnesAndNoble) AS BarnesAndNoble
    ,sum(sbn.BarnesAndNobleTotal) AS BarnesAndNobleTotal
FROM [CBQ2].[cb].[Sellthrough_Barnes_and_Noble] sbn
GROUP BY
    sbn.[WEEK]
ORDER BY [WEEK] DESC;
