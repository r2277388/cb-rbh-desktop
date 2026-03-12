SELECT DISTINCT TOP 10
    bc.[WEEK]
    ,count(bc.week) AS Row_Cnt
    ,sum(bc.Sales) AS Sales
FROM [CBQ2].[cb].[Sellthrough_RollBookscan] bc
GROUP BY
    bc.[WEEK]
ORDER BY [WEEK] DESC;