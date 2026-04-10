from __future__ import annotations


PUBLISHER_EXCLUSION_SQL = """
    AND i.PUBLISHER_CODE IS NOT NULL
    AND i.PUBLISHER_CODE NOT IN (
        'Benefit',
        'AFO LLC',
        'Glam Media',
        'PQ Blackwell',
        'PRINCETON',
        'AMMO Books',
        'San Francisco Art Institute',
        'FareArts',
        'Sager',
        'In Active',
        'Driscolls',
        'Impossible Foods',
        'Moleskine'
    )
    AND i.PRODUCT_TYPE IN ('BK', 'FT', 'RP', 'CP', 'DI')
    AND i.PUBLISHING_GROUP NOT IN ('MKT', 'ZZZ')
"""


def build_sales_query(start_date: str) -> str:
    return f"""
    WITH raw_bookscan AS (
        SELECT
            CAST([Week] AS date) AS [Week],
            LTRIM(RTRIM(CAST([ISBN10] AS varchar(20)))) AS RawISBN,
            SUM(CAST([TotalSales] AS bigint)) AS qty
        FROM [CBQ2].[old].[Sellthrough_Bookscan_bkup]
        WHERE
            [Week] <= '2018-03-31'
            AND [Week] >= '2007-01-01'
        GROUP BY
            CAST([Week] AS date),
            LTRIM(RTRIM(CAST([ISBN10] AS varchar(20))))
        UNION ALL
        SELECT
            CAST([WEEK] AS date) AS [Week],
            LTRIM(RTRIM(CAST([ISBN] AS varchar(20)))) AS RawISBN,
            SUM(CAST([Sales] AS bigint)) AS qty
        FROM [CBQ2].[cb].[Sellthrough_RollBookscan]
        WHERE
            [WEEK] > '2018-03-31'
            AND [WEEK] >= '2007-01-01'
        GROUP BY
            CAST([WEEK] AS date),
            LTRIM(RTRIM(CAST([ISBN] AS varchar(20))))
    )
    SELECT
        [Week],
        RawISBN,
        qty,
        'weekly' AS HistoryType
    FROM raw_bookscan
    WHERE [Week] >= '{start_date}'
    UNION ALL
    SELECT
        DATEFROMPARTS(YEAR([Week]), 12, 31) AS [Week],
        RawISBN,
        SUM(qty) AS qty,
        'yearly' AS HistoryType
    FROM raw_bookscan
    WHERE [Week] < '2023-01-01'
    GROUP BY
        YEAR([Week]),
        RawISBN;
    """


LATEST_WEEK_QUERY = """
WITH all_weeks AS (
    SELECT CAST([Week] AS date) AS [Week]
    FROM [CBQ2].[old].[Sellthrough_Bookscan_bkup]
    WHERE [Week] <= '2018-03-31'
    UNION ALL
    SELECT CAST([WEEK] AS date) AS [Week]
    FROM [CBQ2].[cb].[Sellthrough_RollBookscan]
    WHERE [WEEK] > '2018-03-31'
)
SELECT MAX([Week]) AS latest_week
FROM all_weeks;
"""


DISTINCT_WEEKS_QUERY = """
WITH all_weeks AS (
    SELECT DISTINCT CAST([Week] AS date) AS [Week]
    FROM [CBQ2].[old].[Sellthrough_Bookscan_bkup]
    WHERE [Week] <= '2018-03-31'
    UNION
    SELECT DISTINCT CAST([WEEK] AS date) AS [Week]
    FROM [CBQ2].[cb].[Sellthrough_RollBookscan]
    WHERE [WEEK] > '2018-03-31'
)
SELECT [Week]
FROM all_weeks
ORDER BY [Week];
"""


SOURCE_METADATA_QUERY = """
WITH raw_isbns AS (
    SELECT DISTINCT LTRIM(RTRIM(CAST([ISBN10] AS varchar(20)))) AS RawISBN
    FROM [CBQ2].[old].[Sellthrough_Bookscan_bkup]
    WHERE
        [Week] <= '2018-03-31'
        AND [Week] >= '2007-01-01'
    UNION
    SELECT DISTINCT LTRIM(RTRIM(CAST([ISBN] AS varchar(20)))) AS RawISBN
    FROM [CBQ2].[cb].[Sellthrough_RollBookscan]
    WHERE
        [WEEK] > '2018-03-31'
        AND [WEEK] >= '2007-01-01'
)
SELECT DISTINCT
    COALESCE(i.ITEM_TITLE, r.RawISBN) AS ISBN,
    CASE
        WHEN i.PUBLISHER_CODE = 'Quadrille Publishing Limited' THEN 'Quadrille'
        ELSE i.PUBLISHER_CODE
    END AS Pub,
    i.PRODUCT_TYPE AS PT,
    i.FORMAT AS CAT,
    CASE
        WHEN LEFT(i.PUBLISHING_GROUP, 3) = 'BAR' THEN 'BAR'
        ELSE i.PUBLISHING_GROUP
    END AS pgrp,
    i.SHORT_TITLE AS Title,
    i.PRICE_AMOUNT AS Price,
    CAST(i.AMORTIZATION_DATE AS date) AS PubDate
FROM raw_isbns r
LEFT JOIN ebs.item i
    ON i.ITEM_TITLE = r.RawISBN
WHERE LEN(r.RawISBN) <> 10
{PUBLISHER_EXCLUSION_SQL}
UNION
SELECT DISTINCT
    COALESCE(i.ITEM_TITLE, r.RawISBN) AS ISBN,
    CASE
        WHEN i.PUBLISHER_CODE = 'Quadrille Publishing Limited' THEN 'Quadrille'
        ELSE i.PUBLISHER_CODE
    END AS Pub,
    i.PRODUCT_TYPE AS PT,
    i.FORMAT AS CAT,
    CASE
        WHEN LEFT(i.PUBLISHING_GROUP, 3) = 'BAR' THEN 'BAR'
        ELSE i.PUBLISHING_GROUP
    END AS pgrp,
    i.SHORT_TITLE AS Title,
    i.PRICE_AMOUNT AS Price,
    CAST(i.AMORTIZATION_DATE AS date) AS PubDate
FROM raw_isbns r
LEFT JOIN ebs.item i
    ON i.ISBN = r.RawISBN
WHERE LEN(r.RawISBN) = 10
{PUBLISHER_EXCLUSION_SQL};
"""
