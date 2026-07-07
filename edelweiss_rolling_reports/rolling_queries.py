from __future__ import annotations


def build_sales_query(start_date: str) -> str:
    return f"""
    SELECT
        CAST(ste.[WEEK] AS date) AS [Week],
        LTRIM(RTRIM(CAST(ste.ISBN AS varchar(20)))) AS RawISBN,
        SUM(CAST(ste.WeekEnding AS bigint)) AS qty
    FROM [CBQ2].[cb].[Sellthrough_Edelweiss] ste
    WHERE ste.[WEEK] >= '{start_date}'
    GROUP BY
        CAST(ste.[WEEK] AS date),
        LTRIM(RTRIM(CAST(ste.ISBN AS varchar(20))));
    """


def build_distinct_weeks_since_query(start_date: str) -> str:
    return f"""
    SELECT DISTINCT CAST(ste.[WEEK] AS date) AS [Week]
    FROM [CBQ2].[cb].[Sellthrough_Edelweiss] ste
    WHERE ste.[WEEK] >= '{start_date}'
    ORDER BY [Week];
    """


LATEST_WEEK_QUERY = """
SELECT MAX(CAST(ste.[WEEK] AS date)) AS latest_week
FROM [CBQ2].[cb].[Sellthrough_Edelweiss] ste;
"""


DISTINCT_WEEKS_QUERY = """
SELECT DISTINCT CAST(ste.[WEEK] AS date) AS [Week]
FROM [CBQ2].[cb].[Sellthrough_Edelweiss] ste
ORDER BY [Week];
"""


MISSING_WEEKS_QUERY = """
WITH all_weeks AS (
    SELECT DISTINCT CAST(ste.[WEEK] AS date) AS week_end
    FROM [CBQ2].[cb].[Sellthrough_Edelweiss] ste
),
ordered_weeks AS (
    SELECT
        week_end,
        LEAD(week_end) OVER (ORDER BY week_end) AS next_week_end
    FROM all_weeks
),
gaps AS (
    SELECT
        DATEADD(day, 7, week_end) AS missing_week,
        next_week_end
    FROM ordered_weeks
    WHERE next_week_end IS NOT NULL
      AND DATEADD(day, 7, week_end) < next_week_end

    UNION ALL

    SELECT
        DATEADD(day, 7, missing_week),
        next_week_end
    FROM gaps
    WHERE DATEADD(day, 7, missing_week) < next_week_end
)
SELECT missing_week
FROM gaps
ORDER BY missing_week DESC
OPTION (MAXRECURSION 1000);
"""


SOURCE_METADATA_QUERY = """
WITH latest_source AS (
    SELECT
        LTRIM(RTRIM(CAST(ste.ISBN AS varchar(20)))) AS RawISBN,
        ste.Title,
        ste.Imprint,
        ste.ListPrice,
        ste.PubDate,
        ROW_NUMBER() OVER (
            PARTITION BY LTRIM(RTRIM(CAST(ste.ISBN AS varchar(20))))
            ORDER BY CAST(ste.[WEEK] AS date) DESC, ste.Id DESC
        ) AS rn
    FROM [CBQ2].[cb].[Sellthrough_Edelweiss] ste
)
SELECT DISTINCT
    COALESCE(i.ITEM_TITLE, s.RawISBN) AS ISBN,
    CASE
        WHEN COALESCE(i.PUBLISHER_CODE, s.Imprint) = 'Quadrille Publishing Limited' THEN 'Quadrille'
        ELSE COALESCE(i.PUBLISHER_CODE, s.Imprint)
    END AS Pub,
    i.PRODUCT_TYPE AS PT,
    i.FORMAT AS CAT,
    CASE
        WHEN LEFT(i.PUBLISHING_GROUP, 3) = 'BAR' THEN 'BAR'
        ELSE i.PUBLISHING_GROUP
    END AS PGRP,
    COALESCE(i.SHORT_TITLE, s.Title) AS TITLE,
    COALESCE(i.PRICE_AMOUNT, TRY_CONVERT(decimal(18, 2), s.ListPrice)) AS PRICE,
    COALESCE(CAST(i.AMORTIZATION_DATE AS date), TRY_CONVERT(date, s.PubDate)) AS PubDate
FROM latest_source s
LEFT JOIN ebs.item i
    ON i.ITEM_TITLE = s.RawISBN
WHERE s.rn = 1;
"""


LATEST_INVENTORY_QUERY = """
WITH latest_week AS (
    SELECT MAX(CAST(ste.[WEEK] AS date)) AS [Week]
    FROM [CBQ2].[cb].[Sellthrough_Edelweiss] ste
)
SELECT
    LTRIM(RTRIM(CAST(ste.ISBN AS varchar(20)))) AS RawISBN,
    SUM(CAST(ste.OH AS bigint)) AS [On Hand],
    SUM(CAST(ste.OO AS bigint)) AS [On Order]
FROM [CBQ2].[cb].[Sellthrough_Edelweiss] ste
INNER JOIN latest_week lw
    ON CAST(ste.[WEEK] AS date) = lw.[Week]
GROUP BY
    LTRIM(RTRIM(CAST(ste.ISBN AS varchar(20))));
"""
