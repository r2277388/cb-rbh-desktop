from __future__ import annotations


ROLLOVER_EXCEPTION_WEEKS = (
    "2017-12-09",
    "2017-11-18",
    "2017-08-26",
    "2017-07-08",
    "2017-05-27",
)
ROLLOVER_EXCEPTION_WEEKS_SQL = ", ".join(f"'{week}'" for week in ROLLOVER_EXCEPTION_WEEKS)


def _old_bookscan_week_filter(week_column: str, start_date: str | None = None) -> str:
    conditions = [
        f"{week_column} <= '2018-03-31'",
        f"CAST({week_column} AS date) NOT IN ({ROLLOVER_EXCEPTION_WEEKS_SQL})",
    ]
    if start_date is not None:
        conditions.append(f"{week_column} >= '{start_date}'")
    return "\n            AND ".join(conditions)


def _new_bookscan_week_filter(week_column: str, start_date: str | None = None) -> str:
    conditions = [
        f"({week_column} > '2018-03-31' OR CAST({week_column} AS date) IN ({ROLLOVER_EXCEPTION_WEEKS_SQL}))",
    ]
    if start_date is not None:
        conditions.append(f"{week_column} >= '{start_date}'")
    return "\n            AND ".join(conditions)


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
            {_old_bookscan_week_filter("[Week]", "2007-01-01")}
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
            {_new_bookscan_week_filter("[WEEK]", "2007-01-01")}
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


def build_distinct_weeks_since_query(start_date: str) -> str:
    return f"""
    WITH all_weeks AS (
        SELECT DISTINCT CAST([Week] AS date) AS [Week]
        FROM [CBQ2].[old].[Sellthrough_Bookscan_bkup]
        WHERE
            {_old_bookscan_week_filter("[Week]", start_date)}
        UNION
        SELECT DISTINCT CAST([WEEK] AS date) AS [Week]
        FROM [CBQ2].[cb].[Sellthrough_RollBookscan]
        WHERE
            {_new_bookscan_week_filter("[WEEK]", start_date)}
    )
    SELECT [Week]
    FROM all_weeks
    ORDER BY [Week];
    """


LATEST_WEEK_QUERY = f"""
WITH all_weeks AS (
    SELECT CAST([Week] AS date) AS [Week]
    FROM [CBQ2].[old].[Sellthrough_Bookscan_bkup]
    WHERE {_old_bookscan_week_filter("[Week]")}
    UNION ALL
    SELECT CAST([WEEK] AS date) AS [Week]
    FROM [CBQ2].[cb].[Sellthrough_RollBookscan]
    WHERE {_new_bookscan_week_filter("[WEEK]")}
)
SELECT MAX([Week]) AS latest_week
FROM all_weeks;
"""


DISTINCT_WEEKS_QUERY = f"""
WITH all_weeks AS (
    SELECT DISTINCT CAST([Week] AS date) AS [Week]
    FROM [CBQ2].[old].[Sellthrough_Bookscan_bkup]
    WHERE {_old_bookscan_week_filter("[Week]")}
    UNION
    SELECT DISTINCT CAST([WEEK] AS date) AS [Week]
    FROM [CBQ2].[cb].[Sellthrough_RollBookscan]
    WHERE {_new_bookscan_week_filter("[WEEK]")}
)
SELECT [Week]
FROM all_weeks
ORDER BY [Week];
"""


MISSING_WEEKS_QUERY = f"""
WITH all_weeks AS (
    SELECT DISTINCT CAST([Week] AS date) AS week_end
    FROM [CBQ2].[old].[Sellthrough_Bookscan_bkup]
    WHERE {_old_bookscan_week_filter("[Week]")}

    UNION

    SELECT DISTINCT CAST([WEEK] AS date) AS week_end
    FROM [CBQ2].[cb].[Sellthrough_RollBookscan]
    WHERE {_new_bookscan_week_filter("[WEEK]")}
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


SOURCE_METADATA_QUERY = f"""
WITH raw_isbns AS (
    SELECT DISTINCT LTRIM(RTRIM(CAST([ISBN10] AS varchar(20)))) AS RawISBN
    FROM [CBQ2].[old].[Sellthrough_Bookscan_bkup]
    WHERE
        {_old_bookscan_week_filter("[Week]", "2007-01-01")}
    UNION
    SELECT DISTINCT LTRIM(RTRIM(CAST([ISBN] AS varchar(20)))) AS RawISBN
    FROM [CBQ2].[cb].[Sellthrough_RollBookscan]
    WHERE
        {_new_bookscan_week_filter("[WEEK]", "2007-01-01")}
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
