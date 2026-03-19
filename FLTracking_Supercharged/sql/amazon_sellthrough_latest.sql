WITH Isbn12Exceptions AS (
    SELECT DISTINCT
        LTRIM(RTRIM(i.ITEM_TITLE)) AS ISBN12
    FROM ebs.item i
    WHERE
        i.ISBN IS NOT NULL
        AND LTRIM(RTRIM(i.ITEM_TITLE)) NOT LIKE '%[^0-9]%'
        AND LEN(LTRIM(RTRIM(i.ITEM_TITLE))) = 12
),

LastWeek AS (
    SELECT
        MAX(sta.[WEEK]) AS LastWeek
    FROM [CBQ2].[cb].[Sellthrough_Amazon] sta
),

AmazonNormalized AS (
    SELECT
        sta.[WEEK],
        sta.[UnitShipped],
        sta.[OnHand],
        sta.[ISBN] AS OriginalISBN,
        CleanISBN =
            REPLACE(REPLACE(LTRIM(RTRIM(sta.[ISBN])), '-', ''), ' ', '')
    FROM [CBQ2].[cb].[Sellthrough_Amazon] sta
),

AmazonStandardized AS (
    SELECT
        an.[WEEK],
        an.[UnitShipped],
        an.[OnHand],
        an.OriginalISBN,
        CASE
            WHEN an.CleanISBN NOT LIKE '%[^0-9]%' AND LEN(an.CleanISBN) = 13
                THEN an.CleanISBN

            WHEN an.CleanISBN NOT LIKE '%[^0-9]%'
                 AND LEN(an.CleanISBN) = 12
                 AND EXISTS (
                     SELECT 1
                     FROM Isbn12Exceptions x
                     WHERE x.ISBN12 = an.CleanISBN
                 )
                THEN an.CleanISBN

            WHEN an.CleanISBN NOT LIKE '%[^0-9]%'
                 AND LEN(an.CleanISBN) < 13
                THEN RIGHT(REPLICATE('0', 13) + an.CleanISBN, 13)

            ELSE NULL
        END AS StandardizedISBN
    FROM AmazonNormalized an
),

AmazonLTD AS (
    SELECT
        s.StandardizedISBN,
        SUM(s.[UnitShipped]) AS AmzUnitShipped_LTD
    FROM AmazonStandardized s
    WHERE
        s.StandardizedISBN IS NOT NULL
        AND s.[WEEK] >= DATEFROMPARTS(YEAR(GETDATE()) - 1, 1, 1)
    GROUP BY
        s.StandardizedISBN
),

AmazonLastWeek AS (
    SELECT
        s.StandardizedISBN,
        lw.LastWeek AS AmzWeek,
        SUM(s.[OnHand]) AS AmzOnHand
    FROM AmazonStandardized s
    INNER JOIN LastWeek lw
        ON s.[WEEK] = lw.LastWeek
    WHERE
        s.StandardizedISBN IS NOT NULL
    GROUP BY
        s.StandardizedISBN,
        lw.LastWeek
)

SELECT
    i.ITEM_TITLE AS ISBN,
    ISNULL(altd.AmzUnitShipped_LTD, 0) AS AmzUnitShipped_LTD,
    lw.LastWeek AS AmzLastWeek,
    ISNULL(alw.AmzOnHand, 0) AS AmzOnHand
FROM ebs.item i
LEFT JOIN AmazonLTD altd
    ON altd.StandardizedISBN = LTRIM(RTRIM(i.ITEM_TITLE))
LEFT JOIN AmazonLastWeek alw
    ON alw.StandardizedISBN = LTRIM(RTRIM(i.ITEM_TITLE))
CROSS JOIN LastWeek lw
WHERE
    LTRIM(RTRIM(i.ITEM_TITLE)) NOT LIKE '%[^0-9]%'
    AND LEN(LTRIM(RTRIM(i.ITEM_TITLE))) IN (12, 13)
    AND (
        ISNULL(altd.AmzUnitShipped_LTD, 0) <> 0
        OR ISNULL(alw.AmzOnHand, 0) <> 0
    );