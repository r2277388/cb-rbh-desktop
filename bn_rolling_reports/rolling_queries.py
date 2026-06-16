from __future__ import annotations


def build_customer_sales_query(start_date: str) -> str:
    return f"""
    SELECT
        CAST(sbn.[WEEK] AS date) AS [Week],
        sbn.[ISBN13] AS ISBN,
        i.SHORT_TITLE AS Title,
        CAST(COALESCE(i.AMORTIZATION_DATE, osd.OSD) AS date) AS PubDate,
        i.PRICE_AMOUNT AS Price,
        CASE
            WHEN i.PUBLISHER_CODE = 'Quadrille Publishing Limited' THEN 'Quadrille'
            ELSE i.PUBLISHER_CODE
        END AS Publisher,
        i.PRODUCT_TYPE AS PT,
        i.FORMAT AS CAT,
        CASE
            WHEN LEFT(i.PUBLISHING_GROUP, 3) = 'BAR' THEN 'BAR'
            ELSE i.PUBLISHING_GROUP
        END AS pgrp,
        sbn.[SubjectCode],
        sbn.[DeptCode],
        SUM(ISNULL(sbn.[BarnesAndNobleTotal], 0)) AS qty
    FROM [CBQ2].[cb].[Sellthrough_Barnes_and_Noble] sbn
    INNER JOIN ebs.item i
        ON i.ITEM_TITLE = sbn.ISBN13
    LEFT JOIN (
        SELECT
            tt.ean13 AS ISBN,
            MAX(tt.active_datevalue) AS OSD
        FROM tmm.cb_Import_Title_Tasks tt
        WHERE tt.date_desc = 'On Sale Date'
          AND tt.active_datevalue IS NOT NULL
          AND tt.printingnumber = 1
          AND tt.activeind = '1'
        GROUP BY tt.ean13
    ) osd
        ON osd.ISBN = i.ITEM_TITLE
    WHERE
        sbn.[WEEK] >= '{start_date}'
        AND i.SHORT_TITLE IS NOT NULL
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
    GROUP BY
        CAST(sbn.[WEEK] AS date),
        sbn.[ISBN13],
        i.SHORT_TITLE,
        CAST(COALESCE(i.AMORTIZATION_DATE, osd.OSD) AS date),
        i.PRICE_AMOUNT,
        CASE
            WHEN i.PUBLISHER_CODE = 'Quadrille Publishing Limited' THEN 'Quadrille'
            ELSE i.PUBLISHER_CODE
        END,
        i.PRODUCT_TYPE,
        i.FORMAT,
        CASE
            WHEN LEFT(i.PUBLISHING_GROUP, 3) = 'BAR' THEN 'BAR'
            ELSE i.PUBLISHING_GROUP
        END,
        sbn.[SubjectCode],
        sbn.[DeptCode];
    """
