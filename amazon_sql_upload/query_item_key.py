def item_sql():
    return """
    SELECT
        i.ISBN AS ISBN
    FROM
        ebs.Item i
    WHERE
        i.PRODUCT_TYPE IN ('BK','FT','DI','CP','RP','MI')

    """