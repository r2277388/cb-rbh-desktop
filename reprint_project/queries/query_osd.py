def query_osd():
    return '''
    SELECT
        tt.ean13 AS ISBN,
        tt.active_datevalue AS OSD
    FROM
        tmm.cb_Import_Title_Tasks tt
        INNER JOIN ebs.item i ON i.ITEM_TITLE = tt.ean13
    WHERE
        tt.date_desc = 'On Sale Date'
        AND tt.active_datevalue IS NOT NULL
        AND tt.printingnumber = 1
        AND i.PUBLISHER_CODE = 'Chronicle'
        AND tt.active_datevalue >= DATEADD(year, -1, GETDATE())
    ORDER BY
        OSD DESC;
    '''