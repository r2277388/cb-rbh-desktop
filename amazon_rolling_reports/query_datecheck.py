def check_date():
    """
    Returns a SQL string to check the most recent week loaded
    in [CBQ2].[cb].[Sellthrough_Amazon].
    """
    return """
    SELECT MAX(sta.[Week]) AS LatestWeek
    FROM [CBQ2].[cb].[Sellthrough_Amazon] sta;
    """