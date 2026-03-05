from functions import fetch_data_from_db, get_connection

SQL_LAST_10_WEEKS = """
SELECT TOP 10
    sta.[WEEK],
    COUNT(*) AS [RowCount],
    SUM(sta.[CustomerOrders]) AS [CustomerOrders],
    SUM(sta.[UnitShipped]) AS [UnitShipped],
    SUM(sta.[OnHand]) AS [OnHand]
FROM [CBQ2].[cb].[Sellthrough_Amazon] sta
GROUP BY sta.[WEEK]
ORDER BY sta.[WEEK] DESC;
"""


def main():
    engine = get_connection()
    df = fetch_data_from_db(engine, SQL_LAST_10_WEEKS)

    if df.empty:
        print("No rows returned for the last-10-weeks check.")
        return

    print("\nLast 10 weeks from [CBQ2].[cb].[Sellthrough_Amazon]:")
    print(df.to_string(index=False))


if __name__ == "__main__":
    main()
