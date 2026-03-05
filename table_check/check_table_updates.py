import sys

from sqlalchemy import create_engine
import pandas as pd


def get_connection() -> create_engine:
    engine = create_engine("mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server")
    return engine


def fetch_data_from_db(engine: create_engine, query: str) -> pd.DataFrame:
    raw_connection = engine.raw_connection()
    try:
        cursor = raw_connection.cursor()
        try:
            cursor.execute(query)
            columns = None
            rows = None
            while True:
                if cursor.description:
                    columns = [col[0] for col in cursor.description]
                    rows = cursor.fetchall()
                    break
                if not cursor.nextset():
                    raise RuntimeError("Query did not return a result set.")
        finally:
            cursor.close()
    finally:
        raw_connection.close()

    if columns is None or rows is None:
        return pd.DataFrame()

    return pd.DataFrame.from_records(rows, columns=columns)


SQL_QUERIES = {
    "1": (
        "All Updates",
        """
SELECT TOP 100
    tlu.TableName,
    tlu.LastUpdated
FROM metrics.TableLastUpdated tlu;
""",
    ),
    "2": (
        "Tables for SSR Summary",
        """
SELECT
    tlu.TableName,
    tlu.LastUpdated
FROM metrics.TableLastUpdated tlu
WHERE tlu.TableName IN ('ssr.SalesSSRRow', 'ebs.Sales', 'ebs.Item')
ORDER BY tlu.LastUpdated DESC;
""",
    ),
    "3": (
        "Amazon",
        """
SELECT DISTINCT TOP 10
    [WEEK]
FROM [CBQ2].[cb].[Sellthrough_Amazon] sta
ORDER BY [WEEK] DESC;
""",
    ),
    "4": (
        "Bookscan",
        """
SELECT DISTINCT TOP 10
    [WEEK]
FROM [CBQ2].[cb].[Sellthrough_RollBookscan]
ORDER BY [WEEK] DESC;
""",
    ),
    "5": (
        "Barnes & Noble",
        """
SELECT DISTINCT TOP 10
    [WEEK]
FROM [CBQ2].[cb].[Sellthrough_Barnes_and_Noble]
ORDER BY [WEEK] DESC;
""",
    ),
}


def main():
    if len(sys.argv) < 2:
        print("Please provide a query choice: 1, 2, 3, 4, or 5.")
        return 1

    choice = sys.argv[1].strip()
    if choice not in SQL_QUERIES:
        print(f"Invalid query choice: {choice}. Use 1, 2, 3, 4, or 5.")
        return 1

    report_name, sql_query = SQL_QUERIES[choice]

    engine = get_connection()
    df = fetch_data_from_db(engine, sql_query)

    print(f"\n{report_name} Results:")
    if df.empty:
        print("No rows returned.")
        return 0

    # Improve readability for table-update timestamps (queries 1 and 2).
    if choice in {"1", "2"} and "LastUpdated" in df.columns:
        df["LastUpdated"] = (
            pd.to_datetime(df["LastUpdated"], errors="coerce")
            .dt.strftime("%Y-%m-%d %I:%M:%S %p")
            .fillna("")
        )

    print(df.to_string(index=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
