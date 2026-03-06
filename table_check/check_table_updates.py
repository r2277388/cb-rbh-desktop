import sys
from pathlib import Path

import pandas as pd

# Ensure repo root is importable when this script is executed by file path.
ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from shared.db import fetch_data_from_db, get_connection
from shared.sql import load_sql


SQL_QUERIES = {
    "1": (
        "All Updates",
        load_sql("table_check", "all_updates.sql"),
    ),
    "2": (
        "Tables for SSR Summary",
        load_sql("table_check", "ssr_summary_tables.sql"),
    ),
    "3": (
        "Amazon",
        load_sql("table_check", "amazon_weeks.sql"),
    ),
    "4": (
        "Bookscan",
        load_sql("table_check", "bookscan_weeks.sql"),
    ),
    "5": (
        "Barnes & Noble",
        load_sql("table_check", "bn_weeks.sql"),
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
