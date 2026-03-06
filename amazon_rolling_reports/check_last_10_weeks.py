import sys
from pathlib import Path

# Ensure repo root is importable when this script is executed by file path.
ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from shared.db import fetch_data_from_db, get_connection
from shared.sql import load_sql

SQL_LAST_10_WEEKS = load_sql("amazon_rolling_reports", "last_10_weeks.sql")


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
