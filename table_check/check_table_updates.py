import sys
from datetime import datetime
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
        "Ebs.Sales Prior 5 Days",
        None,
    ),
    "4": (
        "Amazon",
        load_sql("table_check", "amazon_weeks.sql"),
    ),
    "5": (
        "Bookscan",
        load_sql("table_check", "bookscan_weeks.sql"),
    ),
    "6": (
        "Barnes & Noble",
        load_sql("table_check", "bn_weeks.sql"),
    ),
    "7": (
        "Freight Costs",
        None,
    ),
}


def _prior_period_yyyymm(today: datetime | None = None) -> str:
    now = today or datetime.now()
    year = now.year
    month = now.month - 1
    if month == 0:
        year -= 1
        month = 12
    return f"{year}{month:02d}"


def _current_period_yyyymm(today: datetime | None = None) -> str:
    now = today or datetime.now()
    return f"{now.year}{now.month:02d}"


def build_ebs_sales_prior_5_days_sql() -> str:
    current_period = _current_period_yyyymm()
    prior_period = _prior_period_yyyymm()
    return f"""
SELECT TOP (5)
    sd.TRX_DATE AS [Date],
    SUM(CASE WHEN i.PUBLISHER_CODE = 'Chronicle' THEN sd.REVENUE_AMOUNT ELSE 0 END) AS cb_val,
    SUM(CASE WHEN i.PUBLISHER_CODE <> 'Chronicle' THEN sd.REVENUE_AMOUNT ELSE 0 END) AS dp_val,
    SUM(CASE WHEN i.PUBLISHER_CODE = 'Chronicle' THEN 1 ELSE 0 END) AS cb_row_cnt,
    SUM(CASE WHEN i.PUBLISHER_CODE <> 'Chronicle' THEN 1 ELSE 0 END) AS dp_row_cnt
FROM ebs.sales sd
INNER JOIN ebs.item i
    ON i.ITEM_ID = sd.ITEM_ID
WHERE
    sd.PERIOD BETWEEN '{prior_period}' AND '{current_period}'
GROUP BY
    sd.TRX_DATE
ORDER BY
    sd.TRX_DATE DESC;
""".strip()


def build_freight_costs_sql() -> str:
    return """
SELECT TOP (5)
    sr.FISCALPERIOD AS [period],
    SUM(sr.cost) AS [Cost]
FROM cb.shippingreport sr
WHERE TRY_CONVERT(int, sr.FISCALPERIOD) IS NOT NULL
GROUP BY sr.FISCALPERIOD
ORDER BY TRY_CONVERT(int, sr.FISCALPERIOD) DESC;
""".strip()


def _format_for_display(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    force_numeric_cols = {"cb_val", "dp_val", "cb_row_cnt", "dp_row_cnt", "Cost"}
    for col in force_numeric_cols.intersection(out.columns):
        out[col] = pd.to_numeric(out[col], errors="coerce")

    numeric_cols = out.select_dtypes(include=["number"]).columns
    for col in numeric_cols:
        out[col] = pd.to_numeric(out[col], errors="coerce").map(
            lambda x: "" if pd.isna(x) else f"{x:,.0f}"
        )
    return out


def _print_grid_table(df: pd.DataFrame) -> None:
    cols = [str(c) for c in df.columns]
    rows = [[str(v) for v in row] for row in df.fillna("").itertuples(index=False, name=None)]

    widths = []
    for i, col in enumerate(cols):
        max_cell = max((len(r[i]) for r in rows), default=0)
        widths.append(max(len(col), max_cell))

    sep = "+" + "+".join("-" * (w + 2) for w in widths) + "+"
    header = "|" + "|".join(f" {col:<{widths[i]}} " for i, col in enumerate(cols)) + "|"

    print(sep)
    print(header)
    print(sep)
    for row in rows:
        print("|" + "|".join(f" {row[i]:<{widths[i]}} " for i in range(len(cols))) + "|")
    print(sep)


def main():
    if len(sys.argv) < 2:
        print("Please provide a query choice: 1, 2, 3, 4, 5, 6, or 7.")
        return 1

    choice = sys.argv[1].strip()
    if choice not in SQL_QUERIES:
        print(f"Invalid query choice: {choice}. Use 1, 2, 3, 4, 5, 6, or 7.")
        return 1

    report_name, sql_query = SQL_QUERIES[choice]
    if choice == "3":
        sql_query = build_ebs_sales_prior_5_days_sql()
    elif choice == "7":
        sql_query = build_freight_costs_sql()

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

    display_df = _format_for_display(df)
    _print_grid_table(display_df)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
