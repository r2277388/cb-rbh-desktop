import sys
from pathlib import Path

import pandas as pd

from config import OUTPUT_DIR, SQL_DIR
from isbn_utils import normalize_isbn_series


sys.path.append(str(Path(__file__).resolve().parents[2]))

from shared.db.connection import get_connection
from shared.db.query_runner import fetch_data_from_db


SQL_FILE = SQL_DIR / "faire_orders.sql"


def load_faire_orders(sql_path: Path = SQL_FILE) -> pd.DataFrame:
    query = sql_path.read_text(encoding="utf-8").strip()
    engine = get_connection()
    df = fetch_data_from_db(engine, query)

    required_columns = ["ISBN", "Faire_OO_qty"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(
            "Faire orders query is missing required columns: "
            + ", ".join(missing_columns)
        )

    result = df[required_columns].copy()
    result = result.dropna(subset=["ISBN"])
    result["ISBN"] = normalize_isbn_series(result["ISBN"])
    result = result.dropna(subset=["ISBN"])

    return result


def save_faire_orders_output(df: pd.DataFrame) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / "faire_orders.pkl"
    df.to_pickle(output_path)
    return output_path


def main() -> None:
    df = load_faire_orders()
    output_path = save_faire_orders_output(df)

    print(f"Loaded SQL from: {SQL_FILE}")
    print(f"Rows after removing null ISBN values: {len(df)}")
    print(df.head(20).to_string(index=False))
    print(f"\nSaved output to: {output_path}")


if __name__ == "__main__":
    main()
