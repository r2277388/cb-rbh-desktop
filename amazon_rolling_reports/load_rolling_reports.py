import time
from datetime import datetime
from pathlib import Path

import pandas as pd

from functions import get_connection, fetch_data_from_db, save_to_pickle
from query_co import sql_co
from query_us import sql_us


BACKUP_DIR = Path("amazon_rolling_reports") / "backups"
RECENT_HISTORY_WEEKS_TO_CHECK = 12
MIN_NONZERO_PRIOR_WEEKS = 3


def history_columns(df: pd.DataFrame) -> list[str]:
    return [
        col for col in df.columns
        if isinstance(col, str) and len(col) == 10 and col.count("-") == 2
    ]


def recent_nonzero_history_weeks(df: pd.DataFrame, weeks_to_check: int = RECENT_HISTORY_WEEKS_TO_CHECK) -> tuple[int, list[tuple[str, float]]]:
    history_cols = history_columns(df)[:weeks_to_check]
    nonzero_weeks: list[tuple[str, float]] = []
    for col in history_cols:
        week_sum = float(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())
        if week_sum != 0:
            nonzero_weeks.append((col, week_sum))
    return len(nonzero_weeks), nonzero_weeks


def backup_existing_pickle(filename: str) -> Path | None:
    source = Path(filename)
    if not source.exists():
        return None

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"{source.stem}_{timestamp}{source.suffix}"
    backup_path.write_bytes(source.read_bytes())
    return backup_path


def validate_history_against_existing(df_new: pd.DataFrame, filename: str, label: str) -> None:
    existing_path = Path(filename)
    if not existing_path.exists():
        return

    try:
        df_existing = pd.read_pickle(existing_path)
    except Exception as exc:
        print(f"Warning: could not inspect existing pickle {filename}: {exc}")
        return

    existing_nonzero_count, existing_nonzero_weeks = recent_nonzero_history_weeks(df_existing)
    new_nonzero_count, new_nonzero_weeks = recent_nonzero_history_weeks(df_new)

    if existing_nonzero_count >= MIN_NONZERO_PRIOR_WEEKS and new_nonzero_count <= 1:
        latest_week = history_columns(df_new)[:1]
        latest_week_text = latest_week[0] if latest_week else "unknown"
        existing_preview = ", ".join(week for week, _ in existing_nonzero_weeks[:6])
        new_preview = ", ".join(week for week, _ in new_nonzero_weeks[:6]) or "(none beyond latest week)"
        raise ValueError(
            f'{label} SQL refresh looks incomplete and would overwrite good history. '
            f'Existing pickle has {existing_nonzero_count} non-zero recent week columns '
            f'({existing_preview}), but the new query result only has {new_nonzero_count} '
            f'non-zero recent week column(s) near {latest_week_text} ({new_preview}).'
        )

def run_query_and_save(query_func, filename, label):
    print(f'Loading "{label}" data from the SQL database...')
    print('Querying this data can take up to 4 minutes each ...')
    engine = get_connection()
    query = query_func()
    df = fetch_data_from_db(engine, query)
    print('Query completed ...')
    print(f'{label} shape: {df.shape}')
    validate_history_against_existing(df, filename, label)
    backup_path = backup_existing_pickle(filename)
    if backup_path is not None:
        print(f"Backed up existing pickle to {backup_path}")
    save_to_pickle(df, filename)

def main():
    print("Starting the data extraction and saving process...")
    time.sleep(2)
    print("First, we are running the customer orders query ...")
    run_query_and_save(sql_co, "rr_customer_orders.pkl", "Customer Orders")
    time.sleep(2)
    print("Now, we are running the units shipped query ...")
    run_query_and_save(sql_us, "rr_units_shipped.pkl", "Units Shipped")

if __name__ == "__main__":
    main()
