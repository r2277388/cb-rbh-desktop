from pathlib import Path


BASE_SQL_DIR = Path(__file__).resolve().parent


def load_sql(*parts: str) -> str:
    """Load a SQL file relative to shared/sql."""
    sql_path = BASE_SQL_DIR.joinpath(*parts)
    return sql_path.read_text(encoding="utf-8").strip()

