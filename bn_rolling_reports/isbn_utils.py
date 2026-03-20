import re
import sys
from functools import lru_cache
from pathlib import Path

import pandas as pd


sys.path.append(str(Path(__file__).resolve().parents[1]))

from shared.db.connection import get_connection
from shared.db.query_runner import fetch_data_from_db


ISBN12_EXCEPTIONS_SQL = """
SELECT DISTINCT
    LTRIM(RTRIM(i.ITEM_TITLE)) AS ISBN12
FROM ebs.item i
WHERE
    i.ISBN IS NOT NULL
    AND LTRIM(RTRIM(i.ITEM_TITLE)) NOT LIKE '%[^0-9]%'
    AND LEN(LTRIM(RTRIM(i.ITEM_TITLE))) = 12
"""

ISBN12_EXCEPTIONS_FALLBACK = Path(__file__).with_name("isbn12_exceptions.txt")
_ISBN12_NOTICE_PRINTED = False


@lru_cache(maxsize=1)
def load_isbn12_exceptions() -> set[str] | None:
    global _ISBN12_NOTICE_PRINTED
    try:
        engine = get_connection()
        df = fetch_data_from_db(engine, ISBN12_EXCEPTIONS_SQL)
        if "ISBN12" not in df.columns:
            raise ValueError("ISBN12 exception query did not return the ISBN12 column.")
        return {
            str(value).strip()
            for value in df["ISBN12"].dropna().tolist()
            if str(value).strip()
        }
    except Exception as exc:
        if ISBN12_EXCEPTIONS_FALLBACK.exists():
            return {
                line.strip()
                for line in ISBN12_EXCEPTIONS_FALLBACK.read_text().splitlines()
                if line.strip()
            }
        if not _ISBN12_NOTICE_PRINTED:
            print(
                "Note: SQL is unavailable, so 12-digit ISBN exceptions could not be checked. "
                "Any 12-digit numeric EANs will be kept as-is."
            )
            _ISBN12_NOTICE_PRINTED = True
        return None


def normalize_isbn(value: object, isbn12_exceptions: set[str] | None = None) -> str | None:
    if pd.isna(value):
        return None

    clean = re.sub(r"[-\s]", "", str(value).strip())
    if not clean or not clean.isdigit():
        return None

    if len(clean) == 13:
        return clean

    exceptions = isbn12_exceptions if isbn12_exceptions is not None else load_isbn12_exceptions()

    if len(clean) == 12:
        if exceptions is None or clean in exceptions:
            return clean

    if len(clean) < 13:
        return clean.zfill(13)

    return None


def normalize_isbn_series(series: pd.Series) -> pd.Series:
    exceptions = load_isbn12_exceptions()
    normalized = series.map(lambda value: normalize_isbn(value, exceptions))
    return normalized.astype("object")
