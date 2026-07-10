import sys
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from bn_rolling_reports.isbn_utils import normalize_isbn
from paths import process_paths
from shared.db.connection import get_connection
from shared.db.query_runner import fetch_data_from_db

CONINV_PICKLE_FILE = process_paths.CONSOLIDATED_INVENTORY_VERTICALIZATION_FOLDER / "ConsolidatedInventory.pkl"
INVOBS_ALLOWED_ISBNS_SQL = """
SELECT
    i.ITEM_TITLE AS ISBN
FROM ebs.item i
WHERE
    i.ISBN IS NOT NULL
    AND i.PUBLISHING_GROUP NOT IN ('MKT', 'ZZZ')
    AND i.PUBLISHER_CODE = 'Chronicle'
    AND i.PRODUCT_TYPE IN ('BK', 'FT', 'RP', 'CP')
"""
TITLE_LOOKUP_BATCH_SIZE = 1000


def load_isbn_titles(isbns) -> dict[str, str]:
    clean_isbns = sorted({str(isbn).strip() for isbn in isbns if str(isbn).strip()})
    if not clean_isbns:
        return {}

    engine = get_connection()
    title_lookup = {}
    for start in range(0, len(clean_isbns), TITLE_LOOKUP_BATCH_SIZE):
        batch = clean_isbns[start:start + TITLE_LOOKUP_BATCH_SIZE]
        isbn_literals = ", ".join(
            f"'{isbn.replace(chr(39), chr(39) * 2)}'" for isbn in batch
        )
        query = f"""
SELECT
    i.ISBN AS ISBN,
    i.SHORT_TITLE AS Title
FROM ebs.item i
WHERE i.ISBN IN ({isbn_literals})
  AND i.SHORT_TITLE IS NOT NULL
"""
        titles = fetch_data_from_db(engine, query)
        if titles.empty:
            continue
        for _, row in titles.iterrows():
            isbn = str(row["ISBN"]).strip()
            title = str(row["Title"]).strip()
            if isbn and title:
                title_lookup.setdefault(isbn, title)

    return title_lookup


def load_allowed_invobs_isbns() -> set[str]:
    engine = get_connection()
    df = fetch_data_from_db(engine, INVOBS_ALLOWED_ISBNS_SQL)
    if "ISBN" not in df.columns:
        raise ValueError("INVOBS ISBN query must return an ISBN column.")

    return {
        isbn
        for isbn in (normalize_isbn(value) for value in df["ISBN"].dropna().tolist())
        if isbn
    }


def load_period_consolidated_inventory(period: str) -> pd.DataFrame:
    if not CONINV_PICKLE_FILE.exists():
        raise FileNotFoundError(f"ConInv pickle not found: {CONINV_PICKLE_FILE}")

    df = pd.read_pickle(CONINV_PICKLE_FILE).copy()
    df["period"] = df["period"].astype(str)
    df["ISBN"] = df["ISBN"].astype(str).str.strip()
    df["ORG"] = df["ORG"].astype(str).str.lower().str.strip()
    df["Publisher"] = df["Publisher"].astype(str).str.strip()
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0)
    df["Inventory"] = pd.to_numeric(df["Inventory"], errors="coerce").fillna(0)

    df = df[df["period"] == str(period)].copy()
    if df.empty:
        raise ValueError(f"No ConInv rows were found for period {period}.")

    return df


def filter_invobs_inventory_rows(df: pd.DataFrame) -> pd.DataFrame:
    period_label = (
        df["period"].iloc[0]
        if "period" in df.columns and not df.empty
        else "selected period"
    )
    allowed_isbns = load_allowed_invobs_isbns()
    filtered = df[df["ISBN"].isin(allowed_isbns)].copy()
    if filtered.empty:
        raise ValueError(f"No Chronicle BK/FT/RP/CP ISBN rows remained for {period_label}.")

    return filtered


def consolidate_inventory_from_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        raise ValueError("No ConInv rows were provided for consolidation.")

    pivot = (
        df.pivot_table(
            index="ISBN",
            columns="ORG",
            values=["Value", "Inventory"],
            aggfunc="sum",
            fill_value=0,
        )
        .sort_index(axis=1)
        .reset_index()
    )

    pivot.columns = [
        "ISBN" if col == ("ISBN", "") else f"{col[0].lower()}_{col[1].lower()}"
        for col in pivot.columns.to_flat_index()
    ]

    result = pd.DataFrame()
    result["ISBN"] = pivot["ISBN"]
    result["val_cbc"] = pivot["value_cbc"] if "value_cbc" in pivot.columns else 0
    result["val_hbg"] = pivot["value_hbg"] if "value_hbg" in pivot.columns else 0
    result["val_cbp"] = pivot["value_cbp"] if "value_cbp" in pivot.columns else 0
    result["units_cbc"] = pivot["inventory_cbc"] if "inventory_cbc" in pivot.columns else 0
    result["units_hbg"] = pivot["inventory_hbg"] if "inventory_hbg" in pivot.columns else 0
    result["units_cbp"] = pivot["inventory_cbp"] if "inventory_cbp" in pivot.columns else 0
    return result


def consolidate_inventory_from_pickle(period: str) -> pd.DataFrame:
    df = load_period_consolidated_inventory(period)
    df = filter_invobs_inventory_rows(df)
    return consolidate_inventory_from_rows(df)

def consolidate_inventory(period):
    print(">>> consolidate_inventory() started")
    print(f">>> Loading ConInv from pickle for period: {period}")
    df_from_pickle = consolidate_inventory_from_pickle(str(period))
    print(">>> ConInv pickle data loaded!")
    print("Loaded DataFrame shape:", df_from_pickle.shape)
    return df_from_pickle

def main():
    raise SystemExit("Use consolidate_inventory(period) or run main.py with --period.")

if __name__ == "__main__":
    main()
