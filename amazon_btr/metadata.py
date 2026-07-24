from __future__ import annotations

from pathlib import Path
import sys

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from shared.amazon_metadata import (  # noqa: E402
    add_amazon_metadata as _shared_add_amazon_metadata,
    build_asin_metadata as _shared_build_asin_metadata,
)

AMAZON_CATALOG_FOLDER = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Amazon\Catalog"
)

ITEM_SQL = """
SELECT
    CAST(i.ISBN AS varchar(32)) AS ISBN,
    i.SHORT_TITLE AS Title,
    i.PUBLISHER_CODE AS Publisher
FROM ebs.Item i
WHERE i.PRODUCT_TYPE IN ('BK', 'FT', 'DI')
  AND i.AVAILABILITY_STATUS IS NOT NULL;
"""

METADATA_COLUMNS = ["ISBN", "Title", "Publisher"]


def _normalize_identifier(series: pd.Series) -> pd.Series:
    return (
        series.astype("string")
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA})
    )


def load_latest_catalog(catalog_folder: Path) -> pd.DataFrame:
    matches = [
        path
        for path in catalog_folder.glob("*Catalog*csv")
        if not path.name.startswith("~$")
    ]
    if not matches:
        raise FileNotFoundError(
            f"No Amazon Catalog CSV was found in {catalog_folder}."
        )
    catalog_file = max(matches, key=lambda path: path.stat().st_mtime)
    catalog = pd.read_csv(
        catalog_file,
        skiprows=1,
        usecols=["ASIN", "EAN", "ISBN", "Model Number"],
        dtype="string",
        low_memory=False,
    )
    return catalog.rename(columns={"ISBN": "ISBN-13"})


def load_item_metadata() -> pd.DataFrame:
    repo_root = str(Path(__file__).resolve().parents[1])
    if repo_root not in sys.path:
        sys.path.insert(0, repo_root)

    from shared.db import fetch_data_from_db, get_connection

    return fetch_data_from_db(get_connection(), ITEM_SQL)


def build_asin_metadata(
    catalog: pd.DataFrame, item_metadata: pd.DataFrame
) -> pd.DataFrame:
    return _shared_build_asin_metadata(catalog, item_metadata)


def add_title_metadata(
    data: pd.DataFrame, catalog: pd.DataFrame, item_metadata: pd.DataFrame
) -> pd.DataFrame:
    return _shared_add_amazon_metadata(data, catalog, item_metadata)
