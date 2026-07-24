from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


OVERRIDES_FILE = Path(__file__).with_name("amazon_asin_isbn_overrides.json")
METADATA_COLUMNS = ["ISBN", "Title", "Publisher"]


def normalize_identifier(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def normalize_identifier_series(series: pd.Series) -> pd.Series:
    return series.map(normalize_identifier).astype("string").replace("", pd.NA)


def is_valid_isbn10(value: object) -> bool:
    isbn = normalize_identifier(value).replace("-", "")
    if len(isbn) != 10 or not isbn[:9].isdigit():
        return False
    if not (isbn[-1].isdigit() or isbn[-1] == "X"):
        return False
    digits = [int(char) for char in isbn[:9]]
    digits.append(10 if isbn[-1] == "X" else int(isbn[-1]))
    return sum(weight * digit for weight, digit in zip(range(10, 0, -1), digits)) % 11 == 0


def isbn10_to_isbn13(value: object) -> str:
    isbn10 = normalize_identifier(value).replace("-", "")
    if not is_valid_isbn10(isbn10):
        return ""
    first_twelve = f"978{isbn10[:9]}"
    weighted_sum = sum(
        int(char) * (1 if index % 2 == 0 else 3)
        for index, char in enumerate(first_twelve)
    )
    check_digit = (10 - weighted_sum % 10) % 10
    return f"{first_twelve}{check_digit}"


def load_asin_isbn_overrides(path: Path = OVERRIDES_FILE) -> dict[str, str]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as file:
        values = json.load(file)
    return {
        normalize_identifier(asin): normalize_identifier(isbn)
        for asin, isbn in values.items()
        if normalize_identifier(asin) and normalize_identifier(isbn)
    }


def save_asin_isbn_override(
    asin: str, isbn: str, path: Path = OVERRIDES_FILE
) -> None:
    normalized_asin = normalize_identifier(asin)
    normalized_isbn = normalize_identifier(isbn)
    if not normalized_asin or not normalized_isbn:
        raise ValueError("ASIN and ISBN are both required.")
    values = load_asin_isbn_overrides(path)
    values[normalized_asin] = normalized_isbn
    with path.open("w", encoding="utf-8") as file:
        json.dump(dict(sorted(values.items())), file, indent=2)
        file.write("\n")


def resolve_isbn_series(
    data: pd.DataFrame,
    item_metadata: pd.DataFrame,
    candidate_columns: list[str],
    trusted_columns: list[str] | None = None,
    overrides: dict[str, str] | None = None,
) -> pd.Series:
    """Resolve ISBNs with one shared priority while preserving the input index."""
    asins = normalize_identifier_series(data["ASIN"])
    items = normalize_identifier_series(item_metadata["ISBN"])
    valid_isbns = set(items.dropna())
    manual = overrides if overrides is not None else load_asin_isbn_overrides()
    manual = {
        normalize_identifier(asin): normalize_identifier(isbn)
        for asin, isbn in manual.items()
    }

    resolved = asins.map(manual)
    for column in trusted_columns or []:
        if column not in data.columns:
            continue
        candidate = normalize_identifier_series(data[column])
        resolved = resolved.fillna(candidate)

    for column in candidate_columns:
        if column not in data.columns:
            continue
        candidate = normalize_identifier_series(data[column])
        candidate = candidate.where(candidate.isin(valid_isbns))
        resolved = resolved.fillna(candidate)

    direct_isbn = asins.where(asins.isin(valid_isbns))
    resolved = resolved.fillna(direct_isbn)
    converted_isbn10 = asins.map(isbn10_to_isbn13)
    converted_isbn10 = converted_isbn10.where(converted_isbn10.isin(valid_isbns))
    return resolved.fillna(converted_isbn10).astype("string")


def build_asin_metadata(
    catalog: pd.DataFrame,
    item_metadata: pd.DataFrame,
    overrides: dict[str, str] | None = None,
) -> pd.DataFrame:
    """Resolve ASINs to ISBN/title/publisher with shared Amazon fallback rules."""
    catalog = catalog.copy()
    items = item_metadata.copy()
    for column in ["ASIN", "ISBN-13", "EAN", "Model Number"]:
        catalog[column] = normalize_identifier_series(catalog[column])
    items["ISBN"] = normalize_identifier_series(items["ISBN"])
    items = items.dropna(subset=["ISBN"]).drop_duplicates("ISBN", keep="first")
    resolved = resolve_isbn_series(
        catalog,
        items,
        ["ISBN-13", "EAN", "Model Number"],
        overrides=overrides,
    )

    mapping = pd.DataFrame({"ASIN": catalog["ASIN"], "ISBN": resolved})
    mapping = mapping.dropna(subset=["ASIN"]).drop_duplicates("ASIN", keep="first")
    return mapping.merge(items, on="ISBN", how="left")


def add_amazon_metadata(
    data: pd.DataFrame,
    catalog: pd.DataFrame,
    item_metadata: pd.DataFrame,
    overrides: dict[str, str] | None = None,
) -> pd.DataFrame:
    enriched = data.copy()
    enriched["ASIN"] = normalize_identifier_series(enriched["ASIN"])
    catalog_for_resolution = catalog.copy()
    catalog_asins = set(
        normalize_identifier_series(catalog_for_resolution["ASIN"]).dropna()
    )
    missing_asins = [
        asin for asin in enriched["ASIN"].dropna().unique()
        if asin not in catalog_asins
    ]
    if missing_asins:
        missing_catalog_rows = pd.DataFrame(
            {"ASIN": missing_asins, "ISBN-13": "", "EAN": "", "Model Number": ""}
        )
        catalog_for_resolution = pd.concat(
            [catalog_for_resolution, missing_catalog_rows], ignore_index=True
        )
    mapping = build_asin_metadata(
        catalog_for_resolution, item_metadata, overrides=overrides
    )
    enriched = enriched.drop(columns=METADATA_COLUMNS, errors="ignore")
    enriched = enriched.merge(mapping, on="ASIN", how="left", sort=False)
    source_columns = [
        column for column in data.columns if column not in METADATA_COLUMNS
    ]
    return enriched[[*METADATA_COLUMNS, *source_columns]]
