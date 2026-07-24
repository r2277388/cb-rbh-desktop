from __future__ import annotations

from collections.abc import Callable

import pandas as pd

from shared.amazon_metadata import normalize_identifier, save_asin_isbn_override


def find_unresolved_asins(enriched: pd.DataFrame) -> list[str]:
    unresolved = enriched.loc[enriched["ISBN"].isna(), "ASIN"]
    return sorted(
        {
            normalize_identifier(value)
            for value in unresolved
            if normalize_identifier(value)
        }
    )


def amazon_product_url(asin: str) -> str:
    return f"https://www.amazon.com/dp/{normalize_identifier(asin)}"


def review_unresolved_asins(
    enriched: pd.DataFrame,
    *,
    prompt: bool,
    input_func: Callable[[str], str] = input,
    save_func: Callable[[str, str], None] = save_asin_isbn_override,
) -> int:
    unresolved = find_unresolved_asins(enriched)
    if not unresolved:
        return 0

    print("\nUnable to convert these ASINs to ISBNs:")
    for asin in unresolved:
        print(f"  ASIN: {asin}")
        print(f"  Amazon: {amazon_product_url(asin)}")

    if not prompt:
        print("Run interactively to add ISBNs to the shared manual override table.")
        return 0

    saved_count = 0
    print("\nEnter the corresponding ISBN when found.")
    print("Press Enter without an ISBN to leave an ASIN unresolved and continue.")
    for asin in unresolved:
        isbn = input_func(f"ISBN for {asin}: ").strip()
        if not isbn:
            continue
        save_func(asin, isbn)
        print(f"Saved shared override: {asin} -> {isbn}")
        saved_count += 1
    return saved_count
