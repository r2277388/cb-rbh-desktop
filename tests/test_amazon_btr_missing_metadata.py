import pandas as pd

from amazon_btr.missing_metadata import (
    amazon_product_url,
    find_unresolved_asins,
    review_unresolved_asins,
)


def test_unresolved_asins_and_amazon_urls_are_reported():
    enriched = pd.DataFrame(
        [
            {"ASIN": "B000000001", "ISBN": pd.NA},
            {"ASIN": "B000000002", "ISBN": "9780000000001"},
        ]
    )

    assert find_unresolved_asins(enriched) == ["B000000001"]
    assert amazon_product_url("B000000001") == "https://www.amazon.com/dp/B000000001"


def test_interactive_review_saves_supplied_isbn():
    enriched = pd.DataFrame([{"ASIN": "B000000001", "ISBN": pd.NA}])
    saved = []

    count = review_unresolved_asins(
        enriched,
        prompt=True,
        input_func=lambda _: "9780000000001",
        save_func=lambda asin, isbn: saved.append((asin, isbn)),
    )

    assert count == 1
    assert saved == [("B000000001", "9780000000001")]
