import pandas as pd

from shared.amazon_metadata import resolve_isbn_series


def test_trusted_columns_are_preserved_without_item_match():
    data = pd.DataFrame([{"ASIN": "A", "YPTICOD": "TRUSTED"}])
    items = pd.DataFrame([{"ISBN": "DIFFERENT"}])

    result = resolve_isbn_series(
        data,
        items,
        [],
        trusted_columns=["YPTICOD"],
        overrides={},
    )

    assert result.iloc[0] == "TRUSTED"
