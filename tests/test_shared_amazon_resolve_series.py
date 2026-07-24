import pandas as pd

from shared.amazon_metadata import resolve_isbn_series


def test_shared_series_priority_and_unresolved_values():
    data = pd.DataFrame(
        [
            {"ASIN": "MANUAL", "YPTICOD": "111", "EAN": "222"},
            {"ASIN": "CATALOG", "YPTICOD": "", "EAN": "222"},
            {"ASIN": "1761212990", "YPTICOD": "", "EAN": ""},
            {"ASIN": "UNKNOWN", "YPTICOD": "", "EAN": ""},
        ]
    )
    items = pd.DataFrame(
        [
            {"ISBN": "111"},
            {"ISBN": "222"},
            {"ISBN": "9781761212994"},
            {"ISBN": "999"},
        ]
    )

    result = resolve_isbn_series(
        data,
        items,
        ["YPTICOD", "EAN"],
        overrides={"MANUAL": "999"},
    )

    assert result.iloc[0] == "999"
    assert result.iloc[1] == "222"
    assert result.iloc[2] == "9781761212994"
    assert pd.isna(result.iloc[3])
