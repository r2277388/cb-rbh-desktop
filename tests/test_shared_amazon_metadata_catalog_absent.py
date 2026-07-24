import pandas as pd

from shared.amazon_metadata import add_amazon_metadata


def test_isbn10_asin_resolves_even_when_absent_from_catalog():
    data = pd.DataFrame([{"ASIN": "1761212990", "Status": "Accepted"}])
    catalog = pd.DataFrame(
        [
            {
                "ASIN": "SOME_OTHER_ASIN",
                "ISBN-13": "",
                "EAN": "",
                "Model Number": "",
            }
        ]
    )
    items = pd.DataFrame(
        [
            {
                "ISBN": "9781761212994",
                "Title": "The Real Title",
                "Publisher": "The Real Publisher",
            }
        ]
    )

    result = add_amazon_metadata(data, catalog, items, overrides={}).iloc[0]

    assert result["ISBN"] == "9781761212994"
    assert result["Title"] == "The Real Title"
    assert result["Publisher"] == "The Real Publisher"
