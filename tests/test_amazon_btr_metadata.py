import pandas as pd

from amazon_btr.metadata import add_title_metadata, build_asin_metadata


def test_asin_mapping_uses_isbn_then_ean_then_model_number():
    catalog = pd.DataFrame(
        [
            {"ASIN": "A", "ISBN-13": "111", "EAN": "222", "Model Number": "333"},
            {"ASIN": "B", "ISBN-13": "BAD", "EAN": "222", "Model Number": "333"},
            {"ASIN": "C", "ISBN-13": "BAD", "EAN": "BAD", "Model Number": "333"},
        ]
    )
    items = pd.DataFrame(
        [
            {"ISBN": "111", "Title": "First", "Publisher": "P1"},
            {"ISBN": "222", "Title": "Second", "Publisher": "P2"},
            {"ISBN": "333", "Title": "Third", "Publisher": "P3"},
        ]
    )

    mapped = build_asin_metadata(catalog, items).set_index("ASIN")

    assert mapped.loc["A", "ISBN"] == "111"
    assert mapped.loc["B", "ISBN"] == "222"
    assert mapped.loc["C", "ISBN"] == "333"


def test_metadata_columns_are_added_first_without_losing_rows():
    raw = pd.DataFrame(
        [
            {"ASIN": "A", "Status": "Accepted"},
            {"ASIN": "UNKNOWN", "Status": "Accepted"},
        ]
    )
    catalog = pd.DataFrame(
        [{"ASIN": "A", "ISBN-13": "111", "EAN": "", "Model Number": ""}]
    )
    items = pd.DataFrame(
        [{"ISBN": "111", "Title": "A Title", "Publisher": "Chronicle"}]
    )

    enriched = add_title_metadata(raw, catalog, items)

    assert enriched.columns[:3].tolist() == ["ISBN", "Title", "Publisher"]
    assert len(enriched) == 2
    assert enriched.loc[0, "Title"] == "A Title"
    assert pd.isna(enriched.loc[1, "ISBN"])
