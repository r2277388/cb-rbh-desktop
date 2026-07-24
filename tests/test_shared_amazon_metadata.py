import json

import pandas as pd

from shared.amazon_metadata import (
    build_asin_metadata,
    isbn10_to_isbn13,
    is_valid_isbn10,
    load_asin_isbn_overrides,
    save_asin_isbn_override,
)


def test_valid_isbn10_is_converted_to_isbn13():
    assert is_valid_isbn10("1761212990")
    assert isbn10_to_isbn13("1761212990") == "9781761212994"


def test_asin_field_itself_can_resolve_as_isbn10():
    catalog = pd.DataFrame(
        [{"ASIN": "1761212990", "ISBN-13": "", "EAN": "", "Model Number": ""}]
    )
    items = pd.DataFrame(
        [
            {
                "ISBN": "9781761212994",
                "Title": "Resolved Title",
                "Publisher": "Resolved Publisher",
            }
        ]
    )

    result = build_asin_metadata(catalog, items, overrides={}).iloc[0]

    assert result["ISBN"] == "9781761212994"
    assert result["Title"] == "Resolved Title"
    assert result["Publisher"] == "Resolved Publisher"


def test_manual_overrides_are_persistent_and_take_priority(tmp_path):
    path = tmp_path / "overrides.json"
    path.write_text(json.dumps({"ASIN1": "111"}), encoding="utf-8")

    save_asin_isbn_override("asin1", "222", path)

    assert load_asin_isbn_overrides(path) == {"ASIN1": "222"}
