# Shared Amazon ASIN/ISBN metadata

Use `shared.amazon_metadata` whenever an Amazon process needs to resolve an
ASIN to ISBN, Title, and Publisher.

Resolution priority:

1. Shared manual override in `amazon_asin_isbn_overrides.json`
2. Amazon Catalog ISBN-13
3. Amazon Catalog EAN
4. Amazon Catalog Model Number
5. ASIN itself when it is already a valid item-table ISBN
6. ASIN converted from a valid ISBN-10 to ISBN-13

Title and Publisher are joined from `ebs.Item` after ISBN resolution.

Core functions:

- `resolve_isbn_series(data, item_metadata, candidate_columns)`
- `build_asin_metadata(catalog, item_metadata)`
- `add_amazon_metadata(data, catalog, item_metadata)`
- `load_asin_isbn_overrides()`
- `save_asin_isbn_override(asin, isbn)`

`resolve_isbn_series` also supports `trusted_columns` for sources such as
YPTICOD or an ISBN already accepted by an existing report. Trusted values are
preserved without item-table validation; Catalog fallback candidates continue
to require an item-table match.

Current consumers:

- Amazon BTR
- Amazon PreOrders
- Amazon Customer Orders
- Amazon SQL Upload
- Amazon monthly rolling sales
- Amazon AMS monthly campaigns

The legacy `amazon_sql_upload/asin_manual_key.py` module reads the shared JSON
file, so existing Amazon procedures remain compatible.
