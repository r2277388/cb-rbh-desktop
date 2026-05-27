# Cross Gap

Edit `title_groups.json` to add or remove workbook tabs.

Each group can be an explicit ISBN list:

```json
{
  "name": "Dumb Birds",
  "isbns": ["9781452174037", "9781797212272"]
}
```

Or a family-based rule:

```json
{
  "name": "Herve Tullet BB",
  "ip_family_name": "Herve Tullet",
  "formats": ["BB"]
}
```

The same group definitions are used for both the sales query and the Hachette orders query.

Groups can also define supplemental sales-only columns. These columns are added to that
group's tab after the configured ISBN columns and are not included in the open-order query.
They are cached as CSV files in the shared Cross Gap cache folder:

```text
\\sfx\sfny-files\SF\Groups\Sales\2026 Sales Reports\Reports\Cross Gap\cache
```

If a cache file is missing, the report creates it from SQL. Later runs load the cached
CSV instead of re-querying that static period:

```json
{
  "name": "StellaMarigold",
  "isbns": ["9781797233819"],
  "supplemental_sales_columns": [
    {
      "label": "Ivy&Bean_2025",
      "ip_family_name": "Ivy & Bean",
      "start_period": "202501",
      "end_period": "202512"
    }
  ]
}
```
