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
