# Faire Customer Match

This utility matches Faire customer rows to active `ebs.customer` rows and returns the full `PARTYSITENUMBER`.

The workflow is intentionally two-step:

1. Pull active EBS customers into a local CSV cache.
2. Match Faire workbooks against that cache without repeatedly querying the database.

## Commands

Refresh the active customer cache:

```powershell
venv\Scripts\python -m faire_customer_match.main refresh-cache
```

Match the default Faire workbook:

```powershell
venv\Scripts\python -m faire_customer_match.main match
```

Do both in one command:

```powershell
venv\Scripts\python -m faire_customer_match.main run
```

Defaults:

- Input: `Faire_to_HBG_Number_Search.xlsx`
- Cache: `faire_customer_match\active_ebs_customers.csv`
- Output: `Faire_to_HBG_Number_Search_matched.xlsx`

## Matching Logic

The script blocks candidates first so it avoids comparing every Faire row to every active customer.

Candidate blocks use:

- ZIP + state + street number + street key
- ZIP + state + city
- State + street number + street key

Within those small candidate groups, it scores:

- Address similarity
- Account name similarity
- City similarity
- ZIP, state, and country agreement

Rows are labeled `High`, `Medium`, `Review`, or `No candidate`. Treat `Review` rows as a manual queue.
