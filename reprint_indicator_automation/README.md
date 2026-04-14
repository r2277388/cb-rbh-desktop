# Reprint Indicator Automation

This script automates the manual Excel workflow for the Reprint Indicator file:

1. Open the template workbook in Excel
2. Refresh all tables/connections/queries
3. Rebuild `BL_Detail` from `MetaData` rows whose `Release_Group` contains `Backlist`
4. Rebuild `FL_Detail` from `MetaData` rows whose `Release_Group` does not contain `Backlist`
5. Copy sheets from `RPG_Risk_Analyzer` through `Explanation` into a new workbook
6. Save the detached workbook into `G:\OPS\Reprint Indicators\2026`
7. Name the file `reprint indicator_YYYY_MM_DD.xlsx`
8. Break external links and remove workbook connections in the saved copy
9. Close Excel so the saved file is editable

When you choose to refresh, the script can also save the refreshed/rebuilt state back to
`Templates\Reprint indicator TEMPLATE.xlsx` before it creates the detached dated copy.

## Run

```powershell
python reprint_indicator_automation\main.py
```

## Optional

Show Excel while it runs:

```powershell
python reprint_indicator_automation\main.py --visible
```

Use a specific output date:

```powershell
python reprint_indicator_automation\main.py --date 2026_03_30
```

Always save the refreshed/rebuilt results back to the template without the extra prompt:

```powershell
python reprint_indicator_automation\main.py --save-template
```

## Notes

- This uses Excel COM automation through `pywin32`.
- It requires Microsoft Excel to be installed on the machine.
- If a file for the same date already exists, the script adds `_v1`, `_v2`, and so on.
- `BL_Detail` data starts at row 3. ISBNs are written into column `A`, and the template row is autofilled across `A:BO`.
- `FL_Detail` data starts at row 3. `Release_Group` and `ISBN` are written into columns `A:B`, and the template row is autofilled across `A:BY`.
- The export sheet range is currently defined by sheet order from `RPG_Risk_Analyzer` through `Explanation`.
