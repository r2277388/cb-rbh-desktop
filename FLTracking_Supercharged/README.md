# FLTracking Supercharged

This project builds an ISBN-level tracking file by extracting data from separate Excel and SQL sources, validating each source independently, and then combining them in a main process.

## Current source processes

- `inventory_detail`: pulls selected fields from `Inventory Detail.xlsx`
- `amazon_preorders`: pulls `ISBN` and `Orders` from the `nyp` tab of `current_amaz_preorders.xlsx`
- `amazon_sellthrough`: runs the latest-week Amazon sellthrough SQL and returns ISBN-level shipped/on-hand metrics
- `faire_qty`: runs the Faire sales SQL and returns ISBN-level quantity sold
- `faire_orders`: runs the Faire open orders SQL and returns ISBN-level order quantity
- `ingram_daily_report`: pulls the latest Ingram `Daily Report*.xlsx` file and returns ISBN-level 4-week sales, on-hand, and customer BO metrics
- `barnes_noble_weekly`: pulls the latest Barnes & Noble `Week *.xlsx` file and returns ISBN-level store OH, DC OH, and LTD metrics
- `frontlist_main`: anchors on the latest Frontlist Tracking workbook ISBN list and left-merges all extracted source data by ISBN

## Run the first process

```powershell
python -m processes.inventory_detail
python -m processes.amazon_preorders
python -m processes.amazon_sellthrough
python -m processes.faire_qty
python -m processes.faire_orders
python -m processes.ingram_daily_report
python -m processes.barnes_noble_weekly
python -m processes.frontlist_main
```

## Run the main process

```powershell
python main.py
```

## Notes

- The default source path for Inventory Detail is set in `config.py`.
- The Amazon preorders source reads only the `nyp` worksheet.
- The Ingram source automatically picks the most recently modified `Daily Report*.xlsx` file from the configured folder.
- The Barnes & Noble source automatically picks the most recently modified `Week *.xlsx` file from the configured folder.
- The Frontlist main build automatically picks the most recently modified real Frontlist Tracking workbook and merges all source outputs onto that ISBN list only.
- SQL for source-specific database extracts is stored in the local `sql` folder.
- ISBN normalization is centralized in `isbn_utils.py`, including the reusable 12-digit exception query.
- Intermediate source outputs are saved as pickle files; `frontlist_main` is saved as Excel.
- Running `frontlist_main` opens a Save As dialog for the final Excel output and falls back to the local `output` folder if you cancel.
- Each process is designed to be runnable by itself so results can be inspected before combining everything.
