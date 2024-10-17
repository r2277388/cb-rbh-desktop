## Weekly Amazon PO File Processing

### How It Works:

1. **Data Collection**: Five files collect the dataframes needed to create the weekly report:
   - `cleanup_catalog.py`
   - `cleanup_ebs_item.py`
   - `cleanup_inventorydetail.py`
   - `cleanup_po.py`
   - `cleanup_ypticod.py`

2. **Accessing SQL**: The `sql_queries/queries.py` consists of a function that provides the SQL code for running the `cleanup_ebs_item.py` file since this process accesses our SQL server.

3. **Data Cleaning and Combination**: The `combined_file.py` script combines and cleans all of the above files. It creates a combined file with the appropriate ISBNs and PO information.

4. **Report Creation**: The `datadump_cb.py` script uses the cleaned file to create individual views for the weekly report.

5. **Excel Report Creation**: The `excel_report.py` file creates the version that is saved off to the Amazon/PURCHASE ORDERS/YYYY folder.

6. **Last Step**: The `main.py` script will bring everything together and allow us to download the work to excel.