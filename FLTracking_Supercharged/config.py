from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
SQL_DIR = BASE_DIR / "sql"

INVENTORY_DETAIL_DIR = Path(r"G:\OPS\Inventory\Daily\Finance_Only")
INVENTORY_DETAIL_GLOB = "Inventory*.xlsx"

AMAZON_PREORDERS_PATH = Path(
    r"G:\SALES\Amazon\PREORDERS\2026\current_amaz_preorders.xlsx"
)
AMAZON_PREORDERS_SHEET = "nyp"

INGRAM_REPORT_DIR = Path(r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Ingram")
INGRAM_REPORT_GLOB = "Daily Report*.xlsx"

BARNES_NOBLE_DIR = Path(r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Barnes & Noble")
BARNES_NOBLE_GLOB = "Week *.xlsx"
