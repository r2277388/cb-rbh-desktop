from pathlib import Path
from importlib.util import module_from_spec, spec_from_file_location


def _load_shared_paths():
    shared_path = Path(__file__).resolve().parents[1] / "paths" / "process_paths.py"
    spec = spec_from_file_location("_shared_process_paths", shared_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load shared process paths from {shared_path}")
    module = module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_shared = _load_shared_paths()

BASE_FOLDER = _shared.BN_RAW_BASE_FOLDER
RAW_FOLDER_SUFFIX = "_raw_files"
OUTPUT_PREFIX = "pos_combined"
SALES_PREFIX = "sales_working"
INVENTORY_PREFIX = "inventory_working"
REQUIRED_FILE_KEYWORDS = ("toy", "gift", "cal")
REQUIRED_COLUMNS = [
    "EAN",
    "Imprint",
    "OH",
    "DC_To_Stores(OO)",
    "LW",
    "YTD",
    "DC_OH_Tot",
    "DC_OO_Tot",
]
TEXT_COLUMNS = ["EAN", "Imprint"]
NUMERIC_COLUMNS = [
    "OH",
    "DC_To_Stores(OO)",
    "LW",
    "YTD",
    "DC_OH_Tot",
    "DC_OO_Tot",
]
OUTPUT_COLUMN_RENAMES = {
    "EAN": "ISBN",
    "OH": "OH_Stores",
    "DC_To_Stores(OO)": "OO_Stores",
    "DC_OH_Tot": "OH_DC",
    "DC_OO_Tot": "OO_DC",
}
SALES_HEADER_ROW = 6
SALES_TOTAL_ROW = 7
SALES_DATA_START_ROW = 8
SALES_REQUIRED_HEADERS = [
    "ISBN",
    "Title",
    "Author",
    "Pub Date",
    "List Price Value",
    "B&N Subject Code",
    "B&N Dept Code",
    "Barnes & Noble Total",
    "BN College",
    "BN.com",
    "Barnes & Noble",
    "Barnes & Noble Total",
    "BN College",
    "BN.com",
    "Barnes & Noble",
    "Barnes & Noble Total",
    "BN College",
    "BN.com",
    "Barnes & Noble",
    "Barnes & Noble Total",
    "BN College",
    "BN.com",
    "Barnes & Noble",
]
INVENTORY_HEADER_ROW = 5
INVENTORY_TOTAL_ROW = 6
INVENTORY_DATA_START_ROW = 7
INVENTORY_REQUIRED_HEADERS = [
    "ISBN",
    "Title",
    "Author",
    "Pub Date",
    "List Price Value",
    "BN College",
    "BN Distribution Center",
    "BN.com",
    "Barnes & Noble",
    "BN College",
    "BN Distribution Center",
    "BN.com",
    "Barnes & Noble",
    "BN College",
    "BN Distribution Center",
    "BN.com",
    "Barnes & Noble",
    "BN College",
    "BN Distribution Center",
    "BN.com",
    "Barnes & Noble",
    "BN College",
    "BN Distribution Center",
    "BN.com",
    "Barnes & Noble",
    "BN College",
    "BN Distribution Center",
    "BN.com",
    "Barnes & Noble",
    "BN College",
    "BN Distribution Center",
    "BN.com",
    "Barnes & Noble",
]
