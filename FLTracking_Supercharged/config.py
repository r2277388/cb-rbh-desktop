from pathlib import Path
from importlib.util import module_from_spec, spec_from_file_location


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
SQL_DIR = BASE_DIR / "sql"

def _load_shared_paths():
    shared_path = Path(__file__).resolve().parents[1] / "paths" / "process_paths.py"
    spec = spec_from_file_location("_shared_process_paths", shared_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load shared process paths from {shared_path}")
    module = module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_shared = _load_shared_paths()

INVENTORY_DETAIL_DIR = _shared.INVENTORY_DAILY_FINANCE_ONLY_FOLDER
INVENTORY_DETAIL_GLOB = "Inventory*.xlsx"

AMAZON_PREORDERS_PATH = _shared.CURRENT_AMAZON_PREORDERS_FILE
AMAZON_PREORDERS_SHEET = "nyp"

INGRAM_REPORT_DIR = _shared.INGRAM_DAILY_REPORT_FOLDER
INGRAM_REPORT_GLOB = "Daily Report*.xlsx"

BARNES_NOBLE_DIR = _shared.BN_WEEKLY_REPORT_FOLDER
BARNES_NOBLE_GLOB = "Week *.xlsx"
