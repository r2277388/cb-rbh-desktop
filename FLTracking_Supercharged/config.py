from pathlib import Path
from importlib.util import module_from_spec, spec_from_file_location
import sys


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier FL_SuperCharged\cache")
FRONTLIST_MAIN_OUTPUT_DIR = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier FL_SuperCharged"
)
SQL_DIR = BASE_DIR / "sql"

def _load_shared_paths():
    repo_root = Path(__file__).resolve().parents[1]
    shared_path = repo_root / "paths" / "process_paths.py"
    repo_root_text = str(repo_root)
    if repo_root_text not in sys.path:
        sys.path.insert(0, repo_root_text)
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

BOOKSHOP_PREORDERS_DIR = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Bookshop\Bookshop_PreOrders"
)
BOOKSHOP_PREORDERS_GLOB = "*.csv"

INGRAM_REPORT_DIR = _shared.INGRAM_DAILY_REPORT_FOLDER
INGRAM_REPORT_GLOB = "Daily Report*.xlsx"

BARNES_NOBLE_DIR = _shared.BN_WEEKLY_REPORT_FOLDER
BARNES_NOBLE_GLOB = "Week *.xlsx"
