from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path
import sys


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

desktop_path = str(_shared.USER_DESKTOP)
amazon_rolling_folder = str(_shared.AMAZON_ROLLING_OUTPUT_FOLDER)
amazon_po_folder = str(_shared.AMAZON_PO_FOLDER)
dp_folders = {name: str(path) for name, path in _shared.AMAZON_ROLLING_DP_FOLDERS.items()}
oracle_ypticod_file = _shared.ORACLE_YPTICOD_FILE
amazon_rolling_cache_dir = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier AmazonRolling\cache")
amazon_rolling_backup_dir = amazon_rolling_cache_dir / "backups"
amazon_po_pickle_file = amazon_rolling_cache_dir / "latest_amazon_po.pkl"
customer_orders_pickle_file = amazon_rolling_cache_dir / "rr_customer_orders.pkl"
units_shipped_pickle_file = amazon_rolling_cache_dir / "rr_units_shipped.pkl"
monthly_sales_parquet_file = (
    _shared.AMAZON_MONTHLY_SALES_ROOT
    / "cache"
    / _shared.AMAZON_MONTHLY_SALES_PARQUET_NAME
)
