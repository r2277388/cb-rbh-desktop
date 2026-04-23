from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path


def _load_shared_paths():
    shared_path = Path(__file__).resolve().parents[1] / "paths" / "process_paths.py"
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
amazon_rolling_cache_dir = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier AmazonRolling\cache")
amazon_rolling_backup_dir = amazon_rolling_cache_dir / "backups"
amazon_po_pickle_file = amazon_rolling_cache_dir / "latest_amazon_po.pkl"
customer_orders_pickle_file = amazon_rolling_cache_dir / "rr_customer_orders.pkl"
units_shipped_pickle_file = amazon_rolling_cache_dir / "rr_units_shipped.pkl"
