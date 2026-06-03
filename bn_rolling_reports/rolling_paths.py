from __future__ import annotations

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

bn_rolling_folder = Path(_shared.BN_WEEKLY_REPORT_FOLDER)
bn_dp_folders = {name: Path(path) for name, path in _shared.BN_ROLLING_DP_FOLDERS.items()}
cache_dir = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier BarnesNoble\cache")
local_review_dir = Path(__file__).resolve().parent / "review_output"
sales_cache_file = cache_dir / "bn_customer_sales.parquet"
inventory_cache_file = cache_dir / "bn_inventory_snapshots.parquet"
manual_missing_weeks_file = cache_dir / "bn_manual_missing_weeks.parquet"
