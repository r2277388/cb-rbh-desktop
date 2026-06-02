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

DOWNLOADS_FOLDER = _shared.DOWNLOADS_FOLDER
AMAZON_PREORDERS_OUTPUT_FOLDER = _shared.AMAZON_PREORDERS_OUTPUT_FOLDER
ATELIER_AMAZON_CATALOG_FOLDER = _shared.ATELIER_AMAZON_CATALOG_FOLDER
ATELIER_AMAZON_INVENTORY_FOLDER = _shared.ATELIER_AMAZON_INVENTORY_FOLDER
