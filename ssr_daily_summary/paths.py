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


def saved_query_location():
    return str(_shared.SSR_QUERY_OUTPUT_FILE)


def saved_viz_location():
    return str(_shared.SSR_VIZ_OUTPUT_FILE)
