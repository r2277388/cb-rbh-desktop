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

folder_path = str(_shared.UK_ROLLING_SOURCE_FOLDER)
output_path = str(_shared.UK_ROLLING_OUTPUT_FILE)
