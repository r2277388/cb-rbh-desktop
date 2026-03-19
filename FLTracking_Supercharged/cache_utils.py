import json
from pathlib import Path
from typing import Callable

import pandas as pd


def build_source_signature(source_path: Path, extra: dict | None = None) -> dict:
    stat = source_path.stat()
    signature = {
        "source_path": str(source_path),
        "mtime_ns": stat.st_mtime_ns,
        "size": stat.st_size,
    }
    if extra:
        signature.update(extra)
    return signature


def load_cached_dataframe(
    cache_path: Path,
    signature: dict,
    build_func: Callable[[], pd.DataFrame],
) -> tuple[pd.DataFrame, bool]:
    meta_path = cache_path.with_suffix(".json")

    if cache_path.exists() and meta_path.exists():
        try:
            cached_signature = json.loads(meta_path.read_text(encoding="utf-8"))
            if cached_signature == signature:
                return pd.read_pickle(cache_path), True
        except Exception:
            pass

    df = build_func()
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_pickle(cache_path)
    meta_path.write_text(json.dumps(signature, indent=2), encoding="utf-8")
    return df, False
