"""Backward-compatible access to the shared Amazon ASIN/ISBN overrides."""

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from shared.amazon_metadata import load_asin_isbn_overrides  # noqa: E402

asin_isbn_manual_key = load_asin_isbn_overrides()
