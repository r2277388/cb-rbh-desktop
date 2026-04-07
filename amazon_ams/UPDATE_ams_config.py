from __future__ import annotations

import re
from pathlib import Path


REPORTS_ROOT = Path(
    r"G:\SALES\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN"
)
DEFAULT_TAB = "USE_main"
DEFAULT_SKIPROWS = 1

# Only add entries here when auto-discovery needs help for a specific month.
MONTH_OVERRIDES = {
    # "2026-03": {
    #     "file": r"G:\...\2026 - 03 - March - Performance by ASIN_ALL.xlsx",
    #     "tab": "USE_main",
    #     "skiprows": 1,
    # },
}

# Add YYYY-MM values here if a month should be skipped from processing.
IGNORED_MONTHS = set()

MONTH_PATTERN = re.compile(r"(?P<year>20\d{2})\s*-\s*(?P<month>\d{2})")


def _candidate_files() -> list[Path]:
    if not REPORTS_ROOT.exists():
        return []

    candidates: list[Path] = []
    for path in REPORTS_ROOT.rglob("*.xlsx"):
        name = path.name.lower()
        if path.name.startswith("~$"):
            continue
        if "performance by asin" not in name:
            continue
        candidates.append(path)
    return candidates


def _extract_month(path: Path) -> str | None:
    match = MONTH_PATTERN.search(path.stem)
    if not match:
        return None
    return f"{match.group('year')}-{match.group('month')}"


def _path_mtime(path: Path) -> float:
    try:
        return path.stat().st_mtime
    except OSError:
        return -1


def discover_monthly_files() -> dict[str, dict[str, object]]:
    discovered: dict[str, dict[str, object]] = {}

    for path in sorted(_candidate_files()):
        month = _extract_month(path)
        if month is None or month in IGNORED_MONTHS:
            continue

        existing = discovered.get(month)
        if existing is not None:
            existing_path = Path(str(existing["file"]))
            if _path_mtime(existing_path) >= _path_mtime(path):
                continue

        discovered[month] = {
            "tab": DEFAULT_TAB,
            "skiprows": DEFAULT_SKIPROWS,
            "file": str(path),
        }

    return discovered


def build_tab_dict() -> dict[str, dict[str, object]]:
    tab_dict = discover_monthly_files()

    for month in IGNORED_MONTHS:
        tab_dict.pop(month, None)

    for month, override in MONTH_OVERRIDES.items():
        if month in IGNORED_MONTHS:
            continue
        tab_dict[month] = {
            "tab": override.get("tab", DEFAULT_TAB),
            "skiprows": int(override.get("skiprows", DEFAULT_SKIPROWS)),
            "file": str(override["file"]),
        }

    return dict(sorted(tab_dict.items()))


tab_dict = build_tab_dict()
month_list = sorted(tab_dict.keys())
