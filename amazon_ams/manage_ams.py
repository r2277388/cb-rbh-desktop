import importlib
import re
import subprocess
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

import UPDATE_ams_config as ams_config
from processor import run_full_rebuild, run_incremental_update


CONFIG_PATH = Path(__file__).with_name("UPDATE_ams_config.py")
BASELINE_YEAR = "2025"


def reload_config():
    global ams_config
    ams_config = importlib.reload(ams_config)
    return ams_config


def write_config(month_overrides, ignored_months):
    lines = [
        "from __future__ import annotations",
        "",
        "import re",
        "from pathlib import Path",
        "",
        "",
        "REPORTS_ROOT = Path(",
        '    r"G:\\SALES\\Amazon\\AMAZON ADVERTISING\\MONTHLY REPORTS\\MONTHLY REPORTS - PERFORMANCE BY ASIN"',
        ")",
        'DEFAULT_TAB = "USE_main"',
        "DEFAULT_SKIPROWS = 1",
        "",
        "# Only add entries here when auto-discovery needs help for a specific month.",
        "MONTH_OVERRIDES = {",
    ]

    for month in sorted(month_overrides):
        entry = month_overrides[month]
        tab = entry.get("tab", "USE_main")
        skiprows = int(entry.get("skiprows", 1))
        file_path = str(entry.get("file", ""))
        lines.extend(
            [
                f'    "{month}": {{',
                f'        "file": r"{file_path}",',
                f'        "tab": "{tab}",',
                f'        "skiprows": {skiprows},',
                "    },",
            ]
        )

    lines.extend(
        [
            "}",
            "",
            "# Add YYYY-MM values here if a month should be skipped from processing.",
            f"IGNORED_MONTHS = {repr(set(sorted(ignored_months)))}",
            "",
            'MONTH_PATTERN = re.compile(r"(?P<year>20\\d{2})\\s*-\\s*(?P<month>\\d{2})")',
            "",
            "",
            "def _candidate_files() -> list[Path]:",
            "    if not REPORTS_ROOT.exists():",
            "        return []",
            "",
            "    candidates: list[Path] = []",
            '    for path in REPORTS_ROOT.rglob("*.xlsx"):',
            "        name = path.name.lower()",
            '        if path.name.startswith("~$"):',
            "            continue",
            '        if "performance by asin" not in name:',
            "            continue",
            "        candidates.append(path)",
            "    return candidates",
            "",
            "",
            "def _extract_month(path: Path) -> str | None:",
            "    match = MONTH_PATTERN.search(path.stem)",
            "    if not match:",
            "        return None",
            '    return f"{match.group(\'year\')}-{match.group(\'month\')}"',
            "",
            "",
            "def _path_mtime(path: Path) -> float:",
            "    try:",
            "        return path.stat().st_mtime",
            "    except OSError:",
            "        return -1",
            "",
            "",
            "def discover_monthly_files() -> dict[str, dict[str, object]]:",
            "    discovered: dict[str, dict[str, object]] = {}",
            "",
            "    for path in sorted(_candidate_files()):",
            "        month = _extract_month(path)",
            "        if month is None or month in IGNORED_MONTHS:",
            "            continue",
            "",
            "        existing = discovered.get(month)",
            "        if existing is not None:",
            '            existing_path = Path(str(existing["file"]))',
            "            if _path_mtime(existing_path) >= _path_mtime(path):",
            "                continue",
            "",
            '        discovered[month] = {"tab": DEFAULT_TAB, "skiprows": DEFAULT_SKIPROWS, "file": str(path)}',
            "",
            "    return discovered",
            "",
            "",
            "def build_tab_dict() -> dict[str, dict[str, object]]:",
            "    tab_dict = discover_monthly_files()",
            "",
            "    for month in IGNORED_MONTHS:",
            "        tab_dict.pop(month, None)",
            "",
            "    for month, override in MONTH_OVERRIDES.items():",
            "        if month in IGNORED_MONTHS:",
            "            continue",
            '        tab_dict[month] = {"tab": override.get("tab", DEFAULT_TAB), "skiprows": int(override.get("skiprows", DEFAULT_SKIPROWS)), "file": str(override["file"])}',
            "",
            "    return dict(sorted(tab_dict.items()))",
            "",
            "",
            "tab_dict = build_tab_dict()",
            "month_list = sorted(tab_dict.keys())",
            "",
        ]
    )

    CONFIG_PATH.write_text("\n".join(lines), encoding="utf-8")


def read_columns_from_entry(entry):
    df = pd.read_excel(
        entry["file"],
        sheet_name=entry["tab"],
        skiprows=entry["skiprows"],
        header=0,
        engine="openpyxl",
    )
    df.columns = df.columns.str.strip().str.lower()
    return set(df.columns)


def baseline_columns(tab_dict):
    baseline_months = sorted([m for m in tab_dict if m.startswith(f"{BASELINE_YEAR}-")])
    if not baseline_months:
        raise RuntimeError(f"No baseline months found for {BASELINE_YEAR}.")

    month_to_cols = {}
    for month in baseline_months:
        month_to_cols[month] = read_columns_from_entry(tab_dict[month])

    signatures = Counter(tuple(sorted(cols)) for cols in month_to_cols.values())
    expected_tuple, _ = signatures.most_common(1)[0]
    expected = set(expected_tuple)

    if len(signatures) > 1:
        print("WARNING: baseline year has more than one column shape.")
        print("Using most common baseline schema for validation.")

    return expected


def validate_month_columns(month, entry, tab_dict):
    expected = baseline_columns(tab_dict)
    actual = read_columns_from_entry(entry)

    missing = sorted(expected - actual)
    extra = sorted(actual - expected)

    print("\nExpected columns (baseline):")
    print(sorted(expected))
    print("\nActual columns (candidate file):")
    print(sorted(actual))

    if not missing and not extra:
        print(f"\nValidation passed for {month}: columns match baseline.")
        return True

    print(f"\nValidation failed for {month}:")
    print(f"Missing columns: {missing}")
    print(f"Extra columns: {extra}")
    return False


def print_last_10_months():
    config = reload_config()
    last_10 = sorted(config.month_list, reverse=True)[:10]
    print("\nLast 10 discovered/active months:")
    for month in last_10:
        entry = config.tab_dict[month]
        override_label = " override" if month in config.MONTH_OVERRIDES else ""
        print(f"{month}{override_label}: {entry['file']}")


def ignore_month():
    config = reload_config()
    print("\nConfigured months:")
    print(", ".join(sorted(config.month_list)))
    month = input("Enter month to ignore (YYYY-MM): ").strip()
    if month not in config.tab_dict and month not in config.MONTH_OVERRIDES:
        print(f"Month not found: {month}")
        return

    confirm = input(f"Ignore {month} for future processing? (y/n): ").strip().lower()
    if confirm not in {"y", "yes"}:
        print("Cancelled.")
        return

    overrides = dict(config.MONTH_OVERRIDES)
    overrides.pop(month, None)
    ignored = set(config.IGNORED_MONTHS)
    ignored.add(month)
    write_config(overrides, ignored)
    print(f"{month} will be skipped in future processing.")


def add_or_update_month_override():
    config = reload_config()

    month = input("Month (YYYY-MM): ").strip()
    if not re.fullmatch(r"\d{4}-\d{2}", month):
        print("Invalid month format. Use YYYY-MM.")
        return

    tab = input('Tab name [default USE_main]: ').strip() or "USE_main"
    skiprows_raw = input("Rows to skip [default 1]: ").strip() or "1"
    try:
        skiprows = int(skiprows_raw)
    except ValueError:
        print("Rows to skip must be an integer.")
        return

    default_path = ""
    if month in config.tab_dict:
        default_path = str(config.tab_dict[month]["file"])

    prompt = "Full file path"
    if default_path:
        prompt += f" [default {default_path}]"
    prompt += ": "
    file_path = input(prompt).strip() or default_path
    if not file_path:
        print("File path is required.")
        return
    if not Path(file_path).exists():
        print(f"File not found: {file_path}")
        return

    candidate_entry = {"tab": tab, "skiprows": skiprows, "file": file_path}
    validation_tab_dict = dict(config.tab_dict)
    validation_tab_dict[month] = candidate_entry

    print("\nRunning column validation before saving override...")
    ok = validate_month_columns(month, candidate_entry, validation_tab_dict)
    if not ok:
        print("Override not saved due to column mismatch.")
        return

    overrides = dict(config.MONTH_OVERRIDES)
    overrides[month] = candidate_entry
    ignored = set(config.IGNORED_MONTHS)
    ignored.discard(month)
    write_config(overrides, ignored)
    print(f"Saved override for {month} in UPDATE_ams_config.py")


def validate_latest_month():
    config = reload_config()
    if not config.month_list:
        print("No months available.")
        return

    month = sorted(config.month_list)[-1]
    print(f"Validating latest month: {month}")
    validate_month_columns(month, config.tab_dict[month], config.tab_dict)


def run_ams_full():
    print("Running full amazon_ams process...")
    subprocess.run([sys.executable, str(Path(__file__).with_name("main.py"))], check=True)


def run_ams_incremental():
    config = reload_config()
    if not config.month_list:
        print("No months available.")
        return

    default_month = sorted(config.month_list)[-1]
    month = input(f"Month to process [default {default_month}]: ").strip() or default_month
    if month not in config.tab_dict:
        print(f"Month not found: {month}")
        return

    print(f"Processing incremental month: {month}")
    _, archived = run_incremental_update(month)
    if archived:
        print("Archived previous outputs:")
        for path in archived:
            print(f"- {path}")


def run_ams():
    print("\nRun AMS")
    print()
    print("    1. Incremental (append/replace one month only)")
    print("    2. Full rerun (all discovered months)")
    print("    3. Back")
    print()
    choice = input("Choose an option: ").strip().lower()
    if choice == "1":
        run_ams_incremental()
    elif choice == "2":
        run_ams_full()
    elif choice in {"3", "back", "b"}:
        return
    else:
        print("Invalid choice.")


def main():
    while True:
        print("\nAmazon AMS Manager")
        print()
        print("    1. Show last 10 discovered months")
        print("    2. Ignore a month from processing")
        print("    3. Add/update month override")
        print("    4. Validate latest month against baseline")
        print("    5. Run AMS processing")
        print("    6. Back to launcher")
        print()
        choice = input("Choose an option: ").strip().lower()

        try:
            if choice == "1":
                print_last_10_months()
            elif choice == "2":
                ignore_month()
            elif choice == "3":
                add_or_update_month_override()
            elif choice == "4":
                validate_latest_month()
            elif choice == "5":
                run_ams()
            elif choice in {"6", "back", "b", "exit", "quit", "q"}:
                return
            else:
                print("Invalid choice.")
        except Exception as exc:
            print(f"Operation failed: {exc}")


if __name__ == "__main__":
    main()
