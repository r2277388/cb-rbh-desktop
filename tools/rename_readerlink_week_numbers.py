from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from shared.bookscan_calendar import bookscan_week


READERLINK_ROOT = Path(r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Readerlink")
READERLINK_PATTERN = re.compile(
    r"^Week (?P<week>\d{1,2}) - (?P<year>\d{4}) "
    r"Rolling Readerlink \((?P<date>\d{6})\)(?P<suffix>.*)\.xlsx$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class RenamePlan:
    source: Path
    target: Path
    old_week: str
    new_week: str
    old_year: str
    new_year: str


def planned_rename(path: Path) -> RenamePlan | None:
    match = READERLINK_PATTERN.match(path.name)
    if match is None:
        return None

    week_end = pd.to_datetime(match.group("date"), format="%m%d%y", errors="raise")
    bookscan = bookscan_week(week_end)
    old_week = match.group("week")
    old_year = match.group("year")
    new_week = f"{bookscan.week:02d}"
    new_year = str(bookscan.year)

    if old_week == new_week and old_year == new_year:
        return None

    new_name = READERLINK_PATTERN.sub(
        rf"Week {new_week} - {new_year} Rolling Readerlink (\g<date>)\g<suffix>.xlsx",
        path.name,
    )
    return RenamePlan(
        source=path,
        target=path.with_name(new_name),
        old_week=old_week,
        new_week=new_week,
        old_year=old_year,
        new_year=new_year,
    )


def collect_plans() -> list[RenamePlan]:
    if not READERLINK_ROOT.exists():
        raise FileNotFoundError(f"Readerlink report folder not found: {READERLINK_ROOT}")
    plans: list[RenamePlan] = []
    for path in READERLINK_ROOT.rglob("*.xlsx"):
        plan = planned_rename(path)
        if plan is not None:
            plans.append(plan)
    return sorted(plans, key=lambda item: str(item.source).lower())


def main() -> int:
    parser = argparse.ArgumentParser(description="Rename Readerlink rolling report week numbers to BookScan weeks.")
    parser.add_argument("--apply", action="store_true", help="Actually rename files. Omit for dry run.")
    args = parser.parse_args()

    plans = collect_plans()
    if not plans:
        print("No Readerlink weekly report filenames need week-number changes.")
        return 0

    renamed = 0
    skipped = 0
    for plan in plans:
        action = "RENAME" if args.apply else "DRY-RUN"
        print(f"{action}: {plan.source}")
        print(f"    -> {plan.target.name}")
        if not args.apply:
            continue
        if plan.target.exists():
            print("    skipped: target already exists")
            skipped += 1
            continue
        try:
            plan.source.rename(plan.target)
            renamed += 1
        except PermissionError:
            print("    skipped: file is locked or not writable")
            skipped += 1
        except OSError as exc:
            print(f"    skipped: {exc}")
            skipped += 1

    if args.apply:
        print(f"Done. Renamed {renamed} file(s); skipped {skipped} file(s).")
    else:
        print(f"Dry run complete. {len(plans)} file(s) would be renamed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
