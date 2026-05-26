from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from paths import process_paths
from shared.bookscan_calendar import bookscan_week


WEEKLY_AMAZON_PATTERN = re.compile(
    r"^Week (?P<week>\d{1,2})-(?P<year>\d{4}) "
    r"Rolling Amazon \((?P<month>\d{2})_(?P<day>\d{2})_(?P<date_year>\d{4})\) "
    r"- (?P<report_type>Customer Orders|Units Shipped)\.xlsx$",
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


def amazon_roots() -> list[Path]:
    roots = [process_paths.AMAZON_ROLLING_OUTPUT_FOLDER]
    roots.extend(process_paths.AMAZON_ROLLING_DP_FOLDERS.values())
    return roots


def planned_rename(path: Path) -> RenamePlan | None:
    match = WEEKLY_AMAZON_PATTERN.match(path.name)
    if match is None:
        return None

    week_end = f"{match.group('date_year')}-{match.group('month')}-{match.group('day')}"
    bookscan = bookscan_week(week_end)
    old_week = match.group("week")
    old_year = match.group("year")
    new_week = f"{bookscan.week:02d}"
    new_year = str(bookscan.year)

    if old_week == new_week and old_year == new_year:
        return None

    target_name = WEEKLY_AMAZON_PATTERN.sub(
        rf"Week {new_week}-{new_year} Rolling Amazon (\g<month>_\g<day>_\g<date_year>) - \g<report_type>.xlsx",
        path.name,
    )
    return RenamePlan(
        source=path,
        target=path.with_name(target_name),
        old_week=old_week,
        new_week=new_week,
        old_year=old_year,
        new_year=new_year,
    )


def collect_plans() -> tuple[list[RenamePlan], list[Path]]:
    plans: list[RenamePlan] = []
    missing_roots: list[Path] = []
    for root in amazon_roots():
        if not root.exists():
            missing_roots.append(root)
            continue
        for path in root.rglob("*.xlsx"):
            plan = planned_rename(path)
            if plan is not None:
                plans.append(plan)
    return sorted(plans, key=lambda item: str(item.source).lower()), missing_roots


def main() -> int:
    parser = argparse.ArgumentParser(description="Rename Amazon rolling report week numbers to BookScan weeks.")
    parser.add_argument("--apply", action="store_true", help="Actually rename files. Omit for dry run.")
    args = parser.parse_args()

    plans, missing_roots = collect_plans()
    for root in missing_roots:
        print(f"Missing folder, skipped: {root}")

    if not plans:
        print("No Amazon weekly report filenames need week-number changes.")
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
            print(f"    skipped: target already exists")
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
