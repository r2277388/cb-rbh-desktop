from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

try:
    from paths import process_paths
except ModuleNotFoundError:
    import sys

    sys.path.append(str(Path(__file__).resolve().parent.parent))
    from paths import process_paths


POWER_BI_EXTENSIONS = {".pbix", ".pbit", ".pbip"}


@dataclass(frozen=True)
class ReportFolder:
    section: str
    label: str
    path: Path


REPORT_FOLDERS = [
    ReportFolder("Visual Dashboards", "Visual Dashboards", process_paths.POWER_BI_REPORTS_FOLDER),
    *[
        ReportFolder("Barrett Reports", label, folder)
        for label, folder in process_paths.POWER_BI_BARRETT_REPORT_FOLDERS.items()
    ],
]


@dataclass(frozen=True)
class PathTimestamp:
    display: str

    @classmethod
    def from_epoch(cls, epoch_seconds: float) -> "PathTimestamp":
        value = datetime.fromtimestamp(epoch_seconds)
        hour = value.hour % 12 or 12
        am_pm = "AM" if value.hour < 12 else "PM"
        return cls(
            f"{value.month}/{value.day}/{value.year} "
            f"{hour}:{value.minute:02d}:{value.second:02d} {am_pm}"
        )


@dataclass(frozen=True)
class PowerBIReport:
    path: Path
    last_modified: str

    def display_path(self, folder: Path) -> str:
        try:
            return str(self.path.relative_to(folder))
        except ValueError:
            return str(self.path)


def find_power_bi_reports(folder: Path) -> list[PowerBIReport]:
    try:
        if not folder.exists():
            raise FileNotFoundError(f"Power BI reports folder not found: {folder}")
    except PermissionError as exc:
        raise PermissionError(f"Power BI reports folder access denied: {folder}") from exc

    reports: list[PowerBIReport] = []
    for file_path in folder.rglob("*"):
        if not file_path.is_file():
            continue
        if file_path.suffix.lower() not in POWER_BI_EXTENSIONS:
            continue
        modified = file_path.stat().st_mtime
        reports.append(
            PowerBIReport(
                path=file_path,
                last_modified=PathTimestamp.from_epoch(modified).display,
            )
        )

    return sorted(reports, key=lambda report: report.display_path(folder).lower())


def print_reports_table(reports: list[PowerBIReport], folder: Path) -> None:
    if not reports:
        print("No Power BI files were found.")
        return

    file_header = "File"
    modified_header = "Last modified"
    file_width = max(len(file_header), *(len(report.display_path(folder)) for report in reports))
    modified_width = max(len(modified_header), *(len(report.last_modified) for report in reports))

    print(f"{file_header:<{file_width}}  {modified_header:<{modified_width}}")
    print(f"{'-' * file_width}  {'-' * modified_width}")
    for report in reports:
        print(f"{report.display_path(folder):<{file_width}}  {report.last_modified:<{modified_width}}")


def print_report_folder(report_folder: ReportFolder) -> None:
    print(f"\n{report_folder.label}")
    print(f"Folder: {report_folder.path}")
    print()
    try:
        reports = find_power_bi_reports(report_folder.path)
    except (FileNotFoundError, PermissionError) as exc:
        print(exc)
        return
    print_reports_table(reports, report_folder.path)


def main() -> None:
    print("\nPower BI Reports")
    current_section = None
    for report_folder in REPORT_FOLDERS:
        if report_folder.section != current_section:
            current_section = report_folder.section
            print(f"\n{current_section}")
            print("=" * len(current_section))
        print_report_folder(report_folder)


if __name__ == "__main__":
    main()
