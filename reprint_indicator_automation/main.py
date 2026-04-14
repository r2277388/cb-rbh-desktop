from __future__ import annotations

import argparse
import gc
import sys
import time
from datetime import datetime
from pathlib import Path
from tkinter import Tk, messagebox

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from shared import send_outlook_mail

try:
    import pythoncom
    import win32com.client
except ModuleNotFoundError as exc:
    missing_module = exc.name or "pywin32"
    print(
        "Error: Excel automation requires pywin32. "
        f"Missing module: {missing_module}. "
        f"Current Python: {sys.executable}",
        file=sys.stderr,
    )
    raise SystemExit(1) from exc


DEFAULT_TEMPLATE = Path(
    r"G:\OPS\Reprint Indicators\Templates\Reprint indicator TEMPLATE.xlsx"
)
DEFAULT_OUTPUT_DIR = Path(r"G:\OPS\Reprint Indicators\2026")
DEFAULT_FILENAME_PREFIX = "reprint indicator"

XL_LINK_TYPE_EXCEL = 1
XL_LINK_TYPE_OLE = 2
XL_OPEN_XML_WORKBOOK = 51
XL_UP = -4162
DETAIL_START_ROW = 3
METADATA_START_ROW = 2
BL_DETAIL_LAST_COL = "BO"
FL_DETAIL_LAST_COL = "BY"
EXPORT_FIRST_SHEET = "RPG_Risk_Analyzer"
EXPORT_LAST_SHEET = "Explanation"
EMAIL_TO = ["Mary_OHara@chroniclebooks.com"]
EMAIL_CC = [
    "john_carlson@chroniclebooks.com",
    "Kate_BreitingSchmitz@chroniclebooks.com",
]
EMAIL_SUBJECT_PREFIX = "Reprint Indicator Updated thru"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Refresh the Reprint Indicator workbook in Excel, break external links, "
            "and save a dated output copy."
        )
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=DEFAULT_TEMPLATE,
        help=f"Template workbook path. Default: {DEFAULT_TEMPLATE}",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help=f"Output directory. Default: {DEFAULT_OUTPUT_DIR}",
    )
    parser.add_argument(
        "--date",
        dest="date_text",
        help="Override output date using YYYY_MM_DD or YYYY-MM-DD.",
    )
    parser.add_argument(
        "--prefix",
        default=DEFAULT_FILENAME_PREFIX,
        help=f'Output filename prefix. Default: "{DEFAULT_FILENAME_PREFIX}"',
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Show Excel while the automation runs.",
    )
    parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=1800,
        help="Maximum time to wait for refresh/calculation to finish. Default: 1800.",
    )
    parser.add_argument(
        "--save-template",
        action="store_true",
        help="Save refreshed/rebuilt changes back to the template workbook before exporting.",
    )
    return parser.parse_args()


def normalize_output_date(date_text: str | None) -> str:
    if not date_text:
        return datetime.now().strftime("%Y_%m_%d")

    for fmt in ("%Y_%m_%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(date_text, fmt).strftime("%Y_%m_%d")
        except ValueError:
            continue

    raise ValueError("Date must be YYYY_MM_DD or YYYY-MM-DD.")


def prompt_refresh_required() -> bool:
    while True:
        response = input("Refresh workbook data first? [y/n]: ").strip().lower()
        if response in {"y", "yes"}:
            return True
        if response in {"n", "no"}:
            return False
        print("Please enter y or n.")


def prompt_save_template() -> bool:
    while True:
        response = (
            input(
                "Save refreshed/rebuilt changes back to the template workbook too? [y/n]: "
            )
            .strip()
            .lower()
        )
        if response in {"y", "yes"}:
            return True
        if response in {"n", "no"}:
            return False
        print("Please enter y or n.")


def log_step(message: str) -> None:
    print(message, flush=True)


def log_warning(message: str) -> None:
    print(f"Warning: {message}", flush=True)


def show_popup(title: str, message: str) -> None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        messagebox.showinfo(title, message, parent=root)
    finally:
        root.destroy()


def format_display_date(output_date: str) -> str:
    return datetime.strptime(output_date, "%Y_%m_%d").strftime("%m/%d/%Y")


def open_notification_draft(
    *,
    output_path: Path,
    output_date: str,
    bl_count: int,
    fl_count: int,
) -> None:
    subject = f"{EMAIL_SUBJECT_PREFIX} {format_display_date(output_date)}"
    body = "\n".join(
        [
            "Enjoy this week's updated version:",
            "",
            f'"{output_path}"',
            "",
            "Summary:",
            f"BL_Detail rows: {bl_count}",
            f"FL_Detail rows: {fl_count}",
        ]
    )
    send_outlook_mail(
        to=EMAIL_TO,
        cc=EMAIL_CC,
        subject=subject,
        body=body,
        display_before_send=True,
    )


def unique_output_path(output_dir: Path, prefix: str, output_date: str) -> Path:
    base = output_dir / f"{prefix}_{output_date}.xlsx"
    if not base.exists():
        return base

    counter = 1
    while True:
        candidate = output_dir / f"{prefix}_{output_date}_v{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def normalize_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def last_used_row(worksheet, column_index: int) -> int:
    return worksheet.Cells(worksheet.Rows.Count, column_index).End(XL_UP).Row


def find_header_column(worksheet, header_name: str, header_row: int = 1) -> int:
    last_col = worksheet.Cells(header_row, worksheet.Columns.Count).End(-4159).Column
    for col in range(1, last_col + 1):
        if normalize_text(worksheet.Cells(header_row, col).Value) == header_name:
            return col
    raise ValueError(f'Header "{header_name}" not found on sheet "{worksheet.Name}".')


def get_export_sheet_names(workbook) -> list[str]:
    names = [worksheet.Name for worksheet in workbook.Worksheets]
    try:
        start_index = names.index(EXPORT_FIRST_SHEET)
        end_index = names.index(EXPORT_LAST_SHEET)
    except ValueError as exc:
        raise ValueError(
            f"Could not find export sheet range {EXPORT_FIRST_SHEET}..{EXPORT_LAST_SHEET}."
        ) from exc

    if start_index > end_index:
        raise ValueError(
            f"Export sheet order is invalid: {EXPORT_FIRST_SHEET} comes after {EXPORT_LAST_SHEET}."
        )

    return names[start_index : end_index + 1]


def metadata_rows(workbook) -> tuple[list[str], list[tuple[str, str]]]:
    metadata_ws = workbook.Worksheets("MetaData")
    release_group_col = find_header_column(metadata_ws, "Release_Group")
    isbn_col = find_header_column(metadata_ws, "ISBN")
    last_row = max(
        last_used_row(metadata_ws, release_group_col),
        last_used_row(metadata_ws, isbn_col),
    )

    if last_row < METADATA_START_ROW:
        return [], []

    values = metadata_ws.Range(
        metadata_ws.Cells(METADATA_START_ROW, release_group_col),
        metadata_ws.Cells(last_row, isbn_col),
    ).Value
    if not isinstance(values, tuple):
        values = (values,)

    backlist_isbns: list[str] = []
    frontlist_rows: list[tuple[str, str]] = []

    for row in values:
        if not isinstance(row, tuple):
            row = (row,)

        release_group = normalize_text(row[0] if len(row) >= 1 else "")
        isbn = normalize_text(
            row[isbn_col - release_group_col]
            if len(row) > (isbn_col - release_group_col)
            else ""
        )
        if not isbn:
            continue

        if "backlist" in release_group.lower():
            backlist_isbns.append(isbn)
        else:
            frontlist_rows.append((release_group, isbn))

    return backlist_isbns, frontlist_rows


def clear_detail_rows(worksheet, start_row: int, end_row: int, last_col: str) -> None:
    if end_row < start_row:
        return
    worksheet.Range(f"A{start_row}:{last_col}{end_row}").ClearContents()


def autofill_template_row(
    worksheet, start_row: int, target_end_row: int, last_col: str
) -> None:
    if target_end_row < start_row:
        return
    source_range = worksheet.Range(f"A{start_row}:{last_col}{start_row}")
    destination_range = worksheet.Range(f"A{start_row}:{last_col}{target_end_row}")
    source_range.AutoFill(Destination=destination_range)


def write_column_values(
    worksheet, column_letter: str, start_row: int, values: list[str]
) -> None:
    if not values:
        return
    range_ref = (
        f"{column_letter}{start_row}:{column_letter}{start_row + len(values) - 1}"
    )
    worksheet.Range(range_ref).Value = tuple((value,) for value in values)


def write_two_column_values(
    worksheet, start_row: int, rows: list[tuple[str, str]]
) -> None:
    if not rows:
        return
    range_ref = f"A{start_row}:B{start_row + len(rows) - 1}"
    worksheet.Range(range_ref).Value = tuple(rows)


def rebuild_bl_detail(workbook, backlist_isbns: list[str]) -> int:
    worksheet = workbook.Worksheets("BL_Detail")
    current_last_row = max(last_used_row(worksheet, 1), DETAIL_START_ROW)
    clear_detail_rows(
        worksheet, DETAIL_START_ROW + 1, current_last_row, BL_DETAIL_LAST_COL
    )

    if not backlist_isbns:
        clear_detail_rows(
            worksheet, DETAIL_START_ROW, DETAIL_START_ROW, BL_DETAIL_LAST_COL
        )
        return 0

    target_end_row = DETAIL_START_ROW + len(backlist_isbns) - 1
    autofill_template_row(
        worksheet, DETAIL_START_ROW, target_end_row, BL_DETAIL_LAST_COL
    )
    write_column_values(worksheet, "A", DETAIL_START_ROW, backlist_isbns)
    if current_last_row > target_end_row:
        clear_detail_rows(
            worksheet, target_end_row + 1, current_last_row, BL_DETAIL_LAST_COL
        )
    return len(backlist_isbns)


def rebuild_fl_detail(workbook, frontlist_rows: list[tuple[str, str]]) -> int:
    worksheet = workbook.Worksheets("FL_Detail")
    current_last_row = max(last_used_row(worksheet, 1), DETAIL_START_ROW)
    clear_detail_rows(
        worksheet, DETAIL_START_ROW + 1, current_last_row, FL_DETAIL_LAST_COL
    )

    if not frontlist_rows:
        clear_detail_rows(
            worksheet, DETAIL_START_ROW, DETAIL_START_ROW, FL_DETAIL_LAST_COL
        )
        return 0

    target_end_row = DETAIL_START_ROW + len(frontlist_rows) - 1
    autofill_template_row(
        worksheet, DETAIL_START_ROW, target_end_row, FL_DETAIL_LAST_COL
    )
    write_two_column_values(worksheet, DETAIL_START_ROW, frontlist_rows)
    if current_last_row > target_end_row:
        clear_detail_rows(
            worksheet, target_end_row + 1, current_last_row, FL_DETAIL_LAST_COL
        )
    return len(frontlist_rows)


def wait_for_excel(workbook, app, timeout_seconds: int) -> None:
    started = time.time()
    while True:
        pythoncom.PumpWaitingMessages()

        try:
            calculating = app.CalculationState != 0
        except Exception:
            calculating = False

        try:
            refreshing = workbook.Refreshing
        except Exception:
            refreshing = False

        try:
            async_queries_done = app.CalculateUntilAsyncQueriesDone() is None
        except Exception:
            async_queries_done = True

        if not calculating and not refreshing and async_queries_done:
            return

        if time.time() - started > timeout_seconds:
            raise TimeoutError(
                f"Excel refresh/calculation did not finish within {timeout_seconds} seconds."
            )

        time.sleep(1)


def refresh_workbook(workbook, app, timeout_seconds: int) -> None:
    workbook.RefreshAll()
    app.CalculateUntilAsyncQueriesDone()
    wait_for_excel(workbook, app, timeout_seconds)


def break_external_links(workbook) -> int:
    broken_count = 0
    for link_type in (XL_LINK_TYPE_EXCEL, XL_LINK_TYPE_OLE):
        try:
            links = workbook.LinkSources(link_type)
        except Exception:
            links = None

        if not links:
            continue

        if isinstance(links, str):
            links = [links]

        for link in links:
            workbook.BreakLink(Name=link, Type=link_type)
            broken_count += 1

    return broken_count


def copy_export_sheets_to_new_workbook(source_workbook, sheet_names: list[str]):
    try:
        source_workbook.Worksheets(tuple(sheet_names)).Copy()
    except Exception:
        source_workbook.Worksheets(sheet_names[0]).Copy()
        output_workbook = source_workbook.Application.ActiveWorkbook
        for sheet_name in sheet_names[1:]:
            source_workbook.Worksheets(sheet_name).Copy(
                After=output_workbook.Worksheets(output_workbook.Worksheets.Count)
            )
        return output_workbook

    return source_workbook.Application.ActiveWorkbook


def save_reprint_indicator(
    template_path: Path,
    output_dir: Path,
    output_path: Path,
    visible: bool,
    timeout_seconds: int,
    refresh_first: bool,
    save_template: bool,
) -> tuple[Path, int, int, int, list[str], int]:
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    output_dir.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    excel = None
    template_wb = None
    output_wb = None

    try:
        log_step("Starting Excel...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = visible
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        try:
            excel.Calculation = -4105
        except Exception:
            pass

        log_step(f"Opening template: {template_path}")
        template_wb = excel.Workbooks.Open(
            str(template_path),
            UpdateLinks=0,
            ReadOnly=False,
        )
        bl_count = 0
        fl_count = 0
        if refresh_first:
            log_step("Refreshing workbook data...")
            refresh_workbook(template_wb, excel, timeout_seconds)
            log_step("Reading MetaData rows...")
            backlist_isbns, frontlist_rows = metadata_rows(template_wb)
            log_step("Rebuilding BL_Detail...")
            bl_count = rebuild_bl_detail(template_wb, backlist_isbns)
            log_step("Rebuilding FL_Detail...")
            fl_count = rebuild_fl_detail(template_wb, frontlist_rows)
            log_step("Waiting for Excel calculations to finish...")
            wait_for_excel(template_wb, excel, timeout_seconds)
            if save_template:
                log_step("Saving refreshed template workbook...")
                template_wb.Save()

        log_step("Resolving export sheets...")
        export_sheet_names = get_export_sheet_names(template_wb)
        log_step("Copying export sheets into a new workbook...")
        output_wb = copy_export_sheets_to_new_workbook(template_wb, export_sheet_names)
        log_step("Breaking external links in new workbook...")
        broken_count = break_external_links(output_wb)
        log_step(f"Saving final workbook: {output_path}")
        output_wb.SaveAs(str(output_path), FileFormat=XL_OPEN_XML_WORKBOOK)
        log_step("Closing saved workbook...")
        output_wb.Close(SaveChanges=False)
        output_wb = None

        return output_path, broken_count, bl_count, fl_count, export_sheet_names, 0
    finally:
        if output_wb is not None:
            try:
                output_wb.Close(SaveChanges=False)
            except Exception:
                pass

        if template_wb is not None:
            try:
                template_wb.Close(SaveChanges=False)
            except Exception:
                pass

        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

        output_wb = None
        template_wb = None
        excel = None
        gc.collect()

        pythoncom.CoUninitialize()


def main() -> int:
    args = parse_args()

    try:
        output_date = normalize_output_date(args.date_text)
        output_path = unique_output_path(args.output_dir, args.prefix, output_date)
        refresh_first = prompt_refresh_required()
        save_template = args.save_template
        if refresh_first and not save_template:
            save_template = prompt_save_template()
        (
            saved_path,
            broken_count,
            bl_count,
            fl_count,
            exported_sheets,
            _removed_connection_count,
        ) = save_reprint_indicator(
            template_path=args.template,
            output_dir=args.output_dir,
            output_path=output_path,
            visible=args.visible,
            timeout_seconds=args.timeout_seconds,
            refresh_first=refresh_first,
            save_template=save_template,
        )
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Saved: {saved_path}")
    print(f"BL_Detail rows: {bl_count}")
    print(f"FL_Detail rows: {fl_count}")
    print(f"Exported sheets: {', '.join(exported_sheets)}")
    print(f"Broken links: {broken_count}")
    try:
        open_notification_draft(
            output_path=saved_path,
            output_date=output_date,
            bl_count=bl_count,
            fl_count=fl_count,
        )
        print("Opened Outlook draft notification.")
        show_popup(
            "Reprint Indicator Complete",
            (
                "The Reprint Indicator process completed successfully.\n\n"
                "An Outlook draft email has been created for review."
            ),
        )
    except Exception as exc:
        log_warning(f"Could not open Outlook draft notification: {exc}")
        show_popup(
            "Reprint Indicator Complete",
            (
                "The Reprint Indicator process completed successfully.\n\n"
                "The Outlook draft email could not be created."
            ),
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
