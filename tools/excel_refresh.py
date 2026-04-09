from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

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


XL_DONE = 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Inspect or refresh Excel workbook queries/connections via Excel COM."
    )
    parser.add_argument("workbook", type=Path, help="Path to the workbook.")
    parser.add_argument(
        "--connection",
        help="Workbook connection name to refresh. Defaults to RefreshAll().",
    )
    parser.add_argument(
        "--table",
        help="Excel table (ListObject) name to refresh after the query refresh.",
    )
    parser.add_argument(
        "--inspect",
        action="store_true",
        help="List workbook connections and table names, then exit.",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Show Excel while the automation runs.",
    )
    parser.add_argument(
        "--read-only",
        action="store_true",
        help="Open the workbook read-only. Implies no save.",
    )
    parser.add_argument(
        "--save-as",
        type=Path,
        help="Optional output path. If omitted, saves back to the source workbook.",
    )
    parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=1800,
        help="Maximum time to wait for refresh/calculation to finish. Default: 1800.",
    )
    return parser.parse_args()


def log(message: str) -> None:
    print(message, flush=True)


def normalize_name(value) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def iter_connections(workbook):
    for i in range(1, workbook.Connections.Count + 1):
        yield workbook.Connections.Item(i)


def iter_tables(workbook):
    for worksheet in workbook.Worksheets:
        try:
            list_objects = worksheet.ListObjects
            count = list_objects.Count
        except Exception:
            continue

        for i in range(1, count + 1):
            yield worksheet, list_objects.Item(i)


def list_workbook_objects(workbook) -> None:
    log("Connections:")
    connection_count = 0
    for connection in iter_connections(workbook):
        connection_count += 1
        log(f"  - {connection.Name}")

    if not connection_count:
        log("  (none)")

    log("Tables:")
    table_count = 0
    for worksheet, table in iter_tables(workbook):
        table_count += 1
        log(f"  - {worksheet.Name}!{table.Name}")

    if not table_count:
        log("  (none)")


def wait_for_excel(workbook, app, timeout_seconds: int) -> None:
    started = time.time()

    while True:
        calculating = app.CalculationState != XL_DONE

        refreshing = False
        for connection in iter_connections(workbook):
            try:
                ole_db = getattr(connection, "OLEDBConnection", None)
                if ole_db is not None and ole_db.Refreshing:
                    refreshing = True
                    break
            except Exception:
                pass

            try:
                odbc = getattr(connection, "ODBCConnection", None)
                if odbc is not None and odbc.Refreshing:
                    refreshing = True
                    break
            except Exception:
                pass

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


def find_connection(workbook, connection_name: str):
    target = normalize_name(connection_name)
    matches = [
        connection
        for connection in iter_connections(workbook)
        if normalize_name(connection.Name) == target
    ]

    if not matches:
        available = ", ".join(connection.Name for connection in iter_connections(workbook))
        raise ValueError(
            f"Connection '{connection_name}' not found. "
            f"Available connections: {available or '(none)'}"
        )

    return matches[0]


def find_table(workbook, table_name: str):
    target = normalize_name(table_name)
    matches = [
        (worksheet, table)
        for worksheet, table in iter_tables(workbook)
        if normalize_name(table.Name) == target
    ]

    if not matches:
        available = ", ".join(
            f"{worksheet.Name}!{table.Name}" for worksheet, table in iter_tables(workbook)
        )
        raise ValueError(
            f"Table '{table_name}' not found. Available tables: {available or '(none)'}"
        )

    return matches[0]


def refresh_target(workbook, app, connection_name: str | None, table_name: str | None, timeout_seconds: int) -> None:
    if connection_name:
        connection = find_connection(workbook, connection_name)
        log(f"Refreshing connection: {connection.Name}")
        connection.Refresh()
    else:
        log("Refreshing workbook data with RefreshAll()")
        workbook.RefreshAll()

    if table_name:
        worksheet, table = find_table(workbook, table_name)
        log(f"Refreshing table: {worksheet.Name}!{table.Name}")
        try:
            table.QueryTable.Refresh(False)
        except Exception:
            # Not every ListObject exposes QueryTable; Power Query-backed tables
            # usually update through the workbook connection refresh above.
            log("Table does not expose QueryTable directly; relying on connection refresh.")

    app.CalculateUntilAsyncQueriesDone()
    wait_for_excel(workbook, app, timeout_seconds)


def main() -> int:
    args = parse_args()

    workbook_path = args.workbook
    if not workbook_path.exists():
        print(f"Workbook not found: {workbook_path}", file=sys.stderr)
        return 1

    save_changes = not args.read_only and not args.inspect
    save_as_path = args.save_as

    pythoncom.CoInitialize()
    excel = None
    workbook = None

    try:
        log("Starting Excel...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = args.visible
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        try:
            excel.Calculation = -4105
        except Exception:
            pass

        log(f"Opening workbook: {workbook_path}")
        workbook = excel.Workbooks.Open(
            str(workbook_path),
            UpdateLinks=0,
            ReadOnly=args.read_only,
        )

        if args.inspect:
            list_workbook_objects(workbook)
            return 0

        refresh_target(
            workbook=workbook,
            app=excel,
            connection_name=args.connection,
            table_name=args.table,
            timeout_seconds=args.timeout_seconds,
        )

        if save_as_path:
            log(f"Saving refreshed workbook as: {save_as_path}")
            workbook.SaveAs(str(save_as_path))
        elif save_changes:
            log("Saving refreshed workbook...")
            workbook.Save()
        else:
            log("Skipping save.")

        log("Refresh complete.")
        return 0
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass

        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
