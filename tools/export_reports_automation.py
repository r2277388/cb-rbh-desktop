from __future__ import annotations
import argparse, shutil, subprocess, sys, time
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path: sys.path.insert(0, str(ROOT))
from paths import process_paths

def excel_python():
    for candidate in (sys.executable, shutil.which("python"), r"C:\Python310\python.exe"):
        if candidate and subprocess.run([candidate, "-c", "import pythoncom"], capture_output=True).returncode == 0:
            return str(Path(candidate).resolve())
    raise RuntimeError("No Python interpreter with pywin32/pythoncom was found.")

def query_task():
    result = subprocess.run(["schtasks", "/Query", "/TN", process_paths.EXPORT_REPORT_TASK_NAME, "/FO", "LIST", "/V"], capture_output=True, text=True)
    if result.returncode: return None
    return {line.split(":", 1)[0].strip(): line.split(":", 1)[1].strip() for line in result.stdout.splitlines() if ":" in line}

def status():
    details = query_task()
    print("\nWeekly Export Reports Query & Pivot Refresh")
    print(f"  Schedule: {process_paths.EXPORT_REPORT_SCHEDULE_DESCRIPTION}")
    for path in process_paths.EXPORT_REPORT_WORKBOOKS: print(f"  Workbook: {path}")
    print(f"  Status: {'Scheduled' if details else 'Not currently scheduled'}")
    if details:
        for key in ("Next Run Time", "Last Run Time", "Last Result"): print(f"  {key}: {details.get(key, 'Unknown')}")
    return 0

def register():
    command = f'\"{excel_python()}\" \"{Path(__file__).resolve()}\" run'
    return subprocess.run(["schtasks", "/Create", "/TN", process_paths.EXPORT_REPORT_TASK_NAME, "/SC", "WEEKLY", "/D", process_paths.EXPORT_REPORT_SCHEDULE_DAY, "/ST", process_paths.EXPORT_REPORT_SCHEDULE_TIME, "/TR", command, "/RL", "LIMITED", "/IT", "/F"]).returncode

def disable():
    return subprocess.run(["schtasks", "/Change", "/TN", process_paths.EXPORT_REPORT_TASK_NAME, "/Disable"]).returncode

def wait(excel):
    started = time.monotonic()
    while excel.CalculationState != 0:
        try: excel.CalculateUntilAsyncQueriesDone()
        except Exception: pass
        if time.monotonic() - started > 1800: raise TimeoutError("Excel refresh exceeded 30 minutes.")
        time.sleep(1)

def run():
    import pythoncom, win32com.client
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible, excel.DisplayAlerts = False, False
    failures = []
    try:
        for path in process_paths.EXPORT_REPORT_WORKBOOKS:
            workbook = None
            try:
                fallback = path.with_name(f"{path.stem}_v1{path.suffix}")
                if fallback.exists():
                    print(f"Removing prior fallback: {fallback}", flush=True)
                    fallback.unlink()
                workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True)
                print(f"Refreshing queries: {path}", flush=True)
                workbook.RefreshAll()
                excel.CalculateUntilAsyncQueriesDone()
                wait(excel)
                pivots = 0
                for sheet in workbook.Worksheets:
                    tables = sheet.PivotTables()
                    for index in range(1, tables.Count + 1):
                        tables.Item(index).RefreshTable(); pivots += 1
                wait(excel)
                if workbook.ReadOnly:
                    print(f"Main file is open; saving: {fallback}", flush=True)
                    workbook.SaveAs(str(fallback), FileFormat=workbook.FileFormat)
                else: workbook.Save()
                print(f"Complete ({pivots} pivots): {path}", flush=True)
            except Exception as exc:
                failures.append(f"{path}: {exc}"); print(f"FAILED: {path}: {exc}", file=sys.stderr)
            finally:
                if workbook is not None: workbook.Close(SaveChanges=False)
    finally:
        excel.Quit(); pythoncom.CoUninitialize()
    return 1 if failures else 0

def menu():
    while True:
        status()
        print("\n    1. Run Now\n    2. Enable or Update Weekly Automation\n    3. Disable Weekly Automation\n    4. Return to Automation Processes\n")
        choice = input("Choose an option: ").strip().lower()
        if choice == "1": run()
        elif choice == "2": register()
        elif choice == "3": disable()
        elif choice in {"4", "b", "back", "return"}: return 0
        else: print("Invalid choice.")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("command", choices=("run", "status", "register", "disable", "menu"))
    command = parser.parse_args().command
    return {"run": run, "status": status, "register": register, "disable": disable, "menu": menu}[command]()

if __name__ == "__main__": raise SystemExit(main())
