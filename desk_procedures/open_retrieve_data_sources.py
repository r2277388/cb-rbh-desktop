from __future__ import annotations

import os
import webbrowser
from pathlib import Path


PROCEDURE_PATH = Path(__file__).resolve().parent / "procedures" / "retrieve_rolling_report_data_sources.html"


def main() -> None:
    if not PROCEDURE_PATH.exists():
        raise FileNotFoundError(f"Data source instructions file not found: {PROCEDURE_PATH}")

    uri = PROCEDURE_PATH.as_uri()
    try:
        if webbrowser.open_new_tab(uri):
            print(f"Opened data source instructions: {PROCEDURE_PATH}")
            return
    except webbrowser.Error:
        pass

    os.startfile(str(PROCEDURE_PATH))
    print(f"Opened data source instructions: {PROCEDURE_PATH}")


if __name__ == "__main__":
    main()
