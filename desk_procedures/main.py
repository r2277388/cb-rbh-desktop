import os
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent

PROCEDURES = {
    "1": {
        "title": "Update Freight Costs",
        "type": "html",
        "path": BASE_DIR / "procedures" / "update_freight_costs.html",
    },
    "2": {
        "title": "Update Barnes & Noble Rolling Reports",
        "type": "html",
        "path": BASE_DIR / "procedures" / "update_bn_rolling_reports.html",
    },
    "3": {
        "title": "Launcher Process Runbook",
        "type": "html",
        "path": BASE_DIR / "procedures" / "launcher_process_runbook.html",
    },
    "4": {
        "title": "Update Consolidated Inventory Manager (ConInv)",
        "type": "html",
        "path": BASE_DIR / "procedures" / "update_invobs_consolidated_inventory.html",
    },
    "5": {
        "title": "Update FLTracking Supercharged",
        "type": "html",
        "path": BASE_DIR / "procedures" / "update_fltracking_supercharged.html",
    },
    "6": {
        "title": "How To Add Procedures",
        "type": "text",
        "steps": [
            "Open desk_procedures/main.py.",
            "Add a new item to the PROCEDURES dictionary.",
            "For an HTML procedure, set type to 'html' and point path to the file.",
            "For a text procedure, add a list of step strings.",
        ],
    },
}


def display_menu():
    print("\nDesk Procedures")
    print()
    for key, procedure in PROCEDURES.items():
        print(f"    {int(key):02d}. {procedure['title']}")
    print("    99. Back to main menu")


def display_text_procedure(procedure: dict):
    print()
    print(procedure["title"])
    print()
    for index, step in enumerate(procedure["steps"], start=1):
        print(f"    {index}. {step}")


def open_html_procedure(procedure: dict):
    html_path = Path(procedure["path"]).resolve()
    if not html_path.exists():
        print(f"Procedure file not found: {html_path}")
        return

    try:
        os.startfile(str(html_path))
        print(f"Opened procedure: {html_path}")
    except OSError as exc:
        print(f"Unable to open procedure in browser: {exc}")
        print(f"Procedure file: {html_path}")


def open_procedure(choice: str):
    procedure = PROCEDURES.get(choice)
    if not procedure:
        print("Invalid choice. Please select a valid option.")
        return

    if procedure["type"] == "html":
        open_html_procedure(procedure)
        return

    display_text_procedure(procedure)
    print()
    input("Press Enter to return to the procedures menu...")


def main():
    while True:
        display_menu()
        print()
        choice = input("Choose a procedure: ").strip().lower()

        if choice.isdigit():
            choice = str(int(choice))

        if choice in {"99", "back", "b", "exit", "quit", "q"}:
            return

        open_procedure(choice)


if __name__ == "__main__":
    main()
