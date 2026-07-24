import ast
from pathlib import Path
import sys


SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from shared.amazon_metadata import (  # noqa: E402
    load_asin_isbn_overrides,
    save_asin_isbn_override,
)

def add_to_removal_list():
    removal_file = SCRIPT_DIR / "asin_removal_list.py"
    # Load current list (supports both set and list formats)
    with open(removal_file, "r") as f:
        content = f.read()
    # Extract the right-hand side and safely evaluate
    rhs = content.split('=')[1].strip()
    try:
        current_list = set(ast.literal_eval(rhs))
    except Exception:
        # fallback for set literal (not supported by ast.literal_eval in Python <3.12)
        current_list = set(eval(rhs))
    print("Current ASINs to delete:", sorted(current_list))
    while True:
        asin = input("Enter ASIN to add to removal list (or press Enter to finish): ").strip()
        if not asin:
            break
        current_list.add(asin)
    # Write updated list as a Python list for compatibility
    with open(removal_file, "w") as f:
        f.write(f"asins_to_delete_list = {sorted(list(current_list))}\n")
    print("The ASIN removal list has been updated!")

def add_to_manual_key():
    current_dict = load_asin_isbn_overrides()
    print("Current ASIN/ISBN manual pairs:", current_dict)
    while True:
        asin = input("Enter an ASIN first (or press Enter to finish): ").strip()
        if not asin:
            break
        isbn = input(f"Now enter the associated ISBN for that ASIN {asin}: ").strip()
        save_asin_isbn_override(asin, isbn)
        current_dict[asin.upper()] = isbn
    print("Updated shared/amazon_asin_isbn_overrides.json!")

def main():
    print("Would you like to add any ASINs to a list of ASINs that should be removed?")
    response = input("Type 'y' for yes or 'n' for no: ").lower()
    if response == 'y':
        add_to_removal_list()
    elif response == 'n':
        print("Skipping ASIN removal list update.")

    print("Would you like to add any ASIN-->ISBN pairs that will be associated in the future?")
    response = input("Type 'y' for yes or 'n' for no: ").lower()
    if response == 'y':
        add_to_manual_key()
    elif response == 'n':
        print("Skipping ASIN-->ISBN manual key update.")

if __name__ == "__main__":
    main()
