import ast

def add_to_removal_list():
    removal_file = "asin_removal_list.py"
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
    manual_file = "asin_manual_key.py"
    # Load current dictionary
    with open(manual_file, "r") as f:
        content = f.read()
    rhs = content.split('=')[1].strip()
    current_dict = ast.literal_eval(rhs)
    print("Current ASIN/ISBN manual pairs:", current_dict)
    while True:
        asin = input("Enter an ASIN first (or press Enter to finish): ").strip()
        if not asin:
            break
        isbn = input(f"Now enter the associated ISBN for that ASIN {asin}: ").strip()
        current_dict[asin] = isbn
    # Write updated dictionary
    with open(manual_file, "w") as f:
        f.write(f"asin_isbn_manual_key = {repr(current_dict)}\n")
    print("Updated asin_manual_key.py!")

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