import pandas as pd
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from asin_isbn_conversion import asin_isbn_conversion

def main():
    # Run the main data pipeline to get the cleaned DataFrame
    df = asin_isbn_conversion()

    # Get all float columns for later selection and reporting
    float_cols = df.select_dtypes(include='float64').columns.tolist()

    # Reorder columns: ASIN, ISBN, then all float columns
    df = df[['ASIN', 'ISBN'] + float_cols]

    # Rename columns for clarity in the final report
    rename = {
        'ISBN': 'External ID',
        'Ordered Units': 'Customer Orders',
        'Shipped Units': 'Units Shipped',
        'Sellable On Hand Units': 'Units at Amazon',
        'Open Purchase Order Quantity': 'Open PO qty'
    }
    df = df.rename(columns=rename)

    # --- Excel Save Dialog ---
    # Create a hidden Tkinter root window for dialogs
    root = tk.Tk()
    root.withdraw()
    # Ask user where to save the main Excel file
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Save Amazon SQL Upload Excel File"
    )
    if file_path:
        # Save the DataFrame to the chosen Excel file, starting at row 4
        with pd.ExcelWriter(file_path) as writer:
            df.to_excel(writer, index=False, startrow=3)
        print(f"✅ Saved Excel file: {file_path}")
    else:
        print("❌ Save cancelled.")

    # --- NO_ISBN Reporting ---
    # Count how many titles have NO_ISBN
    no_isbn_count = (df['External ID'] == "NO_ISBN").sum()
    # Select columns to show for NO_ISBN reporting
    cols_to_show = ['ASIN', 'External ID'] + float_cols
    df_no_isbn = df[df['External ID'] == "NO_ISBN"][cols_to_show]

    # Ask user if they want to view or save NO_ISBN titles
    msg = (
        f"{no_isbn_count} titles have NO_ISBN.\n\n"
        "Would you like to:\n"
        "1. View them on screen\n"
        "2. Save them as Excel\n\n"
        "Enter 1 or 2:"
    )
    user_choice = simpledialog.askstring("NO_ISBN Titles", msg)

    if user_choice == "1":
        # Print top 20 NO_ISBN titles to the console
        print(df_no_isbn.head(20))
    elif user_choice == "2":
        # Ask user where to save the NO_ISBN Excel file
        no_isbn_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save NO_ISBN Titles Excel File"
        )
        if no_isbn_file:
            # Save NO_ISBN titles to Excel
            df_no_isbn.to_excel(no_isbn_file, index=False)
            messagebox.showinfo("Saved", f"✅ Saved Excel file: {no_isbn_file}")
        else:
            print("❌ Save cancelled.")
    else:
        print("No action taken.")

    # --- Option to Update Dictionaries ---
    # Ask user if they want to update ASIN removal list or manual key
    update_dicts = simpledialog.askstring(
        "Update Dictionaries",
        "Would you like to update the ASIN removal list or ASIN-->ISBN manual key?\nType 'y' for yes or 'n' for no:"
    )
    if update_dicts and update_dicts.lower() == 'y':
        # Run the dictionary update script as a subprocess
        subprocess.run(["python", "asin_add_to_dictionaries.py"])
    else:
        print("Skipping dictionary update.")

if __name__ == "__main__":
    main()