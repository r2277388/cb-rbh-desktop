import argparse
import sys
from pathlib import Path

import pandas as pd
from tkinter import Tk
from tkinter.filedialog import asksaveasfilename
from dict_cdu import create_cdu_dict
from dict_unit_cost2 import df_to_nested_dict
from load_consolidated_inventory import (
    consolidate_inventory_from_rows,
    filter_invobs_inventory_rows,
    load_period_consolidated_inventory,
)

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from paths import process_paths

VALUE_ORGS = ("cbc", "hbg", "cbp")
VALUE_TOLERANCE = 0.004


def process_inventory(df_inventory, cdu_dict, dict_uc):
    # Initialize a list to store total units and values for each component ISBN
    result = []

    # Iterate over each row in df_inventory
    for _, row in df_inventory.iterrows():
        isbn = row['ISBN']
        units_cbc = row['units_cbc']
        units_hbg = row['units_hbg']
        units_cbp = row['units_cbp']

        # Check if the ISBN is a CDU (i.e., it has components in cdu_dict)
        if isbn in cdu_dict:
            # This is a CDU, so we need to break it down into its components
            for component_isbn, component_qty in cdu_dict[isbn].items():
                # Calculate the total units for each component
                total_units_cbc = units_cbc * component_qty
                total_units_hbg = units_hbg * component_qty
                total_units_cbp = units_cbp * component_qty

                # Get the unit costs from dict_uc
                if component_isbn in dict_uc:
                    uc_cbc = dict_uc[component_isbn].get('uc_cbc', 0)
                    uc_hbg = dict_uc[component_isbn].get('uc_hbg', 0)
                    uc_cbp = dict_uc[component_isbn].get('uc_cbp', 0)

                    # Calculate the value for each component
                    value_cbc = total_units_cbc * uc_cbc
                    value_hbg = total_units_hbg * uc_hbg
                    value_cbp = total_units_cbp * uc_cbp

                    # Append the results for this component as a dictionary
                    result.append({
                        'ISBN': component_isbn,
                        'total_units_cbc': total_units_cbc,
                        'total_units_hbg': total_units_hbg,
                        'total_units_cbp': total_units_cbp,
                        'total_value_cbc': value_cbc,
                        'total_value_hbg': value_hbg,
                        'total_value_cbp': value_cbp
                    })
        else:
            # The ISBN is not a CDU, process it normally
            # Get the unit costs from dict_uc
            if isbn in dict_uc:
                uc_cbc = dict_uc[isbn].get('uc_cbc', 0)
                uc_hbg = dict_uc[isbn].get('uc_hbg', 0)
                uc_cbp = dict_uc[isbn].get('uc_cbp', 0)

                # Calculate the value for each bucket
                value_cbc = units_cbc * uc_cbc
                value_hbg = units_hbg * uc_hbg
                value_cbp = units_cbp * uc_cbp

                # Append the results for this ISBN as a dictionary
                result.append({
                    'ISBN': isbn,
                    'total_units_cbc': units_cbc,
                    'total_units_hbg': units_hbg,
                    'total_units_cbp': units_cbp,
                    'total_value_cbc': value_cbc,
                    'total_value_hbg': value_hbg,
                    'total_value_cbp': value_cbp
                })

    # Convert the result into a pandas DataFrame
    result_df = pd.DataFrame(result)
    result_df = result_df[['ISBN', 'total_value_cbc', 'total_units_cbc', \
                           'total_value_hbg', 'total_units_hbg', 'total_value_cbp', 'total_units_cbp']]
    result_df['total_value'] = result_df['total_value_cbc'] + result_df['total_value_hbg'] + result_df['total_value_cbp']
    result_df['total_units'] = result_df['total_units_cbc'] + result_df['total_units_hbg'] + result_df['total_units_cbp']
    return result_df


def build_value_variance(source_rows, aggregated_inventory):
    original_pivot = (
        source_rows.pivot_table(
            index="ISBN",
            columns="ORG",
            values="Value",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )
    original_pivot.columns = ["ISBN" if col == "ISBN" else str(col).lower() for col in original_pivot.columns]

    original = pd.DataFrame({"ISBN": original_pivot["ISBN"]})
    for org in VALUE_ORGS:
        original[f"original_value_{org}"] = original_pivot[org] if org in original_pivot.columns else 0

    final = pd.DataFrame({"ISBN": aggregated_inventory["ISBN"]})
    for org in VALUE_ORGS:
        source_col = f"total_value_{org}"
        final[f"final_value_{org}"] = aggregated_inventory[source_col] if source_col in aggregated_inventory.columns else 0

    variance = original.merge(final, on="ISBN", how="outer").fillna(0)
    for org in VALUE_ORGS:
        variance[f"variance_value_{org}"] = variance[f"final_value_{org}"] - variance[f"original_value_{org}"]

    original_cols = [f"original_value_{org}" for org in VALUE_ORGS]
    final_cols = [f"final_value_{org}" for org in VALUE_ORGS]
    variance_cols = [f"variance_value_{org}" for org in VALUE_ORGS]
    variance["original_total_value"] = variance[original_cols].sum(axis=1)
    variance["final_total_value"] = variance[final_cols].sum(axis=1)
    variance["variance_total_value"] = variance[variance_cols].sum(axis=1)

    def variance_status(row):
        original_total = row["original_total_value"]
        final_total = row["final_total_value"]
        if abs(original_total) <= VALUE_TOLERANCE and abs(final_total) > VALUE_TOLERANCE:
            return "Added in final/CDU component"
        if abs(original_total) > VALUE_TOLERANCE and abs(final_total) <= VALUE_TOLERANCE:
            return "Removed or replaced by cleanup"
        if abs(row["variance_total_value"]) > VALUE_TOLERANCE:
            return "Changed value"
        return "No total variance"

    variance["variance_status"] = variance.apply(variance_status, axis=1)
    variance["abs_variance_total_value"] = variance["variance_total_value"].abs()

    detail_cols = [
        "ISBN",
        "variance_status",
        "original_total_value",
        "final_total_value",
        "variance_total_value",
        "original_value_cbc",
        "final_value_cbc",
        "variance_value_cbc",
        "original_value_hbg",
        "final_value_hbg",
        "variance_value_hbg",
        "original_value_cbp",
        "final_value_cbp",
        "variance_value_cbp",
    ]
    variance_detail = variance[
        (variance[variance_cols].abs() > VALUE_TOLERANCE).any(axis=1)
    ].sort_values("abs_variance_total_value", ascending=False)
    variance_detail = variance_detail[detail_cols]

    summary = pd.DataFrame(
        [
            {"Metric": "Original INVOBS input rows used for variance", "Value": len(source_rows)},
            {"Metric": "Original INVOBS input value", "Value": variance["original_total_value"].sum()},
            {"Metric": "Final aggregated value", "Value": variance["final_total_value"].sum()},
            {"Metric": "Variance: final minus original", "Value": variance["variance_total_value"].sum()},
            {"Metric": "ISBNs with value variance", "Value": len(variance_detail)},
        ]
    )
    return summary, variance_detail


def write_value_variance_sheet(writer, source_rows, aggregated_inventory):
    summary, variance_detail = build_value_variance(source_rows, aggregated_inventory)
    sheet_name = "Value Variance"
    summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)

    startrow = len(summary) + 3
    variance_detail.to_excel(writer, sheet_name=sheet_name, index=False, startrow=startrow)

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    money_format = workbook.add_format({"num_format": "$#,##0.00;[Red]($#,##0.00);-"})
    integer_format = workbook.add_format({"num_format": "#,##0"})
    header_format = workbook.add_format({"bold": True})

    for row_num in (0, 4):
        worksheet.write_number(row_num + 1, 1, float(summary.at[row_num, "Value"]), integer_format)
    for row_num in (1, 2, 3):
        worksheet.write_number(row_num + 1, 1, float(summary.at[row_num, "Value"]), money_format)

    worksheet.write(startrow - 1, 0, "ISBN-level variance detail", header_format)
    worksheet.set_column(0, 0, 18)
    worksheet.set_column(1, 1, 34)
    worksheet.set_column(2, 13, 16, money_format)

def run(period):
    print(">>> Running pickle-based INVOBS flow")
    # Create dictionaries
    dict_cdu = create_cdu_dict()
    print(f">>> Loading original ConInv rows for period: {period}")
    df_original_inventory = load_period_consolidated_inventory(str(period))
    print("Original DataFrame shape:", df_original_inventory.shape)

    df_invobs_source_inventory = filter_invobs_inventory_rows(df_original_inventory)
    df_full_inventory = consolidate_inventory_from_rows(df_invobs_source_inventory)
    if df_full_inventory is None:
        return
    dict_uc = df_to_nested_dict(df_full_inventory)

    df_inventory = df_full_inventory[['ISBN', 'units_cbc', 'units_hbg', 'units_cbp']]

    # Process the inventory data and get the detailed result
    df_result_inventory = process_inventory(df_inventory, dict_cdu, dict_uc)

    # Create the aggregated result by grouping by 'ISBN' and summing
    df_aggregated_inventory = df_result_inventory.groupby('ISBN').sum().reset_index()

    # Use a file dialog to select the save location
    Tk().withdraw()  # Hide the root Tkinter window
    output_file_path = asksaveasfilename(
        title="Save the Consolidated Inventory File",
        initialfile=f"Consolidated_Inventory_{period}.xlsx",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )

    if output_file_path:
        # Save detailed, aggregated, original, and variance results to Excel.
        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            df_result_inventory.to_excel(writer, sheet_name='Detailed_Results', index=False)
            df_aggregated_inventory.to_excel(writer, sheet_name='Aggregated_Results', index=False)
            df_original_inventory.to_excel(writer, sheet_name='Original_Consolidated_Inventory', index=False)
            write_value_variance_sheet(writer, df_invobs_source_inventory, df_aggregated_inventory)

        print(f"Results saved to {output_file_path}")
    else:
        print("No save location selected. Exiting.")


def main():
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--period", required=True)
    args, _ = parser.parse_known_args()
    run(period=args.period)


if __name__ == "__main__":
    main()
