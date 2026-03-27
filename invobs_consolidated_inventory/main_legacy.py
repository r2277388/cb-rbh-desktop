import argparse
import sys
from pathlib import Path

import pandas as pd
from tkinter import Tk
from tkinter.filedialog import asksaveasfilename

from dict_cdu import create_cdu_dict
from dict_unit_cost2 import df_to_nested_dict
from load_consolidated_inventory_legacy import consolidate_inventory


REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


def process_inventory(df_inventory, cdu_dict, dict_uc):
    result = []

    for _, row in df_inventory.iterrows():
        isbn = row["ISBN"]
        units_cbc = row["units_cbc"]
        units_hbg = row["units_hbg"]
        units_cbp = row["units_cbp"]

        if isbn in cdu_dict:
            for component_isbn, component_qty in cdu_dict[isbn].items():
                total_units_cbc = units_cbc * component_qty
                total_units_hbg = units_hbg * component_qty
                total_units_cbp = units_cbp * component_qty

                if component_isbn in dict_uc:
                    uc_cbc = dict_uc[component_isbn].get("uc_cbc", 0)
                    uc_hbg = dict_uc[component_isbn].get("uc_hbg", 0)
                    uc_cbp = dict_uc[component_isbn].get("uc_cbp", 0)

                    result.append(
                        {
                            "ISBN": component_isbn,
                            "total_units_cbc": total_units_cbc,
                            "total_units_hbg": total_units_hbg,
                            "total_units_cbp": total_units_cbp,
                            "total_value_cbc": total_units_cbc * uc_cbc,
                            "total_value_hbg": total_units_hbg * uc_hbg,
                            "total_value_cbp": total_units_cbp * uc_cbp,
                        }
                    )
        else:
            if isbn in dict_uc:
                uc_cbc = dict_uc[isbn].get("uc_cbc", 0)
                uc_hbg = dict_uc[isbn].get("uc_hbg", 0)
                uc_cbp = dict_uc[isbn].get("uc_cbp", 0)

                result.append(
                    {
                        "ISBN": isbn,
                        "total_units_cbc": units_cbc,
                        "total_units_hbg": units_hbg,
                        "total_units_cbp": units_cbp,
                        "total_value_cbc": units_cbc * uc_cbc,
                        "total_value_hbg": units_hbg * uc_hbg,
                        "total_value_cbp": units_cbp * uc_cbp,
                    }
                )

    result_df = pd.DataFrame(result)
    result_df = result_df[
        [
            "ISBN",
            "total_value_cbc",
            "total_units_cbc",
            "total_value_hbg",
            "total_units_hbg",
            "total_value_cbp",
            "total_units_cbp",
        ]
    ]
    result_df["total_value"] = (
        result_df["total_value_cbc"]
        + result_df["total_value_hbg"]
        + result_df["total_value_cbp"]
    )
    result_df["total_units"] = (
        result_df["total_units_cbc"]
        + result_df["total_units_hbg"]
        + result_df["total_units_cbp"]
    )
    return result_df


def run(source_file=None):
    print(">>> Running legacy INVOBS workbook flow")
    dict_cdu = create_cdu_dict()
    df_full_inventory = consolidate_inventory(source_file)
    if df_full_inventory is None:
        return

    dict_uc = df_to_nested_dict(df_full_inventory)
    df_inventory = df_full_inventory[["ISBN", "units_cbc", "units_hbg", "units_cbp"]]
    df_result_inventory = process_inventory(df_inventory, dict_cdu, dict_uc)
    df_aggregated_inventory = df_result_inventory.groupby("ISBN").sum().reset_index()

    Tk().withdraw()
    output_file_path = asksaveasfilename(
        title="Save the Consolidated Inventory File",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
    )

    if output_file_path:
        with pd.ExcelWriter(output_file_path, engine="xlsxwriter") as writer:
            df_result_inventory.to_excel(writer, sheet_name="Detailed_Results", index=False)
            df_aggregated_inventory.to_excel(writer, sheet_name="Aggregated_Results", index=False)
        print(f"Results saved to {output_file_path}")
    else:
        print("No save location selected. Exiting.")


def main():
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--source-file")
    args, _ = parser.parse_known_args()
    run(source_file=args.source_file)


if __name__ == "__main__":
    main()
