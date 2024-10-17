import pandas as pd
from dict_cdu import create_cdu_dict
from dict_unit_cost import df_to_nested_dict
from load_consolidated_inventory import consolidate_inventory

# Create dictionaries
dict_cdu = create_cdu_dict()
dict_uc = df_to_nested_dict()

# Load inventory data
df_inventory = consolidate_inventory()
df_inventory = df_inventory[['ISBN', 'units_cbc', 'units_hbg', 'units_cbp']]

# Assuming df_inventory, cdu_dict, and dict_uc are already defined

def process_inventory(df_inventory, cdu_dict, dict_uc):
    # Initialize a dictionary to store total units and values for each component ISBN
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
    return result_df



# Define cdu_dict and dict_uc based on your inputs
# cdu_dict and dict_uc will be predefined dictionaries based on your example

# Convert the processed inventory into a dataframe
df_result_inventory = process_inventory(df_inventory, dict_cdu, dict_uc)

# Save the result dataframe to an Excel file
output_file_path = r'C:\Users\rbh\Desktop\consolidate_inventory.xlsx'
df_result_inventory.to_excel(output_file_path, index=False)

# Display the result dataframe
print(df_result_inventory)
