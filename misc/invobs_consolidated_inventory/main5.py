import pandas as pd
from dict_cdu import create_cdu_dict
from dict_unit_cost2 import df_to_nested_dict
from load_consolidated_inventory import consolidate_inventory

# Create dictionaries
dict_cdu = create_cdu_dict()
dict_uc = df_to_nested_dict()

# Load inventory data
df_inventory = consolidate_inventory()
df_inventory = df_inventory[['ISBN', 'units_cbc', 'units_hbg', 'units_cbp']]

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

# Process the inventory data and get the detailed result
df_result_inventory = process_inventory(df_inventory, dict_cdu, dict_uc)

# Create the aggregated result by grouping by 'ISBN' and summing
df_aggregated_inventory = df_result_inventory.groupby('ISBN').sum().reset_index()

# Save both detailed and aggregated results to Excel with two tabs
output_file_path = r'C:\Users\rbh\Desktop\consolidated_inventory_September.xlsx'
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    # First sheet: detailed results
    df_result_inventory.to_excel(writer, sheet_name='Detailed_Results', index=False)
    
    # Second sheet: aggregated results
    df_aggregated_inventory.to_excel(writer, sheet_name='Aggregated_Results', index=False)

# Display the result dataframe
# print(df_result_inventory)
# print(df_aggregated_inventory)
print(df_result_inventory[df_result_inventory.ISBN == '0810073340527'])
print(df_aggregated_inventory[df_aggregated_inventory.ISBN == '0810073340527'])
