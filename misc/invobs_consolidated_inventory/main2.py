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

# Initialize a list to collect rows for the result DataFrame
result_rows = []

# Iterate through df_inventory
for index, row in df_inventory.iterrows():
    isbn = row['ISBN']
    units_cbc = row['units_cbc']
    units_hbg = row['units_hbg']
    units_cbp = row['units_cbp']
    
    if isbn in dict_cdu:
        # Handle CDUs
        for component_isbn, quantity in dict_cdu[isbn].items():
            component_units_cbc = units_cbc * quantity
            component_units_hbg = units_hbg * quantity
            component_units_cbp = units_cbp * quantity
            
            unit_costs = dict_uc.get(component_isbn, {'uc_cbc': 0.0, 'uc_hbg': 0.0, 'uc_cbp': 0.0})
            value_cbc = component_units_cbc * unit_costs['uc_cbc']
            value_hbg = component_units_hbg * unit_costs['uc_hbg']
            value_cbp = component_units_cbp * unit_costs['uc_cbp']
            
            result_rows.append({
                'ISBN': component_isbn,
                'units_cbc': component_units_cbc,
                'units_hbg': component_units_hbg,
                'units_cbp': component_units_cbp,
                'value_cbc': value_cbc,
                'value_hbg': value_hbg,
                'value_cbp': value_cbp
            })
    else:
        # Handle regular ISBNs
        unit_costs = dict_uc.get(isbn, {'uc_cbc': 0.0, 'uc_hbg': 0.0, 'uc_cbp': 0.0})
        value_cbc = units_cbc * unit_costs['uc_cbc']
        value_hbg = units_hbg * unit_costs['uc_hbg']
        value_cbp = units_cbp * unit_costs['uc_cbp']
        
        result_rows.append({
            'ISBN': isbn,
            'units_cbc': units_cbc,
            'units_hbg': units_hbg,
            'units_cbp': units_cbp,
            'value_cbc': value_cbc,
            'value_hbg': value_hbg,
            'value_cbp': value_cbp
        })

# Create the result DataFrame from the collected rows
df_result = pd.DataFrame(result_rows)

# Aggregate results by ISBN
df_result = df_result.groupby('ISBN').sum().reset_index()

print(df_result)