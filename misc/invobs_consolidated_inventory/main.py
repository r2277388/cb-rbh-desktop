from dict_cdu import create_cdu_dict
from dict_unit_cost import df_to_nested_dict
from load_consolidated_inventory import consolidate_inventory

dict_cdu = create_cdu_dict()
dict_uc = df_to_nested_dict()

df_inventory = consolidate_inventory()
df_inventory = df_inventory[['ISBN','units_cbc','units_hbg','units_cbp']]



print(df_inventory.head())


