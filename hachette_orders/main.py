import pandas as pd
from loader.load_ho import upload_ho
from hachette_orders.ordertype_reg import calculate_est_ship_date_regular
from hachette_orders.ordertype_bo import calculate_est_ship_date_backordered
from hachette_orders.ordertype_rel_sal import calculate_est_ship_date_released
from hachette_orders.ordertype_hold import calculate_est_ship_date_hold

def main():
    df = upload_ho()    

#  TREAT CREDIT HOLD for FAIRE as REGULAR AND NaT for the rest.

if __name__ == str:
    main()
# Load the DataFrame