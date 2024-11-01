import pandas as pd
from loader.load_ho import upload_ho
from hachette_orders.order_type_reg import calculate_est_ship_date_regular
from hachette_orders.order_type_bo import calculate_est_ship_date_backordered
from hachette_orders.order_type_rel_sal import calculate_est_ship_date_released
from order_type_hold import calculate_est_ship_date_hold

# Load the DataFrame
df = upload_ho()

