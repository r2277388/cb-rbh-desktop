import pandas as pd
from loader.load_ho import upload_ho
from order_type_regular import calculate_est_ship_date_regular
from order_type_backordered import calculate_est_ship_date_backordered
from order_type_released import calculate_est_ship_date_released
from order_type_hold import calculate_est_ship_date_hold

# Load the DataFrame
df = upload_ho()

