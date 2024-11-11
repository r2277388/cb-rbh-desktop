import pandas as pd

file_name = r'C:\Users\rbh\code\hachette_orders\ho_20241111_1318.pkl'
df = pd.read_pickle(file_name)

print(df.info())
