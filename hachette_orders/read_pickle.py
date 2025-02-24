import pandas as pd

file_name = r'C:\Users\rbh\code\hachette_orders\ho_20250220_0825.pkl'
df = pd.read_pickle(file_name)

print(df.info())

print(df.head())

excel_file_name = fr'C:\Users\rbh\Desktop\ho_20250220_0825.xlsx'

df.to_excel(excel_file_name, index=False)