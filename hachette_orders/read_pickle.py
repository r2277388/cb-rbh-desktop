import pandas as pd
import os

file_name = fr'C:\Users\rbh\code\ho_20250226_1414.pkl'

base_name = os.path.basename(file_name)
base_data = base_name.replace('.pkl', '')

print(f'Loading {base_data}...')

df = pd.read_pickle(file_name)

print(df.info())

print(df.head())

excel_file_name = fr'C:\Users\rbh\Desktop\{base_data}.xlsx'

df.to_excel(excel_file_name, index=False)

print(f'Excel Version of Hachette Orders saved to: {excel_file_name}')
