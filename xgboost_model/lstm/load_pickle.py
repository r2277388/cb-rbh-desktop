import numpy as np
import pickle
from pathlib import Path

# Directly set the path in Colab or Jupyter if __file__ is not available
full_path = '/content/drive/MyDrive/code_xgboost/pickle_raw_data/df_pickle.pkl'

# # Use pickle to load the DataFrame
# with open(full_path, 'rb') as file:
#     df_raw = pickle.load(file)

# print(df_raw.head())


from joblib import load

df_raw = load(full_path)