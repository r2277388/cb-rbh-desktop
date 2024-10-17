import time
import pandas as pd
import numpy as np
import pickle
from pathlib import Path

from sklearn.model_selection import train_test_split, TimeSeriesSplit
from sklearn.preprocessing import StandardScaler, LabelEncoder

import tensorflow as tf
from tensorflow import keras
# from tensorflow.keras.models import Sequential, load_model
# from tensorflow.keras.layers import Input,LSTM, Dense
# from tensorflow.keras.callbacks import EarlyStopping

# The path to your pickled file
dir_path = Path('E:\My Drive\Colab Notebooks\cb_forecasting')
file_path = 'df_pickle.pkl'

full_path = dir_path / file_path

# Load the DataFrame from the pickle file
df_raw = pd.read_pickle(full_path)

df_raw.head()