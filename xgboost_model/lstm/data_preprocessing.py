import pandas as pd
import numpy as np
from pathlib import Path

from sklearn.model_selection import train_test_split, TimeSeriesSplit
from sklearn.preprocessing import StandardScaler, LabelEncoder

import os
os.environ["TF_ENABLE_ONEDNN_OPTS"] = "0"
import tensorflow as tf
from tensorflow import keras
from tensorflow.keras.models import Sequential, load_model
from tensorflow.keras.layers import Input,LSTM, Dense
from tensorflow.keras.callbacks import EarlyStopping

tf.config.threading.set_inter_op_parallelism_threads(24)
tf.config.threading.set_intra_op_parallelism_threads(24)

# List all physical GPUs
gpus = tf.config.list_physical_devices('GPU')
if gpus:
    for gpu in gpus:
        gpu_details = tf.config.experimental.get_device_details(gpu)
        print(f"Name: {gpu_details['device_name']}")
        print(f"Memory: {gpu_details['memory_limit'] / (1024 ** 3):.2f} GB")
else:
    print("No GPU found")

# Use Path to construct the relative path correctly
full_path = Path(__file__).resolve().parent.parent / 'pickle_raw_data' / 'df_pickle.pkl'

print(full_path)

if full_path.is_file():
    print("File found!")
else:
    print("File not found. Please check the path.")

# Load the DataFrame from the pickle file
df_raw = pd.read_pickle(full_path)

print(df_raw.head())