import os
import pandas as pd
from pathlib import Path
import tensorflow as tf
import psutil

os.environ["TF_ENABLE_ONEDNN_OPTS"] = "0"
from tensorflow import keras
from tensorflow.keras.models import Sequential, load_model
from tensorflow.keras.layers import Input, LSTM, Dense
from tensorflow.keras.callbacks import EarlyStopping

tf.config.threading.set_inter_op_parallelism_threads(24)
tf.config.threading.set_intra_op_parallelism_threads(24)

# List all physical GPUs
gpus = tf.config.list_physical_devices('GPU')
if gpus:
    for gpu in gpus:
        gpu_details = tf.config.experimental.get_device_details(gpu)
        print(f"GPU Name: {gpu_details['device_name']}")
        print(f"GPU Memory: {gpu_details['memory_limit'] / (1024 ** 3):.2f} GB")
else:
    print("No GPU found")

# Check CPU details
cpu_count = psutil.cpu_count(logical=True)
cpu_freq = psutil.cpu_freq()
cpu_usage = psutil.cpu_percent(interval=1)

print(f"CPU Count: {cpu_count}")
print(f"CPU Frequency: {cpu_freq.current:.2f} MHz")
print(f"CPU Usage: {cpu_usage:.2f}%")