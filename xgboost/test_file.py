import os

def check_file_exists(file_path):
    if os.path.exists(file_path):
        print("File exists.")
    else:
        print("File does not exist.")

file_path = r"E:\My Drive\Colab Notebooks\cb_forecasting\df_pickle.pkl"

check_file_exists(file_path)