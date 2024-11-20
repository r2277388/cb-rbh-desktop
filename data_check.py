import pandas as pd

# Load the pickle file
filename = "data_2024-11-20.pkl"
df = pd.read_pickle(filename)

# Display the first few rows of the DataFrame
print(df.tail())
