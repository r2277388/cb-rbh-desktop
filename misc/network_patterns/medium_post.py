import networkx as nx
import pandas as pd
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt

# Load the DataFrame from the pickle file
df_raw = pd.read_pickle('pickled_saldet.pkl')

df = df_raw.copy()

df = df[df['channel']=='Specialty']

print(df.head())