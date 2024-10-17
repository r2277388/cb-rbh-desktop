import networkx as nx
import pandas as pd
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score

# Load the DataFrame from the pickle file
df = pd.read_pickle('pickled_saldet.pkl')

# Group by ISBN and calculate summary statistics
isbn_stats = df.groupby('ISBN').agg({
    'sales_y1': ['mean', 'median', 'std', 'sum'],
    'acct_cnt': ['mean', 'median', 'std', 'sum']
})

# Display the results
print(isbn_stats)

# Calculate the correlation between sales_y1 and acct_cnt
correlation = df['sales_y1'].corr(df['acct_cnt'])

# Display the correlation coefficient
print(f"Correlation between sales_y1 and acct_cnt: {correlation}")



# Select the features for clustering
X = df[['sales_y1', 'acct_cnt']]

# Apply KMeans clustering
kmeans = KMeans(n_clusters=4, random_state=42)
df['cluster'] = kmeans.fit_predict(X)

# Analyze the clusters
cluster_summary = df.groupby('cluster').agg({
    'ISBN': 'count',
    'sales_y1': ['mean', 'sum'],
    'acct_cnt': ['mean', 'sum']
})

# Display the cluster summary
print(cluster_summary)

# Scatter plot
plt.figure(figsize=(10, 6))
plt.scatter(df['sales_y1'], df['acct_cnt'], c=df['cluster'], cmap='viridis')
plt.xlabel('Sales in First Year (sales_y1)')
plt.ylabel('Account Count (acct_cnt)')
plt.title('Sales vs Account Count by Cluster')
plt.colorbar(label='Cluster')
plt.show()



# Box plot
plt.figure(figsize=(10, 6))
sns.boxplot(x='cluster', y='sales_y1', data=df)
plt.xlabel('Cluster')
plt.ylabel('Sales in First Year (sales_y1)')
plt.title('Sales Distribution by Cluster')
plt.show()


# Prepare the data for regression
X = df[['acct_cnt']]
y = df['sales_y1']

# Fit the regression model
model = LinearRegression()
model.fit(X, y)

# Make predictions
df['predicted_sales_y1'] = model.predict(X)

# Evaluate the model
r2 = r2_score(y, df['predicted_sales_y1'])
print(f"R^2 score: {r2}")

# Plot the regression results
plt.figure(figsize=(10, 6))
plt.scatter(df['acct_cnt'], df['sales_y1'], label='Actual Sales')
plt.plot(df['acct_cnt'], df['predicted_sales_y1'], color='red', label='Predicted Sales')
plt.xlabel('Account Count (acct_cnt)')
plt.ylabel('Sales in First Year (sales_y1)')
plt.title('Regression of Sales on Account Count')
plt.legend()
plt.show()


# Calculate Z-scores for sales_y1
df['z_score'] = (df['sales_y1'] - df['sales_y1'].mean()) / df['sales_y1'].std()

# Identify outliers (e.g., Z-score > 3 or < -3)
outliers = df[df['z_score'].abs() > 3]

# Display the outliers
print(outliers[['ISBN', 'sales_y1', 'acct_cnt', 'z_score']])
