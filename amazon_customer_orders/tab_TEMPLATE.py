from loader.loader_weeklysales import uploader_weeklysales
from loader.loader_traffic import upload_traffic
from loader.loader_item import upload_item
from asin_isbn_converter import asin_isbn_conversion
import pandas as pd
import numpy as np
from datetime import datetime

def top_titles(publisher=None, flbl=None, num_rows=None):
	df_weeklysales = uploader_weeklysales()
	
	df_converter = asin_isbn_conversion()
	df = df_weeklysales.merge(df_converter, on='ASIN', how='left')
	
	df_item = upload_item()
	df = df.merge(df_item, on='ISBN', how='left')
	
	df_glance = upload_traffic()
	df_glance = df_glance[['ASIN', 'Glance Views']]
	df = df.merge(df_glance, on='ASIN', how='left')
	
	df['Conversion Rate'] = np.round((df['Ordered Units'] / df['Glance Views']) * 100, 2)
	
	# Filter by publisher if specified
	if publisher:
		if publisher.startswith('!'):
			df = df[df['publisher'] != publisher[1:]]
		else:
			df = df[df['publisher'] == publisher]
	
	# Add FL/BL categorization
	current_year = datetime.now().year
	if flbl:
		if flbl == "FL":
			df = df[df['Release Date'].dt.year >= current_year]
		elif flbl == "BL":
			df = df[df['Release Date'].dt.year < current_year]
	
	df = df[['ASIN', 'ISBN', 'title', 'Release Date', 'publisher', 'Ordered Units',
			 'Ordered Units - Prior Period', 'Glance Views', 'Conversion Rate']]
	
	# Sort by 'Ordered Units'
	df = df.sort_values(by='Ordered Units', ascending=False)
	
	# Limit the number of rows if specified
	if num_rows:
		df = df.head(num_rows)
	
	return df

def main():
	# Example usage
	df = top_titles(publisher="!Chronicle", flbl="FL", num_rows=10)
	
	print(df.info())
	print(df.head())
	
if __name__ == '__main__':
	main()