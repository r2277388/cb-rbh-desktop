import pandas as pd
import altair as alt
import os
import pickle
from datetime import datetime
from functions import get_connection  # Assuming functions.py has this
from queries import query_viz_daily  # Assuming queries.py has this

def load_data():
    """Connects to the database and loads the daily data."""
    engine = get_connection()
    # Define the filename with the current date
    today = datetime.today().strftime('%Y-%m-%d')
    filename = f"data_{today}.pkl"

    # Check if the file already exists
    if os.path.exists(filename):
        # Load the data from the pickle file
        with open(filename, 'rb') as file:
            df = pickle.load(file)
    else:
        # Load the data from the database
        query_daily = query_viz_daily()
        df = pd.read_sql_query(query_daily, engine)
        # Save the data to a pickle file
        with open(filename, 'wb') as file:
            pickle.dump(df, file)

    return df

def prepare_data(df):
    """Prepare data by extracting year, month, and YTD filters."""
    year_list = sorted(df['year'].unique())
    ty = max(year_list)  # This year
    ly = year_list[-2]   # Last year
    df_ty = df[df['year'] == ty]
    months_ty = sorted(df_ty['month'].unique())
    tm = max(months_ty)  # This month
    df_ytd = df[(df['year'] == ty) & (df['month'].isin(months_ty))]
    return year_list, ty, ly, tm, df_ytd

def calculate_sort_order(df, year):
    """Calculate and add rankings for Publisher and PubGroup."""
    sort_order_pub = df[df['year'] == year].groupby('Publisher').agg({'rev':'sum'}).sort_values(by='rev', ascending=False).reset_index()
    sort_order_pub.index = sort_order_pub.index + 1
    sort_order_pub.rename(columns={'rev': 'Pub_Rank'}, inplace=True)

    sort_order_pgrp = df[df['year'] == year].groupby('PubGroup').agg({'rev':'sum'}).sort_values(by='rev', ascending=False).reset_index()
    sort_order_pgrp.index = sort_order_pgrp.index + 1
    sort_order_pgrp.rename(columns={'rev': 'PGRP_Rank'}, inplace=True)

    sort_order = df[df['year'] == year][['Publisher', 'PubGroup']].drop_duplicates().reset_index(drop=True)
    sort_order = sort_order.merge(sort_order_pub, on='Publisher', how='left').merge(sort_order_pgrp, on='PubGroup', how='left')
    sort_order = sort_order.sort_values(by=['Pub_Rank', 'PGRP_Rank']).reset_index(drop=True)
    sort_order['Rank'] = sort_order.index + 1
    return sort_order[['PubGroup', 'Rank']]

def apply_sort_order(df, sort_order):
    """Applies the sorting order to the main DataFrame."""
    return pd.merge(df, sort_order, on='PubGroup', how='left')

# Visualization Functions
def create_base_chart(df):
    """Creates a base Altair chart to be reused."""
    return alt.Chart(df).mark_bar(
        cornerRadiusTopLeft=3, cornerRadiusTopRight=3
    ).encode(
        x=alt.X('month:O', axis=alt.Axis(title='Month', labelAngle=0)),
        y=alt.Y('sum(rev):Q', title="Total Sales")
    ).properties(width=700, height=300, title='Current Year Core Sales vs Last Year (Line)')

def create_bar_chart(df, ty, ly):
    """Creates bar charts for the current and last year."""
    base = create_base_chart(df)
    bar_ty = base.encode(
        color=alt.Color('PubGroup:N'),
        tooltip=['PubGroup:N', alt.Tooltip('sum(rev):Q', format="$,.0f")]
    ).transform_filter(alt.datum.year == ty)
    
    bar_ly = base.mark_line(color='darkgrey', point=alt.OverlayMarkDef(color="red", size=50)).encode(
        tooltip=[alt.Tooltip('sum(rev):Q', format="$,.0f")]
    ).transform_filter(alt.datum.year == ly)
    
    return bar_ty, bar_ly

def create_heatmaps(df):
    """Create heatmaps for yearly and quarterly data."""
    heat_month = alt.Chart(df).mark_rect(cornerRadius=4).encode(
        x=alt.X('month:O', title='Month', axis=alt.Axis(labelAngle=0)),
        y=alt.Y('year', title='', sort='descending'),
        color=alt.Color('sum(rev)', legend=None, scale=alt.Scale(scheme='blues')),
        tooltip=[alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')]
    ).properties(height=125, width=750, title='Total Month Sales by Year')
    
    heat_quarter = alt.Chart(df).mark_rect(cornerRadius=4).encode(
        x=alt.X('quarter', title='Quarter', axis=alt.Axis(labelAngle=0)),
        y=alt.Y('year', sort='descending', axis=alt.Axis(title='')),
        color=alt.Color('sum(rev)', legend=None, scale=alt.Scale(scheme='purples')),
        tooltip=[alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')]
    ).properties(height=125, width=500, title='Total Quarter Sales by Year')
    
    return heat_month, heat_quarter

def save_chart(chart, filename="ssr_summary_chart.html"):
    """Save the final Altair chart to an HTML file."""
    chart.save(filename)

def main():
    # Load and prepare data
    df_daily = load_data()
    year_list, ty, ly, tm, df_ytd = prepare_data(df_daily)
    
    # Sort orders
    sort_order = calculate_sort_order(df_daily, ty)
    df_daily = apply_sort_order(df_daily, sort_order)
    
    # Generate charts
    bar_ty, bar_ly = create_bar_chart(df_daily, ty, ly)
    heat_month, heat_quarter = create_heatmaps(df_daily)
    
    # Combine and save
    cumulative_chart = (bar_ty + bar_ly) & heat_month & heat_quarter
    save_chart(cumulative_chart)

if __name__ == "__main__":
    main()