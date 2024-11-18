import pandas as pd
import numpy as np
import altair as alt
from queries import query_viz_daily
from functions import get_connection
from IPython.display import display, HTML

def load_data():
    """Connects to the database and loads the daily data."""
    engine = get_connection()
    return pd.read_sql_query(query_viz_daily, engine)

def prepare_data(df):
    """Prepare data by extracting year, month, and YTD filters."""
    year_list = sorted(df['year'].unique())
    ty = max(year_list)  # This year
    ly = year_list[-2]   # Last year
    df_ty = df[df['year'] == ty]
    months_ty = sorted(df_ty['month'].unique())
    tm = max(months_ty)  # This month
    df_ytd = df[df['month'].isin(months_ty)]
    return year_list, ty, ly, tm, df_ytd

def create_sort_order(df, year):
    """Create sort order for publishers and pub groups."""
    sort_order_pub = df[df.year == year].groupby('Publisher').agg({'rev': 'sum'}).sort_values(by='rev', ascending=False).reset_index()
    sort_order_pub.index = sort_order_pub.index + 1
    sort_order_pub.drop('rev', axis=1, inplace=True)
    sort_order_pub.reset_index(inplace=True)

    sort_order_pgrp = df[df.year == year].groupby('PubGroup').agg({'rev': 'sum'}).sort_values(by='rev', ascending=False).reset_index()
    sort_order_pgrp.index = sort_order_pgrp.index + 1
    sort_order_pgrp.drop('rev', axis=1, inplace=True)
    sort_order_pgrp.reset_index(inplace=True)

    sort_order = df[df.year == year][['Publisher', 'PubGroup']].reset_index(drop=True)
    sort_order = pd.merge(sort_order, sort_order_pub, on='Publisher')
    sort_order.columns = ['Publisher', 'PubGroup', 'Pub_Rank']
    sort_order = pd.merge(sort_order, sort_order_pgrp, on='PubGroup')
    sort_order.columns = ['Publisher', 'PubGroup', 'Pub_Rank', 'PGRP_Rank']
    sort_order = sort_order.drop_duplicates().reset_index(drop=True)
    sort_order = sort_order.sort_values(by=['Pub_Rank', 'PGRP_Rank']).reset_index(drop=True)
    sort_order.index = sort_order.index + 1
    sort_order = sort_order.reset_index()
    sort_order.columns = ['Rank', 'Publisher', 'PubGroup', 'Pub_Rank', 'PGRP_Rank']
    return sort_order[['PubGroup', 'Rank']]

def create_base_chart(df, width, height, ty, ly):
    """Create base chart."""
    base = alt.Chart(df).mark_bar(
        cornerRadiusTopLeft=3,
        cornerRadiusTopRight=3).encode(
        alt.X('month:O', axis=alt.Axis(title='Month', labelAngle=0)),
        alt.Y('sum(rev):Q', title="Total Sales")
    ).properties(
        width=width,
        height=height,
        title='Current Year Core Sales vs Last Year (Line)'
    )
    return base

def create_charts(df, width, height, ty, ly, tm, months_ty, sort_order):
    """Create individual charts."""
    base = create_base_chart(df, width, height, ty, ly)
    
    # Bar Chart
    base_ty = base.encode(
        color=alt.Color('PubGroup:N'),
        tooltip=['PubGroup:N', alt.Tooltip('sum(rev):Q', format="$,.0f")]
    ).transform_filter(
        alt.datum.year == ty
    )

    # Last-year Line Chart
    base_line_ly = base.mark_line(color='darkgrey', point=alt.OverlayMarkDef(color="red", size=50)).encode(
        tooltip=[alt.Tooltip('sum(rev):Q', format="$,.0f")]
    ).transform_filter(
        alt.datum.year == ly
    )

    total_sales = base.mark_text(dy=-5).encode(
        text=alt.Text('sum(rev):Q', format='$,.0f')
    ).transform_filter(
        alt.datum.year == ty
    )

    # Monthly Bar Chart
    monthly_bar = base.encode(
        alt.X('PubGroup:N', sort='-y', axis=alt.Axis(labelAngle=0)),
        alt.Y('sum(rev):Q', title="", stack=None),
        tooltip=['PubGroup', 'year', alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')],
        color=alt.Color('year:N', title='Year', scale=alt.Scale(
            domain=[ty, ly],
            range=['#3182bc', 'lightgrey']))
    ).transform_filter(
        alt.datum.month == tm
    ).transform_filter(
        alt.FieldOneOfPredicate(field='year', oneOf=[ty, ly])
    ).properties(
        title='Current Month Sales vs Current Month Last Year', width=600, height=height * 0.5
    )

    # Yearly Bar Chart
    year_bar = base.encode(
        alt.Y('sum(rev):Q', title="", stack=None),
        tooltip=['year', alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')],
        color=alt.Color('year:N', title='Year', legend=None, scale=alt.Scale(
            domain=[ty, ly],
            range=['#3182bc', 'lightgrey']))
    ).transform_filter(
        alt.datum.month == tm
    ).transform_filter(
        alt.FieldOneOfPredicate(field='year', oneOf=[ty, ly])
    ).properties(
        title='TM', width=50, height=height * 0.5
    )

    return base_ty, base_line_ly, total_sales, monthly_bar, year_bar

def combine_charts(base_ty, base_line_ly, total_sales, monthly_bar, year_bar, df, width, height, ty, ly, tm, months_ty, sort_order):
    """Combine all charts into one."""
    plot_title = alt.TitleParams(
        'Publisher and Pub Group Summary',
        subtitle=['All Data Represents Core Sales', ''],
        align='center',
        anchor='middle',
        fontSize=30
    )

    # A dropdown filter
    pub_options = list(df['Publisher'].unique())
    item = 'Chronicle'
    pub_options.remove(item)
    pub_options.insert(0, item)

    publisher_dropdown = alt.binding_select(options=[None] + pub_options, labels=['All'] + pub_options, name='Publisher: ')
    publisher_select = alt.selection_single(fields=['Publisher'], bind=publisher_dropdown)

    channel_list = ['Amazon', 'Trade', 'Specialty', 'Export', 'Mass', 'Direct']
    channel_dropdown = alt.binding_select(options=[None] + channel_list, labels=['All'] + channel_list, name='Channels')
    channel_select = alt.selection_single(fields=['channel'], bind=channel_dropdown)

    cumulative_chart = ((base_ty + base_line_ly + total_sales)
                        & (year_bar | monthly_bar)
                        ).add_selection(
        publisher_select).transform_filter(
        publisher_select).properties(title=plot_title)

    return cumulative_chart

def main():
    df_raw = load_data()
    year_list, ty, ly, tm, df_ytd = prepare_data(df_raw)
    sort_order = create_sort_order(df_raw, ty)
    df_daily = pd.merge(df_raw, sort_order, on='PubGroup')
    width = 700
    height = 300
    base_ty, base_line_ly, total_sales, monthly_bar, year_bar = create_charts(df_daily, width, height, ty, ly, tm, year_list, sort_order)
    cumulative_chart = combine_charts(base_ty, base_line_ly, total_sales, monthly_bar, year_bar, df_daily, width, height, ty, ly, tm, year_list, sort_order)
    display(cumulative_chart)
    cumulative_chart.save('ssr_summary_chart.html')

if __name__ == "__main__":
    main()