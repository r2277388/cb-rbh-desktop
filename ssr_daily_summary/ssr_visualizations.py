#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import altair as alt
import os
import pickle
from datetime import datetime
from functions import get_connection  # Assuming functions.py has this
from queries import query_viz_daily  # Assuming queries.py has this
from paths import saved_viz_location

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
    df_daily = df.copy()
    year_list = sorted(df_daily.year.unique())
    ty = max(year_list)
    ly = year_list[-2]
    df_ty = df_daily.loc[df_daily['year'] == ty]
    months_ty = list(df_ty.month.unique())
    tm = max(months_ty)
    df_ytd = df_daily[df_daily['month'].isin(months_ty)]
    
    sort_order_pub = df_daily.loc[df_daily.year == ty].groupby('Publisher').agg({'rev': 'sum'}).sort_values(by='rev', ascending=False).reset_index()
    sort_order_pub.index = sort_order_pub.index + 1
    sort_order_pub.drop('rev', axis=1, inplace=True)
    sort_order_pub.reset_index(inplace=True)
    
    sort_order_pgrp = df_daily.loc[df_daily.year == ty].groupby('PubGroup').agg({'rev': 'sum'}).sort_values(by='rev', ascending=False).reset_index()
    sort_order_pgrp.index = sort_order_pgrp.index + 1
    sort_order_pgrp.drop('rev', axis=1, inplace=True)
    sort_order_pgrp.reset_index(inplace=True)
    
    sort_order = df_daily.loc[df_daily.year == ty].loc[:, ['Publisher', 'PubGroup']].reset_index(drop=True)
    sort_order = pd.merge(sort_order, sort_order_pub, on='Publisher')
    sort_order.columns = ['Publisher', 'PubGroup', 'Pub_Rank']
    sort_order = pd.merge(sort_order, sort_order_pgrp, on='PubGroup')
    sort_order.columns = ['Publisher', 'PubGroup', 'Pub_Rank', 'PGRP_Rank']
    sort_order = sort_order.drop_duplicates().reset_index(drop=True)
    sort_order = sort_order.sort_values(by=['Pub_Rank', 'PGRP_Rank']).reset_index(drop=True)
    sort_order.index = sort_order.index + 1
    sort_order = sort_order.reset_index()
    sort_order.columns = ['Rank', 'Publisher', 'PubGroup', 'Pub_Rank', 'PGRP_Rank']
    sort_order = sort_order.loc[:, ['PubGroup', 'Rank']]
    
    df_daily = pd.merge(df_daily, sort_order, on='PubGroup')
    
    return df_daily, ty, ly, months_ty, tm

def create_charts(df_daily, ty, ly, months_ty, tm):
    width = 700
    height = 300

    base = alt.Chart(df_daily).mark_bar(
        cornerRadiusTopLeft=3,
        cornerRadiusTopRight=3).encode(
        alt.X('month:O', axis=alt.Axis(title='Month', labelAngle=0)),
        alt.Y('sum(rev):Q', title="Total Sales")
    ).properties(
        width=width,
        height=height,
        title='Current Year Core Sales vs Last Year (Line)'
    )

    base_ty = base.encode(
        color=alt.Color('PubGroup:N'),
        tooltip=['PubGroup:N', alt.Tooltip('sum(rev):Q', format="$,.0f")]
    ).transform_filter(
        alt.datum.year == ty
    )

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

    monthly_bar = base.encode(
        alt.X('PubGroup:N', sort='-y', axis=alt.Axis(labelAngle=0)),
        alt.Y('sum(rev):Q', title="", stack=None),
        tooltip=['PubGroup', 'year', alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')],
        color=alt.Color('year:N', title='Year', scale=alt.Scale(
            domain=[ty, ly],
            range=['#3182bc', 'lightgrey']
        ))
    ).transform_filter(
        alt.datum.month == tm
    ).transform_filter(
        alt.FieldOneOfPredicate(field='year', oneOf=[ty, ly])
    ).properties(
        title='Current Month Sales vs Current Month Last Year', width=600, height=height * 0.5
    )

    year_bar = base.encode(
        alt.Y('sum(rev):Q', title="", stack=None),
        tooltip=['year', alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')],
        color=alt.Color('year:N', title='Year', legend=None, scale=alt.Scale(
            domain=[ty, ly],
            range=['#3182bc', 'lightgrey']
        ))
    ).transform_filter(
        alt.datum.month == tm
    ).transform_filter(
        alt.FieldOneOfPredicate(field='year', oneOf=[ty, ly])
    ).properties(
        title='TM', width=50, height=height * 0.5
    )

    heat_month = base.mark_rect(
        cornerRadiusTopLeft=4,
        cornerRadiusTopRight=4,
        cornerRadiusBottomLeft=4,
        cornerRadiusBottomRight=4
    ).encode(
        y=alt.Y('year', title='', sort='descending'),
        color=alt.Color('sum(rev)', legend=None, scale=alt.Scale(scheme='blues')),
        tooltip=[alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')],
        text=alt.Text('sum(rev)')
    ).properties(height=125, width=750, title='Total Month Sales by Year')

    heat_quarter = base.mark_rect(
        cornerRadiusTopLeft=4,
        cornerRadiusTopRight=4,
        cornerRadiusBottomLeft=4,
        cornerRadiusBottomRight=4
    ).encode(
        x=alt.X('quarter', title='Quarter', axis=alt.Axis(labelAngle=0)),
        y=alt.Y('year', sort='descending', axis=alt.Axis(title='')),
        color=alt.Color('sum(rev)', legend=None, scale=alt.Scale(scheme='purples')),
        tooltip=[alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')],
        text=alt.X('sum(rev)')
    ).properties(height=125, width=500, title='Total Quarter Sales by Year')

    year_publisher = base.mark_bar(
        cornerRadiusTopLeft=3,
        cornerRadiusTopRight=3,
        cornerRadiusBottomRight=3,
        cornerRadiusBottomLeft=3
    ).encode(
        alt.Y('year:O', sort='descending', axis=alt.Axis(title='', labelAngle=0, labels=False)),
        alt.X('sum(rev):Q', title="FY"),
        tooltip=[alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')],
        text=alt.X('sum(rev)')
    ).properties(
        title='',
        width=70,
        height=125
    )

    ytd_publisher = base.mark_bar(
        cornerRadiusTopLeft=3,
        cornerRadiusTopRight=3,
        cornerRadiusBottomRight=3,
        cornerRadiusBottomLeft=3
    ).encode(
        alt.Y('year:O', sort='descending', axis=alt.Axis(title='', labelAngle=0)),
        alt.X('sum(rev):Q', title="YTD"),
        tooltip=[alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')],
        text=alt.X('sum(rev)')
    ).properties(
        title='',
        width=70,
        height=125
    ).transform_filter(
        alt.FieldOneOfPredicate(field='month', oneOf=months_ty)
    )

    channel_list = ['Amazon', 'Trade', 'Specialty', 'Export', 'Mass', 'Direct']

    channel_bg_bar_month = base.mark_bar(line=True, opacity=0.75).encode(
        alt.X('channel', sort=channel_list, title='', axis=alt.Axis(labelAngle=0)),
        alt.Y('sum(rev):Q', stack=None, title=''),
        alt.Color('channel:N', legend=None),
        tooltip=['year', alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')]
    ).transform_filter(
        alt.datum.month == tm
    ).transform_filter(
        alt.FieldOneOfPredicate(field='year', oneOf=[ty])
    ).properties(
        height=100, width=750, title='MTD - Market Channel Sales vs Last-year Thru Full Month'
    )

    channel_bg_tick_month = base.mark_tick(color='grey', thickness=3, size=100).encode(
        alt.X('channel', sort=channel_list, title='', axis=alt.Axis(labelAngle=0)),
        alt.Y('sum(rev):Q', stack=None, title='', sort=[ly]),
        tooltip=['channel', 'year', alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')]
    ).transform_filter(
        alt.datum.month == tm
    ).transform_filter(
        alt.FieldOneOfPredicate(field='year', oneOf=[ly])
    ).properties(
        height=100, width=750, title=''
    )

    channel_bg_bar_ytd = base.mark_bar(line=True, opacity=1.0).encode(
        alt.X('channel', sort=channel_list, title='', axis=alt.Axis(labelAngle=0)),
        alt.Y('sum(rev):Q', stack=None, title=''),
        alt.Color('channel:N', legend=None),
        tooltip=['year', alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')]
    ).transform_filter(
        alt.FieldOneOfPredicate(field='year', oneOf=[ty])
    ).properties(
        height=100, width=750, title='YTD - Market Channel Sales vs Last-year Thru Full Month'
    ).transform_filter(
        alt.FieldOneOfPredicate(field='month', oneOf=months_ty)
    )

    channel_bg_tick_ytd = base.mark_tick(color='grey', thickness=3, size=100).encode(
        alt.X('channel', sort=channel_list, title='', axis=alt.Axis(labelAngle=0)),
        alt.Y('sum(rev):Q', stack=None, title='', sort=[ly]),
        tooltip=['channel', 'year', alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')]
    ).transform_filter(
        alt.FieldOneOfPredicate(field='year', oneOf=[ly])
    ).properties(
        height=100, width=750, title=''
    ).transform_filter(
        alt.FieldOneOfPredicate(field='month', oneOf=months_ty)
    )

    amaz_rom = alt.Chart(df_daily).mark_area(line=True).encode(
    alt.X('year_month:O', axis=alt.Axis(title='Year and Month')),
    alt.Y('sum(rev):Q', title="Percentage of Total Sales", stack='normalize'),
    color=alt.Color('Group:N', scale=alt.Scale(
        domain=['Amaz', 'ROM'],
        range=['gold', 'darkgrey']
    )),
    tooltip=['Group:N', 'year:O', 'month:O', alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')],
    text=alt.X('sum(rev)')
    ).properties(
        height=height * 0.4,
        width=width * 1.0,
        title='Amazon vs Rom Split'  # Keep only one title argument
    )

    pgrp_line = base.mark_line(point=alt.OverlayMarkDef(size=150)).encode(
        color=alt.Color('PubGroup', legend=None),
        x=alt.X('year', title='Year', axis=alt.Axis(title='', labelAngle=0)),
        y=alt.Y('sum(rev):Q', title="Total Sales"),
        tooltip=['PubGroup', 'year', alt.Tooltip('sum(rev):Q', format="$,.0f", title='Total Sales')],
        text=alt.X('sum(rev)')
    ).properties(
        title='YTD vs Prior Years (Thru Full-Month) Pub Group Comparisons',
        height=height * .6,
        width=750
    ).transform_filter(
        alt.FieldOneOfPredicate(field='month', oneOf=months_ty)
    )

    plot_title = alt.TitleParams(
        'Publisher and Pub Group Summary',
        subtitle=['All Data Represents Core Sales', ''],
        align='center',
        anchor='middle',
        fontSize=30
    )

    pub_options = list(df_daily['Publisher'].unique())
    item = 'Chronicle'
    pub_options.remove(item)
    pub_options.insert(0, item)

    publisher_dropdown = alt.binding_select(options=[None] + pub_options, labels=['All'] + pub_options, name='Publisher: ')
    publisher_select = alt.selection_point(fields=['Publisher'], bind=publisher_dropdown)

    channel_dropdown = alt.binding_select(options=[None] + channel_list, labels=['All'] + channel_list, name='Channels')
    channel_select = alt.selection_point(fields=['channel'], bind=channel_dropdown)

    cumulative_chart = ((base_ty + base_line_ly + total_sales)
                        & (year_bar | monthly_bar)
                        & (channel_bg_bar_month + channel_bg_tick_month)
                        & (channel_bg_bar_ytd + channel_bg_tick_ytd)
                        & pgrp_line & (heat_month)
                        & (heat_quarter | ytd_publisher | year_publisher) & amaz_rom).add_params(
        publisher_select).transform_filter(
        publisher_select).properties(title=plot_title)

    return cumulative_chart

def main():
    df_raw = load_data()
    df_daily, ty, ly, months_ty, tm = prepare_data(df_raw)
    cumulative_chart = create_charts(df_daily, ty, ly, months_ty, tm)
    
    path = saved_viz_location()
    cumulative_chart.save(path)
    
    # Save to the current folder as well
    current_folder_path = 'ssr_summary_chart_v2.html'
    cumulative_chart.save(current_folder_path)

    print(f"Charts saved to {path}\nand\n{current_folder_path}")
    
if __name__ == "__main__":
    main()