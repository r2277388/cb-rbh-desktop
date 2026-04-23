import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio

# Set the default renderer once
pio.renderers.default = 'browser'

def load_and_prepare_data(file_path):
    """
    Loads data from a CSV file and prepares it for visualization.
    
    Parameters:
        file_path (str): The path to the CSV file.
        
    Returns:
        pd.DataFrame: Processed DataFrame ready for plotting.
    """
    # Load data
    df = pd.read_csv(file_path)
    
    # Convert 'Date' to datetime and extract month and year
    df['Date'] = pd.to_datetime(df['Date'])
    df['Month'] = df['Date'].dt.strftime('%b')  # Get abbreviated month names
    df['Year'] = df['Date'].dt.year
    
    # Filter for years of interest
    years_of_interest = [2023, 2024, 2025]
    df_filtered = df[df['Year'].isin(years_of_interest)]
    
    # Drop non-numeric columns before grouping
    df_filtered = df_filtered[['Month', 'Year', 'est']]
    
    # Group by 'Month' and 'Year' and sum 'est' values
    df_grouped = df_filtered.groupby(['Year', 'Month'], as_index=False).sum()
    
    # Ensure months are in correct order
    month_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    df_grouped['Month'] = pd.Categorical(df_grouped['Month'], categories=month_order, ordered=True)
    df_grouped = df_grouped.sort_values(['Year', 'Month'])
    
    return df_grouped

def create_comparative_plot(df_grouped):
    """
    Creates a comparative plot with 2024 as a bar chart and 2023 & 2025 as line charts.
    
    Parameters:
        df_grouped (pd.DataFrame): The processed DataFrame containing 'Year', 'Month', and 'est'.
        
    Returns:
        plotly.graph_objects.Figure: The generated plotly figure.
    """
    # Define color palette
    colors = {
        2023: '#1f77b4',  # Blue
        2024: '#ff7f0e',  # Orange
        2025: '#2ca02c'   # Green
    }
    
    # Initialize figure
    fig = go.Figure()
    
    # Add bar chart for 2024
    df_2024 = df_grouped[df_grouped['Year'] == 2024]
    fig.add_trace(
        go.Bar(
            x=df_2024['Month'],
            y=df_2024['est'],
            name='2024',
            marker_color=colors[2024],
            text=df_2024['est'].apply(lambda x: f"${x/1e6:.1f}M"),  # Format as $X.XM
            textposition='auto'
        )
    )
    
    # Add line charts for 2023 and 2025
    for year in [2023, 2025]:
        df_year = df_grouped[df_grouped['Year'] == year]
        if not df_year.empty:
            fig.add_trace(
                go.Scatter(
                    x=df_year['Month'],
                    y=df_year['est'],
                    mode='lines+markers',
                    name=str(year),
                    marker=dict(size=8),
                    line=dict(width=2),
                    marker_color=colors[year]
                )
            )
    
    # Update layout
    fig.update_layout(
        title='Monthly Estimated Values: 2024 vs 2023 & 2025',
        xaxis=dict(
            title='Month',
            tickmode='linear'
        ),
        yaxis=dict(
            title='Estimated Value',
            tickformat=',',
            hoverformat=',.2f'
        ),
        legend=dict(
            title='Year',
            orientation='h',
            x=0.5,
            xanchor='center',
            y=-0.2
        ),
        bargap=0.2,
        template='plotly_white',
        width=1000,
        height=600
    )
    
    # Add hover template
    fig.update_traces(hovertemplate='Month: %{x}<br>Est: %{y:$,.2f}')
    
    return fig

if __name__ == '__main__':
    file_path = 'pgrp_flbl_forecasts.csv'
    df_grouped = load_and_prepare_data(file_path)
    fig = create_comparative_plot(df_grouped)
    fig.show()