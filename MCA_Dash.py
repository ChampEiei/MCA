from dash import Dash 
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output
import plotly.express as px
import pandas as pd

# Load and prepare data
df_all=pd.read_excel("https://github.com/ChampEiei/MCA/main/%E0%B8%AA%E0%B8%A1%E0%B8%9A%E0%B8%B9%E0%B8%A3%E0%B8%93%E0%B9%8C.xlsx")
df_melted=pd.read_excel("https://github.com/ChampEiei/MCA/main/melted.xlsx")
df_all = df_all.dropna(subset=['P&L Type'])
df_melted = df_melted.dropna(subset=['P&L Type'])
pl_type_options = [{'label': pl_type, 'value': pl_type} for pl_type in df_all['P&L Type'].unique()]

# Load the cost structure data
cost = pd.read_excel("https://github.com/ChampEiei/MCA/main/Cost_structure.xlsx")
cost = pd.melt(cost, id_vars='P&L_types', var_name='Expense_type', value_name='Expense_amount')

# Prepare options for cost structure filter
cost_pl_type_options = [{'label': pl_type, 'value': pl_type} for pl_type in cost['P&L_types'].unique()]

# Define color map
color_map = {
    'บริการจัดกิจกรรมแสดงสินค้าและดิจิทัล': '#00CADC',
    'บริการบรรจุและจัดส่งสินค้า': '#7E80E7',
    'บริการพนักงานแนะนำสินค้า': '#9B57CC',
    'บริการจัดเรียงสินค้า': '#65A6FA'
}

# Initialize the Dash app
app = Dash(__name__)
server = app.server

# Define the layout of the app
app.layout = html.Div(children=[
    html.H1(children='MCA-Interactive Multi-Graph Dashboard', style={'textAlign': 'center', 'color': '#003366'}),

    html.Div(children=''' Margin by Sector and Customer. ''', style={'textAlign': 'center', 'color': '#003366'}),

    # Dropdown for filtering by P&L Type
    html.Div(
        children=[
            dcc.Dropdown(
                id='pl-type-filter',
                options=pl_type_options,
                value=[pl_type['value'] for pl_type in pl_type_options],
                multi=True,
                style={'width': '60%', 'margin': 'auto'}
            )
        ],
        style={'padding': '10px'}
    ),

    # DatePicker for filtering by date range
    html.Div(
        children=[
            dcc.DatePickerRange(
                id='date-picker-range',
                start_date=df_all['Start Date'].min(),
                end_date=df_all['Start Date'].max(),
                display_format='YYYY-MM-DD',
                style={'width': '60%', 'margin': 'auto'}
            )
        ],
        style={'padding': '10px'}
    ),

    # Container for the graphs
    html.Div(
        children=[
            html.Div(children=[dcc.Graph(id='scatter-graph')], style={'width': '48%', 'display': 'inline-block', 'padding': '10px'}),
            html.Div(children=[dcc.Graph(id='pie-chart')], style={'width': '48%', 'display': 'inline-block', 'padding': '10px'}),
        ],
        style={'display': 'flex', 'justify-content': 'space-between', 'flex-wrap': 'wrap', 'padding': '10px'}
    ),
    html.Div(
        children=[
            html.Div(children=[dcc.Graph(id='bar-chart')], style={'width': '48%', 'display': 'inline-block', 'padding': '10px'}),
            html.Div(children=[dcc.Graph(id='monthly-bar-chart')], style={'width': '48%', 'display': 'inline-block', 'padding': '10px'}),
        ],
        style={'display': 'flex', 'justify-content': 'space-between', 'flex-wrap': 'wrap', 'padding': '10px'}
    ),
    # Add new graph for Project Type vs. margin
    html.Div(
        children=[
            html.Div(children=[dcc.Graph(id='project-type-bar-chart')], style={'width': '100%', 'padding': '10px'}),
        ],
        style={'padding': '10px'}
    ),
    # Add new graph for Margin Over Time
    html.Div(
        children=[
            html.Div(children=[dcc.Graph(id='margin-over-time-graph')], style={'width': '100%', 'padding': '10px'}),
        ],
        style={'padding': '10px'}
    ),

    # New Section: Cost Structure Analysis 2023
    html.H2(children='Cost Structure Analysis 2023', style={'textAlign': 'center', 'color': '#003366', 'padding-top': '20px'}),

    # Dropdown specifically for the cost structure graph
    html.Div(
        children=[
            dcc.Dropdown(
                id='cost-pl-type-filter',
                options=cost_pl_type_options,
                value=[pl_type['value'] for pl_type in cost_pl_type_options],
                multi=True,
                style={'width': '50%', 'margin': 'auto'}
            )
        ],
        style={'padding': '10px', 'textAlign': 'center'}
    ),

    # Graphs for Cost Structure Analysis
    html.Div(
        children=[
            html.Div(children=[dcc.Graph(id='cost-bar-chart')], style={'width': '48%', 'display': 'inline-block', 'padding': '10px'}),
            html.Div(children=[dcc.Graph(id='cost-sunburst-chart')], style={'width': '48%', 'display': 'inline-block', 'padding': '10px'}),  # New Sunburst chart
        ],
        style={'display': 'flex', 'justify-content': 'space-between', 'flex-wrap': 'wrap', 'padding': '10px'}
    ),

    # New table for showing group_mar data
    html.H2(children='Aggregated Margin and Revenue Data', style={'textAlign': 'center', 'color': '#003366', 'padding-top': '20px'}),
    html.Div(
        children=[
            dash_table.DataTable(
                id='aggregated-table',
                columns=[
                    {'name': 'P&L Type', 'id': 'P&L Type'},
                    {'name': 'Margin', 'id': 'margin'},
                    {'name': 'Total Revenue In year', 'id': 'Total Revenue In year'},
                    {'name': 'margin ratio', 'id': 'margin ratio'}
                ],
                data=[],
                page_size=10,
                style_table={'overflowX': 'auto', 'margin': 'auto'},
                style_cell={'textAlign': 'left', 'padding': '5px'},
                style_header={'backgroundColor': 'lightblue', 'fontWeight': 'bold'}
            )
        ],
        style={'width': '80%', 'margin': 'auto', 'padding': '20px'}
    )
], style={'backgroundColor': '#E6F0FF', 'font-family': 'Arial'})

# Define the callbacks to update graphs based on filter
@app.callback(
    [Output('scatter-graph', 'figure'),
     Output('pie-chart', 'figure'),
     Output('bar-chart', 'figure'),
     Output('monthly-bar-chart', 'figure'),
     Output('project-type-bar-chart', 'figure'),
     Output('margin-over-time-graph', 'figure'),
     Output('cost-bar-chart', 'figure'),
     Output('cost-sunburst-chart', 'figure'),  # New output for sunburst chart
     Output('aggregated-table', 'data')],
    [Input('pl-type-filter', 'value'),
     Input('date-picker-range', 'start_date'),
     Input('date-picker-range', 'end_date'),
     Input('cost-pl-type-filter', 'value')]
)
def update_graphs(selected_pl_types, start_date, end_date, selected_cost_pl_types):
    if selected_pl_types:
        filtered_df_all = df_all[df_all['P&L Type'].isin(selected_pl_types)]
        filtered_df_melted = df_melted[df_melted['P&L Type'].isin(selected_pl_types)]
    else:
        filtered_df_all = df_all
        filtered_df_melted = df_melted

    # Filter by date range
    if start_date and end_date:
        filtered_df_all = filtered_df_all[(filtered_df_all['Start Date'] >= start_date) & (filtered_df_all['Start Date'] <= end_date)]
        filtered_df_melted = filtered_df_melted[(filtered_df_melted['YearMonth'] >= start_date[:7]) & (filtered_df_melted['YearMonth'] <= end_date[:7])]

    # Ensure margin values are numeric and non-negative
    filtered_df_all = filtered_df_all[pd.to_numeric(filtered_df_all['margin'], errors='coerce').notnull()]
    filtered_df_all = filtered_df_all[filtered_df_all['margin'] > 0]

    # Update scatter plot to use margin for size
    scatter_fig = px.scatter(
        filtered_df_all,
        x='Start Date',
        y='margin',
        color='P&L Type',
        size='margin',
        size_max=30,
        color_discrete_map=color_map
    )
    scatter_fig.update_layout(title='Scatter Plot of Margin Over Time', xaxis_title='Start Date', yaxis_title='Margin', showlegend=False)

    # Update pie chart to use margin
    pie_fig = px.pie(filtered_df_all, names='P&L Type', values='margin', title='Margin Distribution by P&L Type',color='P&L Type',color_discrete_map=color_map)

    # Update bar chart to show top 5 clients by margin
    df_sorted = filtered_df_all.groupby(['Client', 'P&L Type'])['margin'].sum().reset_index()
    df_sorted = df_sorted.sort_values(by='margin', ascending=False)
    top_5_clients = df_sorted.groupby('Client')['margin'].sum().nlargest(5).index
    df_top_5 = df_sorted[df_sorted['Client'].isin(top_5_clients)]

    bar_fig = px.bar(df_top_5, x='Client', y='margin', color='P&L Type',color_discrete_map=color_map)
    bar_fig.update_layout(title='Top 5 Clients by Margin')

    monthly_bar_fig = px.bar(filtered_df_melted, x='YearMonth', y='Amount', color='Metric', barmode='group',
                             labels={'Amount': 'Amount', 'YearMonth': 'Year and Month'},
                             title='Monthly Revenue and Cost Over Time')

    # Create new bar chart for Project Type vs. margin
    df_Project_Type = filtered_df_all.groupby(['Project Type', 'P&L Type'])['margin'].sum().reset_index()
    df_Project_Type = df_Project_Type.sort_values(by='margin', ascending=False)
    fig_Project_Type = px.bar(df_Project_Type, x='P&L Type', y='margin', color='Project Type', title='Margin by Project Type and P&L Type', barmode='group')

    # Apply the color map for the Project Type graph
    fig_Project_Type.for_each_trace(lambda t: t.update(marker_color=color_map.get(t.name, t.marker.color)))

    # Create the Margin Over Time graph
    margin = filtered_df_all.groupby(filtered_df_all['Start Date'].dt.to_period('M'))['margin'].sum().reset_index()
    margin['Start Date'] = margin['Start Date'].dt.strftime('%Y-%m')
    margin['Start Date'] = margin['Start Date'].astype('str') + '-01'
    margin['Start Date'] = pd.to_datetime(margin['Start Date'])
    margin = margin[margin['margin'] != 0]
    margin.reset_index(drop=True, inplace=True)
    fig_margin = px.scatter(margin, x='Start Date', y='margin', title='Margin Over Time', trendline='ols')
    fig_margin.data[1].update(line=dict(color='red', width=2))
    fig_margin.update_layout(title='Margin Over Time', xaxis_title='Month')
    fig_margin.update_xaxes(range=['2021-01-01', '2023-12-31'])

    # Update cost bar chart based on selected P&L Types for cost
    if selected_cost_pl_types:
        filtered_cost = cost[cost['P&L_types'].isin(selected_cost_pl_types)]
    else:
        filtered_cost = cost

    # Sort the filtered_cost DataFrame by 'Expense_amount' in descending order
    filtered_cost = filtered_cost.sort_values(by='Expense_amount', ascending=False)

    fig_cost = px.bar(
        filtered_cost, 
        x='P&L_types', 
        y='Expense_amount', 
        barmode='group', 
        color='Expense_type',
        title='Cost Structure by P&L Type and Expense Type'
    )

    # Sunburst chart for cost structure analysis
    sunburst_fig = px.sunburst(
        filtered_cost,
        path=['P&L_types', 'Expense_type'],
        values='Expense_amount',
        title='Sunburst Chart of Cost Structure',
        color='P&L_types',
        color_discrete_map={'จัดกิจกรรมทางการตลาดและดิจิทัล': '#00CADC',
 'บรรจุและจัดส่งสินค้า': '#7E80E7',
 'พนักงานแนะนำสินค้า': '#9B57CC',
 'จัดเรียงสินค้า': '#65A6FA'}
        
    )

    # Create group_mar DataFrame
    group_mar = filtered_df_all.groupby(['P&L Type'])[['margin', 'Total Revenue In year']].sum().reset_index()
    group_mar['margin ratio'] = group_mar['margin'] / group_mar['Total Revenue In year']
    group_mar = group_mar[group_mar['P&L Type'] != 'Others']
    # Update table data
    table_data = group_mar.to_dict('records')

    return scatter_fig, pie_fig, bar_fig, monthly_bar_fig, fig_Project_Type, fig_margin, fig_cost, sunburst_fig, table_data



if __name__ == '__main__':
    app.run_server(debug=True)
