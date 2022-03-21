
#Dash components
import dash
import dash_table
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc

#Plotly
import plotly.express as px
import plotly.graph_objects as go
#File selection prompt
#import tkinter
#from tkinter import filedialog

import pandas as pd
import numpy as np
import sqlite3
from sqlite3.dbapi2 import DatabaseError
import os



SIDEBAR_STYLE = {
    "position": "fixed",
    "top": 0,
    "left": 0,
    "bottom": 0,
    "width": "18rem",
    "padding": "2rem 1rem",
    "background-color": "#000000",
}
CONTENT_STYLE = {
    "margin-left": "18rem",
    "margin-right": "2rem",
    "padding": "2rem 1rem",
}

PASCODES = ['AU4WFCMW','AU4WF18H','AU4WF1FH','AU4WFDWG','AU4WFNL3','AU4WFNL4','AU4WFNL5','AU4WFNL6','AU4WFR3Y','BP4WFDWC','HH4WFD8Z']

#wwa = "World Wide Averages"
dbname = 'MedicalDashboardDB'
conn   = sqlite3.connect(dbname + '.sqlite')
cur = conn.cursor() #This code creates a cursor which will be used to pull data from our database.

#TODO Change this to name files
xlFile = pd.ExcelFile('WWA_Apr 2021.xlsx')
xlFile2 = pd.ExcelFile('UPMR ao 01 Nov 2.xlsx')
xlFile3 = pd.ExcelFile('316 MDG UMD Jul 21 Excel 97-2003.xls')
xlFile4 = pd.ExcelFile('Gains Listing ao 01 Nov 21.xlsx')
xlFile5 = pd.ExcelFile('Loss Listing ao 01 Nov 21.xlsx')


upmr_sheet = pd.read_excel(xlFile2)
#Filtered out PAScodes we do not care about
upmr_sheet = upmr_sheet[pd.DataFrame(upmr_sheet.PASCODE.tolist()).isin(PASCODES).any(1).values]
umd_sheet = pd.read_excel(xlFile3, ['EXCEL'])
umd_excel_sheet = umd_sheet['EXCEL']
gains_sheet = xlFile4.parse()
loss_sheet = xlFile5.parse()
#Changing column names to values on row 1 in the spreadsheet
gains_sheet.columns = gains_sheet.iloc[1]
loss_sheet.columns = loss_sheet.iloc[1]
umd_excel_afsc = umd_excel_sheet['AFSC']
#Filtering 'UMD' sheet, 'AFSC' column for AFSCs starting with a '0'
umd_excel_sheet.loc[umd_excel_sheet['AFSC'].str.startswith('0'), 'AFSC'] = umd_excel_sheet.loc[umd_excel_sheet['AFSC'].str.startswith('0'), 'AFSC'].str.lstrip('0')

num_auth_upmr = upmr_sheet['AFSC_AUTH'].value_counts()
num_pascode_upmr = upmr_sheet['PASCODE'].value_counts()
num_prim_upmr = upmr_sheet['DAFSC'].value_counts()
num_afsc_umd = umd_excel_sheet['AFSC'].value_counts()

total_auth = len(umd_excel_afsc.index)
total_ass = len(upmr_sheet.dropna(subset=['NAME']))



file_names = ['WWA_Apr 2021.xlsx','Copy of 316 MDSS BLSDM Support Roster 1 Sep 21_Copy.xlsx','316 MDG UMD Jul 21 Excel 97-2003.xls']

sheet = []
for sheetName in xlFile.sheet_names:
    sheet.append(xlFile.parse(sheetName))




#
#Begin Code
#
app=dash.Dash(__name__, external_stylesheets=[dbc.themes.DARKLY], suppress_callback_exceptions=True)
sidebar = html.Div(
    [
        html.H3("316 Medical Group", className="display-4"),
        html.Hr(),
        html.P(
            "Click to view Individual pages:", className="lead"
        ),
        dbc.Nav(
            [
                dbc.NavLink("Home", href="/", active="exact"),
                dbc.NavLink("WWA", href="/WWA", active="exact"),
                dbc.NavLink("Assigned vs Auth", href="/AvA", active="exact"),
                dbc.NavLink("FY Proj", href="/FY_Projections", active="exact"),
            ],
            vertical=True,
            pills=True,
        ),
    ],
    style=SIDEBAR_STYLE,
)
##############################################################
#Home Page content
#
##############################################################
home_page = html.Div([
    dbc.Row([
        dbc.Col(
            dbc.Card(
                dbc.CardBody('Total Assigned Personnel: {}'.format(total_ass)),
                body=True, 
                id='ass-card',
                color="grey"),
        className="mb-4",
        ),

        dbc.Col(
            dbc.Card(
                dbc.CardBody('Total Authorized Personnel: {}'.format(total_auth)),
                body=True, 
                id='auth-card',
                color="grey"),
        className="mb-4"
        )  
    ]),

##Displays all sheet names in excel spreadsheet

    dcc.Dropdown(
        id="dropdown",
        options=[{"label": x, "value": x} for x in xlFile.sheet_names],
        value=xlFile.sheet_names[0],
        clearable=False,
        style={
            'color': '#000000',
            'background-color': '#cdcdcd',
        } 
    ),  

    html.Div(
        id="afscbar",
        style={},
        children= dcc.Graph(id="bar-chart")
    ),
    html.Div(
        id="tablediv",
        children= dash_table.DataTable(
            id="afsctable",
            columns=[{"name": i, "id": i} for i in pd.DataFrame(num_afsc_umd).columns],
            data=pd.DataFrame(num_afsc_umd).to_dict('records'),
            page_size=10,
            style_header={
                'backgroundColor': 'rgb(30, 30, 30)',
                'color': 'white'
            },
            style_data={
                'backgroundColor': 'rgb(50, 50, 50)',
                'color': 'white'
            },
        ),
    )
], style=CONTENT_STYLE)

##############################################################
#Page_1(WWA)content
#
##############################################################

wwa_page = html.Div([
    dbc.Row([
        dbc.Col(
            dbc.Card(
                html.Div(
                    children=[html.Ul(children=[html.Li(i) for i in file_names])],
                    className="text-left text-light bg-grey"),
            body=True, 
            color="grey"),
        className="mb-4"
        )
    ],style={'width': '40%'}),

##Displays all sheet names in excel spreadsheet

    dcc.Dropdown(
        id="dropdown",
        options=[{"label": x, "value": x} for x in xlFile.sheet_names],
        value=xlFile.sheet_names[0],
        clearable=False,
        style={
            'color': '#000000',
            'background-color': '#cdcdcd',
        } 
    ),  

    html.Div(
        id="afscbar",
        style={},
        children= dcc.Graph(id="bar-chart")
    ),
], style=CONTENT_STYLE)

##############################################################
#Assigned vs Authorized Page content
#
##############################################################
ava_page = html.Div([
    dbc.Row([
        dbc.Col(
            dbc.Card(
                html.Div(
                    children=[html.Ul(children=[html.Li(i) for i in file_names])],
                    className="text-left text-light bg-grey"),
                body=True, 
                color="grey"
            ),
            className="mb-4"
        )
    ],style={'width': '40%'}),

##Displays all sheet names in excel spreadsheet
    dbc.Row([
        dbc.Col(
            html.Div(
                dcc.Dropdown(
                    id="dropdown-afsc",
                    options=[{"label": x, "value": x} for x in num_prim_upmr.index],
                    clearable=False,
                    style={
                        'color': '#000000',
                        'background-color': '#cdcdcd',
                    } 
                )
            ),
        ), 
        dbc.Col(
            html.Div(
                dcc.Dropdown(
                    id="dropdown-sqdrn",
                    options=[{"label": x, "value": x} for x in num_pascode_upmr.index],
                    clearable=False,
                    style={
                        'color': '#000000',
                        'background-color': '#cdcdcd',
                    } 
                )
            )
        )]
    ),  

    html.Div(
        dbc.Row(
            children=[
                dbc.Col(
                    id="afscbar",
                    style={},
                    children = dcc.Graph(id="auth-bar")
                ),
                dbc.Col(
                    id="afscbar",
                    style={},
                    children = dcc.Graph(id="ass-bar")
                )
            ]
        ),
    ),
], style=CONTENT_STYLE)


##############################################################
#Fiscal Year Projections Page content
#
##############################################################
fy_proj_page = html.Div([
    dbc.Row([
        dbc.Col(
            dbc.Card(
                html.Div(
                    children=[html.Ul(children=[html.Li(i) for i in file_names])],
                    className="text-left text-light bg-grey"),
            body=True, 
            color="grey"),
        className="mb-4"
        )
    ],style={'width': '40%'}),

##Displays all sheet names in excel spreadsheet
    dcc.Dropdown(
        id="dropdown_FY",
        options=[{"label": x, "value": x} for x in num_pascode_upmr.index],
        value=xlFile.sheet_names[0],
        clearable=False,
        style={
            'color': '#000000',
            'background-color': '#cdcdcd',
        } 
    ),
    html.Div(
        id="tablediv",
        children= dash_table.DataTable(
            id="gains_table",
            columns=[{"name": i, "id": i} for i in gains_sheet.columns],
            data=gains_sheet.to_dict('records'),
            page_size=10,
            style_table={'overflowX': 'scroll'},
            style_header={
                'backgroundColor': 'rgb(30, 30, 30)',
                'color': 'white'
            },
            style_data={
                'backgroundColor': 'rgb(50, 50, 50)',
                'color': 'white'
            },
        ),
    ),
    html.Div(
        id="tablediv",
        children= dash_table.DataTable(
            id="gains_table",
            columns=[{"name": i, "id": i} for i in gains_sheet.columns],
            data=gains_sheet.to_dict('records'),
            page_size=10,
            style_table={'overflowX': 'scroll'},
            style_header={
                'backgroundColor': 'rgb(30, 30, 30)',
                'color': 'white'
            },
            style_data={
                'backgroundColor': 'rgb(50, 50, 50)',
                'color': 'white'
            },
        ),
    ),  
], style=CONTENT_STYLE)

app.layout = html.Div([dcc.Location(id="url"), sidebar, html.Div(children=home_page, id='page-content')])



#This function updates the bar-graph when the dropdown menu is used
#Input:dropdown value for sheet name
#Output:Bar graph of <sheet names> AFSC vs Authorized vs Assigned
@app.callback(
    [Output("ass-bar", "figure"),
    Output("auth-bar", "figure")],
    [Input("dropdown-afsc", "value"),#Dropdown values from UPMR
    Input("dropdown-sqdrn", "value")],prevent_initial_call=True)

def update_ava_chart(afsc,pas):
    figure1 = px.bar(umd_excel_sheet, x=num_afsc_umd.index, y=[num_afsc_umd],
                height=1000)
    figure1.update_layout(title_text='Authorized AFSCs', title_x=0.5)

    if afsc==None and pas != None:
        upmr_pas_filtered = upmr_sheet.loc[upmr_sheet['PASCODE'] == pas]
        upmr_pas_dafsc = upmr_pas_filtered['DAFSC'].value_counts()

        figure2 = go.Figure(data=[
            go.Bar(name='Authorized', x='''list of afscs at pascode''', y=[num_auth_upmr[afsc]]),
            go.Bar(name='Assigned', x='''list of afscs at pascode in upmr''', y=[upmr_pas_dafsc[afsc]])
        ])
    elif afsc!=None and pas == None:
        figure2 = go.Figure(data=[
            go.Bar(name='Authorized', x=[afsc], y=[num_afsc_umd[afsc]]),
            go.Bar(name='Assigned', x=[afsc], y=[num_prim_upmr[afsc]])
        ])
    else:
        upmr_pas_filtered = upmr_sheet.loc[upmr_sheet['PASCODE'] == pas]
        upmr_pas_dafsc = upmr_pas_filtered['DAFSC'].value_counts()
        figure2 = go.Figure(data=[
            go.Bar(name='Authorized', x=[afsc], y=[num_auth_upmr[afsc]]),
            go.Bar(name='Assigned', x=[afsc], y=[upmr_pas_dafsc[afsc]])
        ])
    figure2.update_layout(barmode='group')
    #figure2 = px.bar(num_auth_upmr, x=[code, code], y=[num_auth_upmr[code],num_prim_upmr[code]],
    #            barmode = "group", height=1000)
    #figure2.update_layout(title_text='Authorized AFSC', title_x=0.5)

    return figure1, figure2

#This function updates the card at the top showing the files being used as data.
#Input: url, pathname
#Output: latest files being used
#TODO

#This fuction updates the FY Proj graph to display loss list and gaining
#Input
#Output
@app.callback(
    Output("gains_table", 'data'),
    [Input("dropdown_FY", "value")]
)
def update_fy_table(code):
    return gains_sheet.loc[gains_sheet['GAINING_PAS'] == code].to_dict('records')


#This function updates the bar-graph in the WWA page when the dropdown menu is used
#Input:dropdown value for sheet name
#Output:Bar graph of <sheet names> AFSC vs Authorized vs Assigned
@app.callback(
    Output("bar-chart", "figure"),
    [Input("dropdown", "value")])

def update_wwa_chart(shtName):
    figure = px.bar(xlFile.parse(shtName), x="AFSC", y=["ASGN", "AUTH"],
                 barmode="group", height=1000)
    figure.update_layout(title_text='World Wide Averages', title_x=0.5)
    return figure



#This function updates the page with content related to the tab clicked
#Input: url change
#Output: updated page content
@app.callback(
    Output('page-content', 'children'),
    [Input('url', 'pathname')])

def display_page(pathname):
    if pathname == '/WWA':
        return wwa_page
    elif pathname == '/AvA':
        return ava_page
    elif pathname == '/FY_Projections':
        return fy_proj_page
    else:
        return home_page


#This function updates the input file references used to form the data
#Input: submit button
#Output: File path 
#State: filepath
'''
@app.callback(
    Output(),
    [Input()],
    State()
    )

def file_selector():
    tkinter.Tk().withdraw()
    upmr_filepath = filedialog.askopenfilename()
'''
if __name__ == '__main__':
    app.run_server(debug=True,port=8001)


