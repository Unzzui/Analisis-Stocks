import os
import requests
import dash
import dash_bootstrap_components as dbc
from dash import dcc, html
from dash.dependencies import Input, Output
import pandas as pd

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.LUX])

def download_excel_file(symbol, section):
    url = f"https://stockrow.com/api/companies/{symbol}/financials.xlsx?dimension=A&section={section}&sort=asc"

    response = requests.get(url)
    if response.status_code == 200:
        file_path = f"data/{symbol}_financials_{section.replace(' ', '_')}.xlsx"
        with open(file_path, "wb") as file:
            file.write(response.content)
        return file_path
    else:
        return None

def open_excel_files(metrics_path, balance_sheet_path, income_statement_path):
    df_metrics = pd.read_excel(metrics_path)
    df_balance = pd.read_excel(balance_sheet_path)
    df_income = pd.read_excel(income_statement_path)
    
    return df_metrics, df_balance, df_income

@app.callback(
    Output('output-container', 'children'),
    [Input('analyze-button', 'n_clicks')],
    [dash.dependencies.State('symbol-input', 'value')]
)
def update_output(n_clicks, symbol):
    if n_clicks is None:
        return []
    
    # Verificar si el archivo de métricas ya está descargado
    metrics_path = f"data/{symbol}_financials_metrics.xlsx"
    if not os.path.exists(metrics_path):
        metrics_path = download_excel_file(symbol, "Metrics")
        if metrics_path is None:
            return html.Div("Error al descargar el archivo de métricas.")
    
    # Verificar si el archivo del estado de resultados ya está descargado
    income_statement_path = f"data/{symbol}_financials_income_statement.xlsx"
    if not os.path.exists(income_statement_path):
        income_statement_path = download_excel_file(symbol, "Income Statement")
        if income_statement_path is None:
            return html.Div("Error al descargar el archivo del estado de resultados.")
    
    # Verificar si el archivo del balance ya está descargado
    balance_sheet_path = f"data/{symbol}_financials_balance_sheet.xlsx"
    if not os.path.exists(balance_sheet_path):
        balance_sheet_path = download_excel_file(symbol, "Balance Sheet")
        if balance_sheet_path is None:
            return html.Div("Error al descargar el archivo del balance.")
    
    df_metrics, df_balance, df_income = open_excel_files(metrics_path, balance_sheet_path, income_statement_path)
    
    return generate_layout(df_metrics, df_balance, df_income)


def generate_layout(df_metrics, df_balance, df_income):
    try:
        metrics_table = html.Table(
            [html.Tr([html.Th(col) for col in df_metrics.columns])] +
            [html.Tr([html.Td(df_metrics.iloc[i][col]) for col in df_metrics.columns]) for i in range(len(df_metrics))]
        )
        
        balance_table = html.Table(
            [html.Tr([html.Th(col) for col in df_balance.columns])] +
            [html.Tr([html.Td(df_balance.iloc[i][col]) for col in df_balance.columns]) for i in range(len(df_balance))]
        )
        
        income_table = html.Table(
            [html.Tr([html.Th(col) for col in df_income.columns])] +
            [html.Tr([html.Td(df_income.iloc[i][col]) for col in df_income.columns]) for i in range(len(df_income))]
        )
        
        return html.Div(
            children=[
                html.H1("Análisis Financiero", className="text-center mt-5 mb-4"),
        
                html.H2("Métricas", className="mt-5 mb-3"),
                metrics_table,

                html.H2("Balance Sheet", className="mt-5 mb-3"),
                balance_table,

                html.H2("Income Statement", className="mt-5 mb-3"),
                income_table
            ]
        )
    except KeyError as e:
        return html.Div(f"Error al generar el layout: {str(e)}")


app.layout = html.Div(
    children=[
        html.H2("Análisis Financiero", className="text-center mt-5 mb-4"),
        html.Div(
            children=[
                dbc.Input(id='symbol-input', placeholder='Símbolo', className="mr-2"),
                dbc.Button("Analyze", id='analyze-button', color="primary", outline=True),
            ],
            className="d-flex justify-content-center mt-3"
        ),
        html.Div(id='output-container', className="mt-4")
    ],
    className="container"
)


if __name__ == '__main__':
    app.run_server(debug=True)
