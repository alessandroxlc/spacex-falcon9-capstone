"""Dashboard interactivo de lanzamientos SpaceX con Plotly Dash.

Ejecución:
    python dashboard/spacex_dash_app.py
y abre http://127.0.0.1:8050 en el navegador.

Componentes (los mismos que pide el laboratorio de IBM):
    1. Desplegable para elegir todos los sitios o uno concreto.
    2. Gráfico de pastel: éxitos por sitio (todos) o éxito vs. fallo (un sitio).
    3. Control deslizante de rango para la masa de carga útil.
    4. Dispersión masa de carga vs. resultado, coloreada por versión del booster.
"""
from pathlib import Path

import pandas as pd
import plotly.express as px
from dash import Dash, Input, Output, dcc, html

ROOT = Path(__file__).resolve().parent.parent
LOCAL_CSV = ROOT / "data" / "spacex_launch_dash.csv"
REMOTE_CSV = (
    "https://cf-courses-data.s3.us.cloud-object-storage.appdomain.cloud/"
    "IBM-DS0321EN-SkillsNetwork/datasets/spacex_launch_dash.csv"
)


def load_data() -> pd.DataFrame:
    if LOCAL_CSV.exists():
        return pd.read_csv(LOCAL_CSV)
    df = pd.read_csv(REMOTE_CSV)
    LOCAL_CSV.parent.mkdir(exist_ok=True)
    df.to_csv(LOCAL_CSV, index=False)
    return df


spacex_df = load_data()
max_payload = spacex_df["Payload Mass (kg)"].max()
min_payload = spacex_df["Payload Mass (kg)"].min()
sites = sorted(spacex_df["Launch Site"].unique())

OUTCOME_COLORS = {"Éxito": "#1FA67A", "Fallo": "#D64545"}

app = Dash(__name__, title="SpaceX Launch Records Dashboard")
server = app.server

app.layout = html.Div(
    style={"fontFamily": "Arial, sans-serif", "maxWidth": "1200px", "margin": "0 auto", "padding": "16px"},
    children=[
        html.H1(
            "SpaceX Launch Records Dashboard",
            style={"textAlign": "center", "color": "#0B1220", "fontSize": 36},
        ),
        # TAREA 1: desplegable de sitios de lanzamiento
        dcc.Dropdown(
            id="site-dropdown",
            options=[{"label": "Todos los sitios", "value": "ALL"}]
            + [{"label": site, "value": site} for site in sites],
            value="ALL",
            placeholder="Selecciona un sitio de lanzamiento",
            searchable=True,
            clearable=False,
        ),
        html.Br(),
        # TAREA 2: gráfico de pastel
        dcc.Graph(id="success-pie-chart"),
        html.Br(),
        html.P("Rango de masa de carga útil (kg):"),
        # TAREA 3: control deslizante de rango
        dcc.RangeSlider(
            id="payload-slider",
            min=0,
            max=10000,
            step=1000,
            marks={i: f"{i:,}".replace(",", ".") for i in range(0, 10001, 2500)},
            value=[min_payload, max_payload],
        ),
        # TAREA 4: dispersión masa de carga vs. resultado
        dcc.Graph(id="success-payload-scatter-chart"),
    ],
)


@app.callback(Output("success-pie-chart", "figure"), Input("site-dropdown", "value"))
def get_pie_chart(entered_site: str):
    if entered_site == "ALL":
        successes = spacex_df[spacex_df["class"] == 1].groupby("Launch Site").size().reset_index(name="exitos")
        fig = px.pie(
            successes,
            values="exitos",
            names="Launch Site",
            title="Total de lanzamientos exitosos por sitio",
            color_discrete_sequence=["#F26B1D", "#0B1220", "#94A3B8", "#1FA67A"],
        )
    else:
        site_df = spacex_df[spacex_df["Launch Site"] == entered_site]
        counts = site_df["class"].map({1: "Éxito", 0: "Fallo"}).value_counts().reset_index()
        counts.columns = ["Resultado", "Lanzamientos"]
        fig = px.pie(
            counts,
            values="Lanzamientos",
            names="Resultado",
            title=f"Éxito vs. fallo en {entered_site}",
            color="Resultado",
            color_discrete_map=OUTCOME_COLORS,
        )
    fig.update_traces(textinfo="percent+label")
    return fig


@app.callback(
    Output("success-payload-scatter-chart", "figure"),
    [Input("site-dropdown", "value"), Input("payload-slider", "value")],
)
def get_scatter_chart(entered_site: str, payload_range: list[float]):
    low, high = payload_range
    df = spacex_df[spacex_df["Payload Mass (kg)"].between(low, high)]
    title = "Correlación entre masa de carga y éxito en todos los sitios"
    if entered_site != "ALL":
        df = df[df["Launch Site"] == entered_site]
        title = f"Correlación entre masa de carga y éxito en {entered_site}"
    fig = px.scatter(
        df,
        x="Payload Mass (kg)",
        y="class",
        color="Booster Version Category",
        title=title,
        labels={"class": "Resultado (1 = aterrizó)"},
    )
    fig.update_traces(marker={"size": 12, "opacity": 0.8})
    fig.update_yaxes(tickvals=[0, 1])
    return fig


if __name__ == "__main__":
    app.run(debug=False)
