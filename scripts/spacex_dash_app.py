"""SpaceX Falcon 9 landing dashboard (Plotly Dash).

Run locally:   python scripts/spacex_dash_app.py   then open http://127.0.0.1:8050
Data:          data/spacex_clean.csv (created by notebooks/02_EDA_Map_Prediction.ipynb)
"""
import os
import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "..", "data", "spacex_clean.csv")
df = pd.read_csv(DATA)
df["Landing_class"] = df["Landing_class"].astype("float")
SITES = sorted(df["Launch_Site"].unique())
PAYLOAD_MAX = int(df["Payload_kg"].max())


def success_pie(site):
    d = df if site == "ALL" else df[df["Launch_Site"] == site]
    counts = (d.groupby("Landing_Outcome").size().reset_index(name="launches"))
    fig = px.pie(counts, values="launches", names="Landing_Outcome",
                 title=("Landing outcomes: all sites" if site == "ALL" else f"Landing outcomes: {site}"),
                 hole=0.35)
    fig.update_traces(textinfo="percent+label")
    return fig


def payload_scatter(site, low, high):
    d = df[df["Payload_kg"].between(low, high)]
    if site != "ALL":
        d = d[d["Launch_Site"] == site]
    d = d[d["Landing_class"].notna()].copy()
    d["Outcome"] = d["Landing_class"].map({1.0: "Success", 0.0: "Failure"})
    fig = px.scatter(d, x="Payload_kg", y="Booster_family", color="Outcome", symbol="Orbit",
                     color_discrete_map={"Success": "#3E8E62", "Failure": "#C84455"},
                     hover_data=["Date", "Launch_Site", "Landing_Outcome"],
                     title=f"Payload mass vs landing outcome ({low:,}–{high:,} kg)")
    return fig


app = Dash(__name__)
app.title = "SpaceX Falcon 9 Landing Dashboard"

app.layout = html.Div(style={"fontFamily": "Arial", "maxWidth": "1100px", "margin": "auto"}, children=[
    html.H1("SpaceX Falcon 9 Launch Dashboard", style={"textAlign": "center"}),
    html.Label("Launch site"),
    dcc.Dropdown(id="site-dropdown", options=[{"label": "All sites", "value": "ALL"}] +
                 [{"label": s, "value": s} for s in SITES], value="ALL", clearable=False),
    dcc.Graph(id="success-pie"),
    html.Label("Payload mass range (kg)"),
    dcc.RangeSlider(id="payload-slider", min=0, max=PAYLOAD_MAX, step=500,
                    value=[0, PAYLOAD_MAX], marks={0: "0", 5000: "5,000", 10000: "10,000", 15000: "15,000"},
                    tooltip={"placement": "bottom"}),
    dcc.Graph(id="payload-scatter"),
])


@app.callback(Output("success-pie", "figure"), Input("site-dropdown", "value"))
def update_pie(site):
    return success_pie(site)


@app.callback(Output("payload-scatter", "figure"),
              Input("site-dropdown", "value"), Input("payload-slider", "value"))
def update_scatter(site, payload_range):
    low, high = payload_range
    return payload_scatter(site, low, high)


if __name__ == "__main__":
    app.run(debug=False, port=8050)
