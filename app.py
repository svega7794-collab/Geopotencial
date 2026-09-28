from __future__ import annotations

import math
import sys
import os
import threading
import time
import webbrowser
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path

import numpy as np
import plotly.graph_objects as go
import plotly.io as pio
from flask import Flask, jsonify, request, send_file, Response
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from plotly.subplots import make_subplots
from plotly.offline import get_plotlyjs

APP_NAME = "GeoPotencial 6"
APP_VERSION = "6.0 · PRESTIGE WEB EDITION"
AUTHORS = ["Laura Vargas", "Michael", "Stephany Vega"]

RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
BASE_DIR = Path(__file__).resolve().parent
if os.environ.get("RENDER"):
    OUTPUT_DIR = Path(os.environ.get("GEOPOTENCIAL_OUTPUT_DIR", "/tmp/geopotencial_results"))
else:
    OUTPUT_DIR = Path.home() / "Documents" / "GeoPotencial_Resultados"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
EXCEL_PATH = OUTPUT_DIR / "resultados_geopotencial.xlsx"
REPORT_PATH = OUTPUT_DIR / "reporte_geopotencial_6.html"

app = Flask(__name__)
WORKBOOK_LOCK = threading.Lock()


@dataclass(frozen=True)
class TeacherModel:
    a: float = 6378137.0
    b: float = 6356752.314
    KM: float = 3.98199e14
    e2: float = 0.00669438
    f: float = 0.003352813
    A2: float = 66032801198.0
    A3: float = 5.94179e14
    A4: float = 1.53907e22
    A5: float = 2.42771e27


MODEL = TeacherModel()


@dataclass
class Result:
    phi_deg: float
    lam_deg: float
    h: float
    phi_rad: float
    lam_rad: float
    theta_deg: float
    N: float
    X: float
    Y: float
    Z: float
    r: float
    KM_over_r: float
    CE: float
    CA: float
    CAE: float
    V_CE: float
    V_CA: float
    V_CAE: float
    V_total: float


def fmt(v: float, decimals: int = 4) -> str:
    if not math.isfinite(v):
        return "—"
    if v != 0 and (abs(v) >= 1e12 or abs(v) < 1e-6):
        return f"{v:.6E}"
    return f"{v:.{decimals}f}"


def calculate_model(phi_deg: float, lam_deg: float, h: float) -> Result:
    if not all(math.isfinite(v) for v in (phi_deg, lam_deg, h)):
        raise ValueError("Todos los datos deben ser números finitos.")
    if not -90 <= phi_deg <= 90:
        raise ValueError("La latitud φ debe estar entre −90° y 90°.")
    if not -180 <= lam_deg <= 180:
        raise ValueError("La longitud λ debe estar entre −180° y 180°.")
    if not -1000 <= h <= 1_000_000:
        raise ValueError("Use una altura h entre −1000 m y 1000000 m.")

    phi = math.radians(phi_deg)
    lam = math.radians(lam_deg)
    s = math.sin(phi)
    c = math.cos(phi)

    N = MODEL.a / math.sqrt(1.0 - MODEL.e2 * s**2)
    X = (N + h) * c * math.cos(lam)
    Y = (N + h) * c * math.sin(lam)
    Z = ((1.0 - MODEL.e2) * N + h) * s
    r = math.sqrt(X**2 + Y**2 + Z**2)
    if r <= 0:
        raise ValueError("El radio geocéntrico r debe ser positivo.")

    CE = 1.0 + (MODEL.A2 / r**2) * (1.0 / 3.0 - s**2)

    CA = (
        (MODEL.A3 / r**3) * ((5.0 / 2.0) * s**2 - 3.0 / 2.0)
        + (MODEL.A5 / r**5)
        * (15.0 / 8.0 - (35.0 / 4.0) * s**2 + (63.0 / 8.0) * s**4)
        * s
    )

    CAE = (MODEL.A4 / r**4) * (
        3.0 / 35.0
        + (1.0 / 7.0) * s**2
        - (1.0 / 4.0) * math.sin(2.0 * phi) ** 2
    )

    KM_over_r = MODEL.KM / r
    V_CE = KM_over_r * CE
    V_CA = KM_over_r * CA
    V_CAE = KM_over_r * CAE
    V_total = V_CE + V_CA + V_CAE

    return Result(
        phi_deg=phi_deg,
        lam_deg=lam_deg,
        h=h,
        phi_rad=phi,
        lam_rad=lam,
        theta_deg=90.0 - abs(phi_deg),
        N=N,
        X=X,
        Y=Y,
        Z=Z,
        r=r,
        KM_over_r=KM_over_r,
        CE=CE,
        CA=CA,
        CAE=CAE,
        V_CE=V_CE,
        V_CA=V_CA,
        V_CAE=V_CAE,
        V_total=V_total,
    )


def dms_to_decimal(deg: float, minute: float, sec: float, hemi: str, lat=True) -> float:
    if deg < 0:
        raise ValueError("Escriba los grados positivos y seleccione el hemisferio.")
    if not 0 <= minute < 60:
        raise ValueError("Los minutos deben estar entre 0 y 59.999…")
    if not 0 <= sec < 60:
        raise ValueError("Los segundos deben estar entre 0 y 59.999…")
    limit = 90 if lat else 180
    if deg > limit or (deg == limit and (minute != 0 or sec != 0)):
        raise ValueError(f"Los grados deben estar entre 0° y {limit}°.")
    value = deg + minute / 60 + sec / 3600
    if hemi.upper() in ("S", "W"):
        value *= -1
    return value


def result_payload(r: Result) -> dict:
    s2 = math.sin(r.phi_rad) ** 2
    s4 = math.sin(r.phi_rad) ** 4
    sin2 = math.sin(2 * r.phi_rad) ** 2

    formulas = {
        "CE": "CE = 1 + (A₂/r²) · (1/3 − sen²φ)",
        "CA": "CA = (A₃/r³)·(5/2·sen²φ − 3/2) + (A₅/r⁵)·(15/8 − 35/4·sen²φ + 63/8·sen⁴φ)·senφ",
        "CAE": "CAE = (A₄/r⁴)·(3/35 + 1/7·sen²φ − 1/4·sen²(2φ))",
        "TOTAL": "V = (kM/r) · [CE + CA + CAE]",
    }

    substitutions = {
        "CE": f"CE = 1 + ({MODEL.A2:.4f}/{r.r:.4f}²) · (1/3 − sen²({r.phi_deg:.4f}°)) = {r.CE:.10f}",
        "CA": f"CA = ({MODEL.A3:.6E}/{r.r:.4f}³)·(5/2·{s2:.8f} − 3/2) + ({MODEL.A5:.6E}/{r.r:.4f}⁵)·(15/8 − 35/4·{s2:.8f} + 63/8·{s4:.8f})·sen({r.phi_deg:.4f}°) = {r.CA:.8E}",
        "CAE": f"CAE = ({MODEL.A4:.6E}/{r.r:.4f}⁴)·(3/35 + 1/7·{s2:.8f} − 1/4·{sin2:.8f}) = {r.CAE:.8E}",
        "TOTAL": f"V = {r.KM_over_r:.4f} · [{r.CE:.10f} + ({r.CA:.8E}) + ({r.CAE:.8E})] = {r.V_total:.4f} J/kg",
    }

    return {**asdict(r), "formulas": formulas, "substitutions": substitutions}


# ----------------------------- Excel ---------------------------------

HEAD_FILL = PatternFill("solid", fgColor="3F3A36")
ACCENT_FILL = PatternFill("solid", fgColor="B66C50")
SOFT_FILL = PatternFill("solid", fgColor="F2ECE5")
WHITE_FONT = Font(color="FFFFFF", bold=True)
THIN = Side(style="thin", color="D7CEC4")


def ensure_workbook():
    if EXCEL_PATH.exists():
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"
    ws.append(["GeoPotencial 6", "Prestige Web Edition"])
    ws.append(["Equipo", " · ".join(AUTHORS)])
    ws.append(["Modelo", "WGS84"])
    ws.append(["Ecuación", "V = (kM/r) [CE + CA + CAE]"])

    hist = wb.create_sheet("Historial")
    hist.append(["Resultado", "Fecha", "φ [°]", "λ [°]", "h [m]", "r [m]", "CE", "CA", "CAE", "V [J/kg]"])
    for c in hist[1]:
        c.fill = HEAD_FILL
        c.font = WHITE_FONT
        c.alignment = Alignment(horizontal="center")

    const = wb.create_sheet("Constantes WGS84")
    const.append(["Constante", "Valor", "Unidad"])
    for item in [
        ("a", MODEL.a, "m"), ("b", MODEL.b, "m"), ("e²", MODEL.e2, "—"),
        ("f", MODEL.f, "—"), ("kM", MODEL.KM, "m³/s²"), ("A₂", MODEL.A2, "m²"),
        ("A₃", MODEL.A3, "m³"), ("A₄", MODEL.A4, "m⁴"), ("A₅", MODEL.A5, "m⁵"),
    ]:
        const.append(item)

    for wsx in wb.worksheets:
        for col in range(1, wsx.max_column + 1):
            wsx.column_dimensions[get_column_letter(col)].width = 22
        for row in wsx.iter_rows():
            for c in row:
                c.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    wb.save(EXCEL_PATH)


def register_result(r: Result) -> int:
    ensure_workbook()
    with WORKBOOK_LOCK:
        wb = load_workbook(EXCEL_PATH)
        hist = wb["Historial"]
        n = hist.max_row
        hist.append([
            f"Resultado {n}",
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            round(r.phi_deg, 4),
            round(r.lam_deg, 4),
            round(r.h, 4),
            round(r.r, 4),
            round(r.CE, 10),
            r.CA,
            r.CAE,
            round(r.V_total, 4),
        ])
        for c in hist[hist.max_row]:
            c.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            if c.column in (3, 4, 5, 6, 10):
                c.number_format = "0.0000"
        wb.save(EXCEL_PATH)
    return n


def read_history(limit=100):
    ensure_workbook()
    with WORKBOOK_LOCK:
        wb = load_workbook(EXCEL_PATH, data_only=True)
        hist = wb["Historial"]
        rows = []
        for row in hist.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            rows.append({
                "resultado": row[0], "fecha": row[1], "phi": row[2], "lam": row[3],
                "h": row[4], "r": row[5], "CE": row[6], "CA": row[7],
                "CAE": row[8], "V": row[9],
            })
    return list(reversed(rows[-limit:]))



# ----------------------------- Gráficas -------------------------------

PAPER = "#F7F2EC"
INK = "#2C2926"
MOSS = "#59624E"
CLAY = "#B66C50"
ROSE = "#B98E87"
GOLD = "#B79A68"
SLATE = "#756F67"
GRID = "#D8CFC6"


def base_layout(title):
    return dict(
        title=dict(text=title, x=0.02, xanchor="left", font=dict(size=18, color=INK)),
        paper_bgcolor=PAPER,
        plot_bgcolor=PAPER,
        font=dict(family="Segoe UI, Arial", color=INK),
        margin=dict(l=55, r=35, t=65, b=55),
        xaxis=dict(gridcolor=GRID, zerolinecolor=GRID),
        yaxis=dict(gridcolor=GRID, zerolinecolor=GRID),
        legend=dict(bgcolor="rgba(255,255,255,.55)", bordercolor="#D7CCC1", borderwidth=1),
    )


def graph_profile(r):
    xs = np.linspace(0, 90, 181)
    ys = [calculate_model(float(x), r.lam_deg, r.h).V_total for x in xs]
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=xs, y=ys, mode="lines", line=dict(color=MOSS, width=3),
        fill="tozeroy", fillcolor="rgba(89,98,78,.12)", name="V total"
    ))
    fig.add_trace(go.Scatter(
        x=[abs(r.phi_deg)], y=[r.V_total], mode="markers",
        marker=dict(size=11, color=CLAY, line=dict(color="white", width=2)),
        name="Punto actual"
    ))
    fig.update_layout(**base_layout("Perfil del potencial V(|φ|)"))
    fig.update_xaxes(title="|φ| [°]", range=[0, 90])
    fig.update_yaxes(title="V [J/kg]")
    return fig


def graph_contributions(r):
    """
    Visualización clara de los tres aportes correctivos.
    - Fila 1: resumen en barras del punto actual.
    - Filas 2-4: evolución con la latitud, cada panel con su propia escala.
    Se muestra la corrección A2 de C.E respecto al término central kM/r,
    para que su escala sea comparable con C.A y C.A.E.
    """
    xs = np.linspace(-90, 90, 361)
    ce_corr, ca_vals, cae_vals = [], [], []

    for x in xs:
        rr = calculate_model(float(x), r.lam_deg, r.h)
        ce_corr.append(rr.V_CE - rr.KM_over_r)
        ca_vals.append(rr.V_CA)
        cae_vals.append(rr.V_CAE)

    current_ce = r.V_CE - r.KM_over_r
    current_vals = [current_ce, r.V_CA, r.V_CAE]
    labels = ["A₂ de C.E", "C.A", "C.A.E"]
    colors = [MOSS, CLAY, GOLD]

    fig = make_subplots(
        rows=4, cols=1,
        shared_xaxes=False,
        vertical_spacing=0.08,
        row_heights=[0.22, 0.26, 0.26, 0.26],
        subplot_titles=(
            "Resumen de correcciones en el punto actual",
            "Corrección A₂ de la contribución esférica",
            "Achatamiento C.A (grados 3 y 5)",
            "Asimetría ecuatorial C.A.E (grado 4)",
        ),
    )

    fig.add_trace(
        go.Bar(
            x=labels,
            y=current_vals,
            marker_color=colors,
            text=[f"{v:.4f}" for v in current_vals],
            textposition="outside",
            hovertemplate="%{x}<br>%{y:.6f} J/kg<extra></extra>",
            showlegend=False,
        ),
        row=1, col=1
    )

    series = [
        (ce_corr, MOSS, current_ce),
        (ca_vals, CLAY, r.V_CA),
        (cae_vals, GOLD, r.V_CAE),
    ]

    for row, (ys, color, current_y) in enumerate(series, start=2):
        fig.add_trace(
            go.Scatter(
                x=xs, y=ys,
                mode="lines",
                line=dict(color=color, width=2.8),
                fill="tozeroy",
                fillcolor=color.replace("#", "rgba(") if False else None,
                showlegend=False,
                hovertemplate="φ=%{x:.1f}°<br>Aporte=%{y:.6f} J/kg<extra></extra>",
            ),
            row=row, col=1
        )
        fig.add_trace(
            go.Scatter(
                x=[r.phi_deg], y=[current_y],
                mode="markers",
                marker=dict(size=10, color=ROSE, line=dict(color="white", width=1.8)),
                showlegend=False,
                hovertemplate="Punto actual<br>φ=%{x:.4f}°<br>%{y:.6f} J/kg<extra></extra>",
            ),
            row=row, col=1
        )
        fig.add_vline(
            x=r.phi_deg,
            line_dash="dot",
            line_color=ROSE,
            opacity=.75,
            row=row, col=1
        )

    fig.update_layout(
        height=900,
        paper_bgcolor=PAPER,
        plot_bgcolor=PAPER,
        font=dict(family="Segoe UI, Arial", color=INK),
        margin=dict(l=68, r=40, t=95, b=60),
        title=dict(
            text="Aportes correctivos al potencial",
            x=.02, xanchor="left",
            font=dict(size=20)
        ),
        bargap=.48,
        showlegend=False,
    )

    for rr in range(2, 5):
        fig.update_xaxes(
            range=[-90, 90],
            tickvals=[-90, -60, -30, 0, 30, 60, 90],
            gridcolor=GRID,
            zerolinecolor=GRID,
            row=rr, col=1
        )
        fig.update_yaxes(
            title="J/kg",
            gridcolor=GRID,
            zerolinecolor=GRID,
            row=rr, col=1
        )

    fig.update_xaxes(title="Latitud φ [°]", row=4, col=1)
    fig.update_yaxes(title="J/kg", gridcolor=GRID, row=1, col=1)

    fig.add_annotation(
        text="Nota: en C.E se grafica solo la corrección A₂; el término central kM/r se analiza por separado.",
        xref="paper", yref="paper", x=0.0, y=1.03,
        showarrow=False, align="left",
        font=dict(size=11, color=SLATE),
    )
    return fig

def graph_height(r):
    hs = np.linspace(0, 100000, 121)
    ys, central = [], []
    for h in hs:
        rr = calculate_model(r.phi_deg, r.lam_deg, float(h))
        ys.append(rr.V_total)
        central.append(rr.KM_over_r)

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=hs/1000, y=ys, mode="lines",
                             line=dict(color=CLAY, width=3), name="V total"))
    fig.add_trace(go.Scatter(x=hs/1000, y=central, mode="lines",
                             line=dict(color=SLATE, width=1.8, dash="dash"), name="kM/r"))
    fig.add_trace(go.Scatter(x=[r.h/1000], y=[r.V_total], mode="markers",
                             marker=dict(size=10, color=MOSS, line=dict(color="white", width=2)),
                             name="Punto actual"))
    fig.update_layout(**base_layout(f"Potencial en función de la altura · φ = {r.phi_deg:.4f}°"))
    fig.update_xaxes(title="Altura elipsoidal h [km]")
    fig.update_yaxes(title="V [J/kg]")
    return fig


def graph_ellipsoid(r):
    u = np.linspace(0, 2*np.pi, 78)
    v = np.linspace(-np.pi/2, np.pi/2, 48)
    U, Vv = np.meshgrid(u, v)
    X = MODEL.a*np.cos(Vv)*np.cos(U)/1000
    Y = MODEL.a*np.cos(Vv)*np.sin(U)/1000
    Z = MODEL.b*np.sin(Vv)/1000

    C = np.empty_like(Vv)
    for i in range(Vv.shape[0]):
        for j in range(Vv.shape[1]):
            C[i, j] = calculate_model(float(np.degrees(Vv[i, j])), 0.0, 0.0).V_total

    colorscale = [
        [0.00, "#3E453A"],
        [0.25, "#626955"],
        [0.50, "#B79A68"],
        [0.75, "#B66C50"],
        [1.00, "#6A3C31"],
    ]

    fig = go.Figure()
    fig.add_trace(go.Surface(
        x=X, y=Y, z=Z, surfacecolor=C, colorscale=colorscale,
        colorbar=dict(title="V [J/kg]", len=.65),
        hovertemplate="X=%{x:.1f} km<br>Y=%{y:.1f} km<br>Z=%{z:.1f} km<extra></extra>"
    ))
    fig.add_trace(go.Scatter3d(
        x=[r.X/1000], y=[r.Y/1000], z=[r.Z/1000],
        mode="markers+text", text=["P"], textposition="top center",
        marker=dict(size=7, color="#EBC8BB", line=dict(color=INK, width=2)),
        name="Punto P"
    ))
    fig.update_layout(
        title=dict(text="Elipsoide WGS84 coloreado por potencial", x=.02, xanchor="left"),
        paper_bgcolor=PAPER, font=dict(family="Segoe UI, Arial", color=INK),
        scene=dict(
            bgcolor=PAPER,
            xaxis=dict(title="X [km]", gridcolor=GRID),
            yaxis=dict(title="Y [km]", gridcolor=GRID),
            zaxis=dict(title="Z [km]", gridcolor=GRID),
            aspectmode="data",
        ),
        margin=dict(l=0, r=0, t=60, b=0)
    )
    return fig


def graph_surface(r):
    phis = np.linspace(0, 90, 46)
    hs = np.linspace(0, 100000, 34)
    P, H = np.meshgrid(phis, hs)
    Z = np.empty_like(P)
    for i in range(P.shape[0]):
        for j in range(P.shape[1]):
            Z[i, j] = calculate_model(float(P[i, j]), r.lam_deg, float(H[i, j])).V_total

    fig = go.Figure(go.Surface(
        x=P, y=H/1000, z=Z,
        colorscale=[
            [0.00, "#3E453A"], [0.35, "#74745D"],
            [0.62, "#B79A68"], [0.82, "#B66C50"], [1.00, "#6A3C31"]
        ]
    ))
    fig.update_layout(
        title=dict(text="Superficie V(|φ|, h)", x=.02, xanchor="left"),
        paper_bgcolor=PAPER, font=dict(family="Segoe UI, Arial", color=INK),
        scene=dict(
            bgcolor=PAPER,
            xaxis=dict(title="|φ| [°]", gridcolor=GRID),
            yaxis=dict(title="h [km]", gridcolor=GRID),
            zaxis=dict(title="V [J/kg]", gridcolor=GRID),
        ),
        margin=dict(l=0, r=0, t=60, b=0)
    )
    return fig


def graph_harmonic_shapes(r):
    """
    Tercera visualización 3D: forma angular de cada aporte.
    La deformación radial se exagera deliberadamente para visualizar la
    estructura armónica; no representa la forma física real de la Tierra.
    Incluye C.E(A2), C.A(A3+A5) y C.A.E(A4).
    """
    lon = np.linspace(0, 2*np.pi, 96)
    lat = np.linspace(-np.pi/2, np.pi/2, 64)
    LON, LAT = np.meshgrid(lon, lat)

    s = np.sin(LAT)

    # Factores angulares normalizados de las expresiones del docente.
    ce_shape = (1/3 - s**2)
    ca_shape = ((5/2)*s**2 - 3/2) + 0.65 * (
        15/8 - (35/4)*s**2 + (63/8)*s**4
    ) * s
    cae_shape = 3/35 + (1/7)*s**2 - (1/4)*np.sin(2*LAT)**2

    def normalized_surface(shape, exaggeration=0.42):
        maxabs = np.max(np.abs(shape))
        q = shape / maxabs if maxabs else shape
        rho = 1.0 + exaggeration*q
        X = rho*np.cos(LAT)*np.cos(LON)
        Y = rho*np.cos(LAT)*np.sin(LON)
        Z = rho*np.sin(LAT)
        return X, Y, Z, q

    shapes = [
        ("C.E · término A₂", ce_shape, [[0, "#4B5246"], [0.5, "#E7DDD1"], [1, "#B66C50"]]),
        ("C.A · grados 3 y 5", ca_shape, [[0, "#3E4550"], [0.5, "#E7DDD1"], [1, "#B66C50"]]),
        ("C.A.E · grado 4", cae_shape, [[0, "#59624E"], [0.5, "#E7DDD1"], [1, "#B79A68"]]),
    ]

    fig = go.Figure()

    for i, (name, shape, colorscale) in enumerate(shapes):
        X, Y, Z, q = normalized_surface(shape)
        fig.add_trace(
            go.Surface(
                x=X, y=Y, z=Z,
                surfacecolor=q,
                colorscale=colorscale,
                cmin=-1, cmax=1,
                showscale=True,
                visible=(i == 1),  # abre por defecto en C.A: la "reloj de arena"
                colorbar=dict(
                    title="Aporte normalizado",
                    len=.62,
                    x=1.02
                ),
                name=name,
                hovertemplate=name + "<br>amplitud norm.=%{surfacecolor:.3f}<extra></extra>",
            )
        )

    buttons = []
    for idx, (name, _, _) in enumerate(shapes):
        visible = [False] * len(shapes)
        visible[idx] = True
        buttons.append(dict(
            label=name,
            method="update",
            args=[
                {"visible": visible},
                {"title": {"text": f"Forma 3D de {name} · deformación exagerada", "x": .02, "xanchor": "left"}}
            ],
        ))

    fig.update_layout(
        title=dict(
            text="Forma 3D de C.A · grados 3 y 5 · deformación exagerada",
            x=.02, xanchor="left"
        ),
        paper_bgcolor=PAPER,
        font=dict(family="Segoe UI, Arial", color=INK),
        scene=dict(
            bgcolor=PAPER,
            xaxis=dict(title="X*", gridcolor=GRID),
            yaxis=dict(title="Y*", gridcolor=GRID),
            zaxis=dict(title="Z*", gridcolor=GRID),
            aspectmode="data",
            camera=dict(eye=dict(x=1.45, y=1.45, z=1.05)),
        ),
        updatemenus=[dict(
            type="dropdown",
            direction="down",
            x=.02, y=.98,
            xanchor="left", yanchor="top",
            buttons=buttons,
            bgcolor="#FFFDFC",
            bordercolor=GRID,
            font=dict(color=INK),
        )],
        annotations=[dict(
            text="Visualización cualitativa: la deformación está exagerada para revelar la estructura angular.",
            xref="paper", yref="paper", x=.02, y=.02,
            showarrow=False, align="left",
            font=dict(size=11, color=SLATE),
        )],
        margin=dict(l=0, r=80, t=80, b=20),
    )
    return fig


def parse_result(data):
    mode = data.get("mode", "dms")
    h = float(data.get("h", 2600))

    if mode == "decimal":
        phi = float(data.get("phi", 10))
        lam = float(data.get("lam", -75))
    else:
        phi = dms_to_decimal(
            float(data.get("lat_deg", 10)),
            float(data.get("lat_min", 0)),
            float(data.get("lat_sec", 0)),
            str(data.get("lat_hemi", "N")),
            lat=True
        )
        lam = dms_to_decimal(
            float(data.get("lon_deg", 75)),
            float(data.get("lon_min", 0)),
            float(data.get("lon_sec", 0)),
            str(data.get("lon_hemi", "W")),
            lat=False
        )

    return calculate_model(phi, lam, h)


# ----------------------------- Informe -------------------------------

def generate_report(r):
    payload = result_payload(r)
    history = read_history(50)
    graph1 = graph_contributions(r).to_html(full_html=False, include_plotlyjs=False)
    graph2 = graph_ellipsoid(r).to_html(full_html=False, include_plotlyjs=False)
    plotly_js = get_plotlyjs()

    rows = "".join(
        f"<tr><td>{x['resultado']}</td><td>{x['fecha']}</td><td>{float(x['phi']):.4f}</td>"
        f"<td>{float(x['lam']):.4f}</td><td>{float(x['h']):.4f}</td><td>{float(x['V']):.4f}</td></tr>"
        for x in history
    ) or "<tr><td colspan='6'>Sin resultados registrados.</td></tr>"

    html = f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>GeoPotencial 6 · Informe</title>
<style>
body{{margin:0;background:#F7F2EC;color:#2C2926;font-family:Segoe UI,Arial,sans-serif}}
.wrap{{max-width:1180px;margin:auto;padding:48px 28px}}
.hero{{background:linear-gradient(135deg,#302D2A,#4A4540 55%,#59624E);color:white;border-radius:28px;padding:44px;position:relative;overflow:hidden}}
.hero h1{{font-size:48px;margin:8px 0}} .eyebrow{{color:#D9C39D;letter-spacing:.16em;font-weight:800;font-size:12px}}
.card{{background:#FFFDFC;border:1px solid #D8CFC6;border-radius:20px;padding:20px}}
.grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:18px}}
.section{{margin-top:32px}} .formula{{background:#F0E7DD;border-left:5px solid #B66C50;padding:18px;border-radius:14px;margin:10px 0}}
.total{{background:#59624E;color:white;padding:24px;border-radius:20px;font-size:28px;font-weight:800}}
table{{width:100%;border-collapse:collapse;background:#FFFDFC}} th,td{{padding:12px;border-bottom:1px solid #E1D8D0;text-align:left}} th{{background:#3F3A36;color:white}}
</style><script>{plotly_js}</script></head><body><div class="wrap">
<div class="hero"><div class="eyebrow">GEODESIA FÍSICA · WGS84 · INFORME CIENTÍFICO</div><h1>GeoPotencial 6</h1><div>Potencial gravitacional de la Tierra por aportes</div><p>{' · '.join(AUTHORS)}</p></div>
<div class="section grid">
<div class="card"><small>φ</small><h3>{r.phi_deg:.4f}°</h3></div>
<div class="card"><small>λ</small><h3>{r.lam_deg:.4f}°</h3></div>
<div class="card"><small>h</small><h3>{r.h:.4f} m</h3></div>
<div class="card"><small>r</small><h3>{r.r:.4f} m</h3></div>
</div>
<div class="section"><h2>Resultado general</h2><div class="total">V = {r.V_total:.4f} J/kg</div></div>
<div class="section"><h2>Desarrollo matemático</h2>
<div class="formula"><b>C.E</b><br>{payload['formulas']['CE']}<br><small>{payload['substitutions']['CE']}</small><br><b>VCE = {r.V_CE:.4f} J/kg</b></div>
<div class="formula"><b>C.A</b><br>{payload['formulas']['CA']}<br><small>{payload['substitutions']['CA']}</small><br><b>VCA = {r.V_CA:.4f} J/kg</b></div>
<div class="formula"><b>C.A.E</b><br>{payload['formulas']['CAE']}<br><small>{payload['substitutions']['CAE']}</small><br><b>VCAE = {r.V_CAE:.4f} J/kg</b></div>
<div class="formula"><b>Total</b><br>{payload['formulas']['TOTAL']}<br><small>{payload['substitutions']['TOTAL']}</small></div>
</div>
<div class="section"><h2>Aportes vs latitud</h2>{graph1}</div>
<div class="section"><h2>Representación 3D</h2>{graph2}</div>
<div class="section"><h2>Historial acumulado</h2><table><thead><tr><th>Resultado</th><th>Fecha</th><th>φ</th><th>λ</th><th>h</th><th>V</th></tr></thead><tbody>{rows}</tbody></table></div>
</div></body></html>"""

    REPORT_PATH.write_text(html, encoding="utf-8")
    return REPORT_PATH


# ----------------------------- Rutas --------------------------------

@app.get("/")
def index():
    html_path = RESOURCE_DIR / "index.html"
    return Response(html_path.read_text(encoding="utf-8"), mimetype="text/html")


@app.post("/api/calculate")
def api_calculate():
    try:
        r = parse_result(request.get_json(force=True))
        return jsonify({"ok": True, "result": result_payload(r)})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.post("/api/register")
def api_register():
    try:
        r = parse_result(request.get_json(force=True))
        number = register_result(r)
        return jsonify({"ok": True, "number": number})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.get("/api/history")
def api_history():
    return jsonify({"ok": True, "history": read_history(100)})


@app.post("/api/graph/<kind>")
def api_graph(kind):
    try:
        r = parse_result(request.get_json(force=True))
        if kind == "profile":
            fig = graph_profile(r)
        elif kind == "contributions":
            fig = graph_contributions(r)
        elif kind == "height":
            fig = graph_height(r)
        elif kind == "ellipsoid":
            fig = graph_ellipsoid(r)
        elif kind == "surface":
            fig = graph_surface(r)
        elif kind == "shapes":
            fig = graph_harmonic_shapes(r)
        else:
            raise ValueError("Gráfica no reconocida.")
        return jsonify({"ok": True, "figure": pio.to_json(fig, pretty=False)})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.post("/api/report")
def api_report():
    try:
        r = parse_result(request.get_json(force=True))
        generate_report(r)
        return jsonify({"ok": True, "url": "/download/report"})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.get("/download/excel")
def download_excel():
    ensure_workbook()
    return send_file(EXCEL_PATH, as_attachment=True)


@app.get("/download/report")
def download_report():
    if not REPORT_PATH.exists():
        return "Primero genere el informe.", 404
    return send_file(REPORT_PATH, as_attachment=True)


@app.get("/open-output")
def open_output():
    if os.environ.get("RENDER"):
        return jsonify({
            "ok": False,
            "message": "En la versión web usa Descargar Excel o Generar informe HTML."
        }), 400
    try:
        os.startfile(OUTPUT_DIR)
        return jsonify({"ok": True})
    except Exception as exc:
        return jsonify({"ok": False, "message": str(exc)}), 500


def run_flask():
    host = "0.0.0.0" if os.environ.get("RENDER") else "127.0.0.1"
    port = int(os.environ.get("PORT", "8765"))
    app.run(host=host, port=port, debug=False, use_reloader=False, threaded=True)


def main():
    ensure_workbook()

    if os.environ.get("RENDER"):
        run_flask()
        return

    thread = threading.Thread(target=run_flask, daemon=True)
    thread.start()
    time.sleep(0.8)
    url = "http://127.0.0.1:8765"

    try:
        import webview
        webview.create_window(
            "GeoPotencial 6 · Prestige Web Edition",
            url,
            width=1500,
            height=920,
            min_size=(1150, 720),
        )
        webview.start()
    except Exception:
        webbrowser.open(url)
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass


if __name__ == "__main__":
    main()
