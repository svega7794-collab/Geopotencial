from __future__ import annotations

import math
import sys
import textwrap
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
from openpyxl.chart import BarChart, Reference
from plotly.subplots import make_subplots
from plotly.offline import get_plotlyjs

APP_NAME = "GeoPotencial 6"
APP_VERSION = "6.0 · PRESTIGE WEB EDITION"
TEAM = [
    {"name": "Laura Vargas", "student_code": "20241025003", "list_code": "27"},
    {"name": "Michael Ramirez", "student_code": "20232025071", "list_code": "29"},
    {"name": "Stephany Vega", "student_code": "20241025033", "list_code": "28"},
]
AUTHORS = [member["name"] for member in TEAM]

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

# Paleta Prestige: cálida, sobria y consistente con la interfaz web.
XL_PAPER = "F7F2EC"
XL_CARD = "FFFDFC"
XL_SOFT = "EFE7DE"
XL_INK = "2C2926"
XL_GRAPHITE = "3C3834"
XL_MOSS = "59624E"
XL_CLAY = "B66C50"
XL_ROSE = "B98E87"
XL_GOLD = "B79A68"
XL_LINE = "D8CFC6"
XL_WHITE = "FFFFFF"

THIN = Side(style="thin", color=XL_LINE)
MEDIUM = Side(style="medium", color=XL_GRAPHITE)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

TITLE_FONT = Font(name="Aptos Display", size=22, bold=True, color=XL_INK)
SUBTITLE_FONT = Font(name="Aptos", size=11, italic=True, color="756F67")
SECTION_FONT = Font(name="Aptos", size=11, bold=True, color=XL_WHITE)
HEADER_FONT = Font(name="Aptos", size=10, bold=True, color=XL_WHITE)
LABEL_FONT = Font(name="Aptos", size=10, bold=True, color=XL_GRAPHITE)
VALUE_FONT = Font(name="Aptos", size=10, color=XL_INK)
MATH_FONT = Font(name="Cambria Math", size=11, color=XL_INK)
MATH_BOLD = Font(name="Cambria Math", size=11, bold=True, color=XL_INK)
BIG_RESULT_FONT = Font(name="Aptos Display", size=20, bold=True, color=XL_WHITE)
SMALL_WHITE_FONT = Font(name="Aptos", size=9, color=XL_WHITE)

FILL_PAPER = PatternFill("solid", fgColor=XL_PAPER)
FILL_CARD = PatternFill("solid", fgColor=XL_CARD)
FILL_SOFT = PatternFill("solid", fgColor=XL_SOFT)
FILL_GRAPHITE = PatternFill("solid", fgColor=XL_GRAPHITE)
FILL_MOSS = PatternFill("solid", fgColor=XL_MOSS)
FILL_CLAY = PatternFill("solid", fgColor=XL_CLAY)
FILL_GOLD = PatternFill("solid", fgColor=XL_GOLD)
FILL_ROSE = PatternFill("solid", fgColor=XL_ROSE)

PREMIUM_SHEETS = ["Portada", "Resumen", "Desarrollo matemático", "Historial", "Constantes WGS84"]


def _set_sheet_canvas(ws, zoom=90):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = zoom
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.outlinePr.summaryBelow = True


def _paint(ws, cell_range, fill=FILL_CARD, border=BORDER):
    for row in ws[cell_range]:
        for c in row:
            c.fill = fill
            c.border = border


def _merge_label(ws, cell_range, text, fill=FILL_GRAPHITE, font=SECTION_FONT, align="left"):
    ws.merge_cells(cell_range)
    c = ws[cell_range.split(":")[0]]
    c.value = text
    c.fill = fill
    c.font = font
    c.alignment = Alignment(horizontal=align, vertical="center")
    _paint(ws, cell_range, fill=fill)
    c.font = font


def _title(ws, cell_range, text, subtitle=None):
    ws.merge_cells(cell_range)
    c = ws[cell_range.split(":")[0]]
    c.value = text
    c.font = TITLE_FONT
    c.alignment = Alignment(horizontal="left", vertical="center")
    if subtitle:
        row = c.row + 1
        start_col = c.column
        end_col = ws[cell_range.split(":")[1]].column
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
        sc = ws.cell(row=row, column=start_col)
        sc.value = subtitle
        sc.font = SUBTITLE_FONT
        sc.alignment = Alignment(horizontal="left", vertical="center")


def _style_value_cell(c, number_format=None, bold=False, fill=FILL_CARD, align="left"):
    c.fill = fill
    c.border = BORDER
    c.font = Font(name="Aptos", size=10, bold=bold, color=XL_INK)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    if number_format:
        c.number_format = number_format


def _style_header_row(ws, row, start_col, end_col, fill=FILL_GRAPHITE):
    for col in range(start_col, end_col + 1):
        c = ws.cell(row=row, column=col)
        c.fill = fill
        c.font = HEADER_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _apply_history_row_style(ws, row):
    fill = FILL_CARD if row % 2 == 0 else FILL_PAPER
    for col in range(1, 11):
        c = ws.cell(row=row, column=col)
        c.fill = fill
        c.border = BORDER
        c.font = VALUE_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in (3, 4, 5, 6, 10):
        ws.cell(row=row, column=col).number_format = "0.0000"
    for col in (7, 8, 9):
        ws.cell(row=row, column=col).number_format = "0.0000000000E+00"


def _build_premium_workbook(history_rows=None):
    """Create the complete premium workbook, optionally preserving legacy history rows."""
    history_rows = history_rows or []
    wb = Workbook()
    portada = wb.active
    portada.title = "Portada"
    resumen = wb.create_sheet("Resumen")
    desarrollo = wb.create_sheet("Desarrollo matemático")
    hist = wb.create_sheet("Historial")
    const = wb.create_sheet("Constantes WGS84")

    # ---------------- Portada ----------------
    _set_sheet_canvas(portada, 95)
    portada.sheet_properties.tabColor = XL_CLAY
    widths = {"A": 4, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18, "H": 4}
    for col, width in widths.items():
        portada.column_dimensions[col].width = width
    for r in range(1, 32):
        portada.row_dimensions[r].height = 22

    portada.merge_cells("B2:G4")
    c = portada["B2"]
    c.value = "GeoPotencial 6"
    c.font = Font(name="Aptos Display", size=30, bold=True, color=XL_WHITE)
    c.fill = FILL_GRAPHITE
    c.alignment = Alignment(horizontal="left", vertical="center")
    _paint(portada, "B2:G4", FILL_GRAPHITE, Border(bottom=MEDIUM))
    portada["B2"].font = Font(name="Aptos Display", size=30, bold=True, color=XL_WHITE)

    portada.merge_cells("B5:G5")
    portada["B5"] = "PRESTIGE WEB EDITION · GEODESIA FÍSICA"
    portada["B5"].font = Font(name="Aptos", size=11, bold=True, color=XL_CLAY)
    portada["B5"].alignment = Alignment(horizontal="left")

    _merge_label(portada, "B8:G8", "MODELO MATEMÁTICO", FILL_MOSS)
    portada.merge_cells("B9:G11")
    portada["B9"] = "V = (kM/r) · [ C.E + C.A + C.A.E ]"
    portada["B9"].font = Font(name="Cambria Math", size=18, bold=True, color=XL_INK)
    portada["B9"].alignment = Alignment(horizontal="center", vertical="center")
    _paint(portada, "B9:G11", FILL_CARD)
    portada["B9"].font = Font(name="Cambria Math", size=18, bold=True, color=XL_INK)

    _merge_label(portada, "B13:G13", "INFORMACIÓN DEL PROYECTO", FILL_GRAPHITE)
    team_text = " | ".join(
        f'{m["name"]} — Código {m["student_code"]} — Lista {m["list_code"]}'
        for m in TEAM
    )
    info = [
        ("Integrantes", team_text),
        ("Modelo de referencia", "WGS84"),
        ("Magnitud calculada", "Potencial gravitacional terrestre"),
        ("Unidad final", "J/kg (equivalente a m²/s²)"),
    ]
    for i, (label, value) in enumerate(info, start=14):
        portada.merge_cells(start_row=i, start_column=2, end_row=i, end_column=3)
        portada.merge_cells(start_row=i, start_column=4, end_row=i, end_column=7)
        portada.cell(i, 2, label)
        portada.cell(i, 4, value)
        _paint(portada, f"B{i}:C{i}", FILL_SOFT)
        _paint(portada, f"D{i}:G{i}", FILL_CARD)
        portada.cell(i, 2).font = LABEL_FONT
        portada.cell(i, 4).font = VALUE_FONT
        portada.cell(i, 2).alignment = Alignment(vertical="center")
        portada.cell(i, 4).alignment = Alignment(vertical="center", wrap_text=True)
        if i == 14:
            portada.row_dimensions[i].height = 44

    _merge_label(portada, "B20:G20", "RESULTADO MÁS RECIENTE", FILL_CLAY)
    portada.merge_cells("B21:G24")
    portada["B21"] = "Registre un cálculo desde la aplicación para actualizar este bloque."
    portada["B21"].font = BIG_RESULT_FONT
    portada["B21"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    _paint(portada, "B21:G24", FILL_CLAY)
    portada["B21"].font = BIG_RESULT_FONT

    portada.merge_cells("B27:G28")
    portada["B27"] = (
        "El archivo conserva el historial de cálculos y documenta las ecuaciones, "
        "constantes WGS84, sustituciones numéricas y aportes del potencial."
    )
    portada["B27"].font = SUBTITLE_FONT
    portada["B27"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ---------------- Resumen ----------------
    _set_sheet_canvas(resumen, 90)
    resumen.sheet_properties.tabColor = XL_MOSS
    for col, width in {"A": 23, "B": 18, "C": 13, "D": 34, "E": 23, "F": 18, "G": 13, "H": 16}.items():
        resumen.column_dimensions[col].width = width
    _title(resumen, "A1:H1", "Resumen del cálculo", "Variables, geometría, aportes y potencial total")
    resumen.row_dimensions[1].height = 32

    _merge_label(resumen, "A4:C4", "DATOS DE ENTRADA", FILL_MOSS)
    _merge_label(resumen, "E4:H4", "POTENCIAL TOTAL", FILL_CLAY)
    for rr, label, unit in [(5, "Latitud geodésica φ", "°"), (6, "Longitud geodésica λ", "°"), (7, "Altura elipsoidal h", "m")]:
        resumen.cell(rr, 1, label); resumen.cell(rr, 3, unit)
        _style_value_cell(resumen.cell(rr, 1), bold=True, fill=FILL_SOFT)
        _style_value_cell(resumen.cell(rr, 2), "0.0000", align="right")
        _style_value_cell(resumen.cell(rr, 3), fill=FILL_SOFT, align="center")
    resumen.merge_cells("E5:H7")
    resumen["E5"] = "Sin cálculo registrado"
    resumen["E5"].font = BIG_RESULT_FONT
    resumen["E5"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    _paint(resumen, "E5:H7", FILL_CLAY)
    resumen["E5"].font = BIG_RESULT_FONT

    _merge_label(resumen, "A10:C10", "GEOMETRÍA DEL PUNTO", FILL_GRAPHITE)
    geom = [
        (11, "Radio de curvatura N", "m"),
        (12, "Radio geocéntrico r", "m"),
        (13, "Colatitud θ = 90° − |φ|", "°"),
        (14, "Factor central kM/r", "J/kg"),
    ]
    for rr, label, unit in geom:
        resumen.cell(rr, 1, label); resumen.cell(rr, 3, unit)
        _style_value_cell(resumen.cell(rr, 1), bold=True, fill=FILL_SOFT)
        _style_value_cell(resumen.cell(rr, 2), "0.0000", align="right")
        _style_value_cell(resumen.cell(rr, 3), fill=FILL_SOFT, align="center")

    _merge_label(resumen, "E10:H10", "COORDENADAS ECEF", FILL_GRAPHITE)
    for rr, label in [(11, "X"), (12, "Y"), (13, "Z")]:
        resumen.cell(rr, 5, label)
        resumen.cell(rr, 7, "m")
        _style_value_cell(resumen.cell(rr, 5), bold=True, fill=FILL_SOFT)
        resumen.merge_cells(start_row=rr, start_column=6, end_row=rr, end_column=6)
        _style_value_cell(resumen.cell(rr, 6), "0.0000", align="right")
        _style_value_cell(resumen.cell(rr, 7), fill=FILL_SOFT, align="center")
        _style_value_cell(resumen.cell(rr, 8), fill=FILL_PAPER)

    _merge_label(resumen, "A17:D17", "APORTES AL POTENCIAL", FILL_MOSS)
    headers = ["Aporte", "Factor", "Potencial [J/kg]", "Lectura física"]
    for i, h in enumerate(headers, start=1):
        resumen.cell(18, i, h)
    _style_header_row(resumen, 18, 1, 4, FILL_GRAPHITE)
    aporte_rows = [
        (19, "C.E · Contribución esférica", "Término dominante + corrección A₂"),
        (20, "C.A · Achatamiento", "Correcciones asociadas a A₃ y A₅"),
        (21, "C.A.E · Asimetría ecuatorial", "Corrección asociada a A₄"),
        (22, "TOTAL", "Suma de los tres aportes"),
    ]
    for rr, label, desc in aporte_rows:
        resumen.cell(rr, 1, label); resumen.cell(rr, 4, desc)
        for cc in range(1, 5):
            _style_value_cell(resumen.cell(rr, cc), fill=FILL_CARD)
        resumen.cell(rr, 1).font = LABEL_FONT if rr < 22 else Font(name="Aptos", size=10, bold=True, color=XL_WHITE)
        if rr == 22:
            for cc in range(1, 5):
                resumen.cell(rr, cc).fill = FILL_CLAY
                resumen.cell(rr, cc).font = Font(name="Aptos", size=10, bold=True, color=XL_WHITE)
        resumen.cell(rr, 2).number_format = "0.0000000000E+00"
        resumen.cell(rr, 3).number_format = "0.0000"
        resumen.row_dimensions[rr].height = 34 if rr < 22 else 30

    _merge_label(resumen, "F17:H17", "CORRECCIONES COMPARABLES", FILL_GOLD)
    resumen["F18"] = "Componente"; resumen["G18"] = "J/kg"
    _style_header_row(resumen, 18, 6, 7, FILL_GOLD)
    for rr, label in [(19, "A₂ de C.E"), (20, "C.A"), (21, "C.A.E")]:
        resumen.cell(rr, 6, label)
        _style_value_cell(resumen.cell(rr, 6), fill=FILL_SOFT, bold=True)
        _style_value_cell(resumen.cell(rr, 7), "0.0000", fill=FILL_CARD, align="right")

    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = "Magnitud de correcciones"
    chart.y_axis.title = "Componente"
    chart.x_axis.title = "J/kg"
    chart.height = 5.6
    chart.width = 10.0
    data = Reference(resumen, min_col=7, min_row=18, max_row=21)
    cats = Reference(resumen, min_col=6, min_row=19, max_row=21)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.legend = None
    try:
        chart.series[0].graphicalProperties.solidFill = XL_CLAY
    except Exception:
        pass
    resumen.add_chart(chart, "E24")

    # ---------------- Desarrollo matemático ----------------
    _set_sheet_canvas(desarrollo, 88)
    desarrollo.sheet_properties.tabColor = XL_GOLD
    for col, width in {"A": 26, "B": 74, "C": 16, "D": 16}.items():
        desarrollo.column_dimensions[col].width = width
    _title(desarrollo, "A1:D1", "Desarrollo matemático", "Ecuaciones del docente, sustitución numérica y resultado de cada aporte")
    desarrollo.row_dimensions[1].height = 32

    equations = [
        (4, "Ecuación general", "V = (kM/r) · [C.E + C.A + C.A.E]"),
        (7, "Contribución esférica · C.E", "C.E = 1 + (A₂/r²) · [1/3 − sen²φ]"),
        (12, "Achatamiento · C.A", "C.A = (A₃/r³)·[(5/2)sen²φ − 3/2] + (A₅/r⁵)·[15/8 − (35/4)sen²φ + (63/8)sen⁴φ]·senφ"),
        (18, "Asimetría ecuatorial · C.A.E", "C.A.E = (A₄/r⁴)·[3/35 + (1/7)sen²φ − (1/4)sen²(2φ)]"),
    ]
    for rr, label, equation in equations:
        _merge_label(desarrollo, f"A{rr}:D{rr}", label, FILL_GRAPHITE if rr != 4 else FILL_MOSS)
        desarrollo.merge_cells(start_row=rr+1, start_column=1, end_row=rr+1, end_column=4)
        desarrollo.cell(rr+1, 1, equation)
        desarrollo.cell(rr+1, 1).font = MATH_BOLD
        desarrollo.cell(rr+1, 1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        _paint(desarrollo, f"A{rr+1}:D{rr+1}", FILL_CARD)
        desarrollo.cell(rr+1, 1).font = MATH_BOLD
        desarrollo.row_dimensions[rr+1].height = 34

    _merge_label(desarrollo, "A23:D23", "SUSTITUCIÓN NUMÉRICA DEL ÚLTIMO CÁLCULO", FILL_CLAY)
    rows = [
        (24, "C.E", "Registre un cálculo para ver la sustitución.", ""),
        (26, "C.A", "Registre un cálculo para ver la sustitución.", ""),
        (28, "C.A.E", "Registre un cálculo para ver la sustitución.", ""),
        (30, "TOTAL", "Registre un cálculo para ver la sustitución.", ""),
    ]
    for rr, label, text, result in rows:
        desarrollo.cell(rr, 1, label)
        desarrollo.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
        desarrollo.cell(rr, 2, text)
        _style_value_cell(desarrollo.cell(rr, 1), bold=True, fill=FILL_SOFT, align="center")
        _paint(desarrollo, f"B{rr}:D{rr}", FILL_CARD)
        desarrollo.cell(rr, 2).font = MATH_FONT
        desarrollo.cell(rr, 2).alignment = Alignment(wrap_text=True, vertical="center")
        desarrollo.row_dimensions[rr].height = 48 if rr != 30 else 58

    desarrollo.merge_cells("A33:D36")
    desarrollo["A33"] = "Resultado final pendiente"
    desarrollo["A33"].font = BIG_RESULT_FONT
    desarrollo["A33"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    _paint(desarrollo, "A33:D36", FILL_MOSS)
    desarrollo["A33"].font = BIG_RESULT_FONT

    desarrollo.merge_cells("A38:D40")
    desarrollo["A38"] = (
        "Nota: λ interviene en las coordenadas X e Y del punto, pero este modelo de potencial "
        "zonal depende de φ, r y los coeficientes A₂, A₃, A₄ y A₅; por ello λ no cambia V."
    )
    desarrollo["A38"].font = SUBTITLE_FONT
    desarrollo["A38"].alignment = Alignment(wrap_text=True, vertical="center")
    _paint(desarrollo, "A38:D40", FILL_SOFT)
    desarrollo["A38"].font = SUBTITLE_FONT

    # ---------------- Historial ----------------
    _set_sheet_canvas(hist, 90)
    hist.sheet_properties.tabColor = XL_ROSE
    widths = [16, 21, 12, 12, 14, 17, 18, 18, 18, 18]
    for idx, width in enumerate(widths, start=1):
        hist.column_dimensions[get_column_letter(idx)].width = width
    hist.merge_cells("A1:J2")
    hist["A1"] = "Historial de resultados"
    hist["A1"].font = TITLE_FONT
    hist["A1"].alignment = Alignment(vertical="center")
    hist.row_dimensions[1].height = 26
    hist.append([])  # row 3 spacer
    headers = ["Resultado", "Fecha", "φ [°]", "λ [°]", "h [m]", "r [m]", "CE", "CA", "CAE", "V [J/kg]"]
    for col, val in enumerate(headers, 1):
        hist.cell(4, col, val)
    _style_header_row(hist, 4, 1, 10, FILL_GRAPHITE)
    hist.freeze_panes = "A5"
    hist.auto_filter.ref = "A4:J4"
    for row_data in history_rows:
        hist.append(list(row_data))
        _apply_history_row_style(hist, hist.max_row)
    if history_rows:
        hist.auto_filter.ref = f"A4:J{hist.max_row}"

    # ---------------- Constantes ----------------
    _set_sheet_canvas(const, 92)
    const.sheet_properties.tabColor = XL_GRAPHITE
    for col, width in {"A": 20, "B": 25, "C": 16, "D": 48}.items():
        const.column_dimensions[col].width = width
    _title(const, "A1:D1", "Constantes WGS84 y coeficientes", "Valores fijos utilizados por el modelo")
    const.row_dimensions[1].height = 32
    headers = ["Constante", "Valor", "Unidad", "Descripción"]
    for col, val in enumerate(headers, 1):
        const.cell(4, col, val)
    _style_header_row(const, 4, 1, 4, FILL_GRAPHITE)
    constants = [
        ("a", MODEL.a, "m", "Semieje mayor WGS84"),
        ("b", MODEL.b, "m", "Semieje menor WGS84"),
        ("e²", MODEL.e2, "—", "Primera excentricidad al cuadrado"),
        ("f", MODEL.f, "—", "Achatamiento geométrico"),
        ("kM", MODEL.KM, "m³/s²", "Constante gravitacional por masa terrestre usada en el ejercicio"),
        ("A₂", MODEL.A2, "m²", "Coeficiente de expansión, grado 2"),
        ("A₃", MODEL.A3, "m³", "Coeficiente de expansión, grado 3"),
        ("A₄", MODEL.A4, "m⁴", "Coeficiente de expansión, grado 4"),
        ("A₅", MODEL.A5, "m⁵", "Coeficiente de expansión, grado 5"),
    ]
    for rr, row_data in enumerate(constants, start=5):
        for cc, val in enumerate(row_data, start=1):
            const.cell(rr, cc, val)
            _style_value_cell(const.cell(rr, cc), fill=FILL_CARD)
        const.cell(rr, 1).font = LABEL_FONT
        const.cell(rr, 2).number_format = "0.0000000000E+00" if abs(float(row_data[1])) >= 1e9 else "0.0000000000"
        const.cell(rr, 4).alignment = Alignment(wrap_text=True, vertical="center")
    const.freeze_panes = "A5"

    wb.active = 0
    return wb


def _legacy_history_rows(path: Path):
    """Read history from an older workbook so upgrading the design does not lose records."""
    if not path.exists():
        return []
    try:
        old = load_workbook(path, data_only=False)
        if "Historial" not in old.sheetnames:
            return []
        ws = old["Historial"]
        start_row = 5 if ws["A4"].value == "Resultado" else 2
        rows = []
        for row in ws.iter_rows(min_row=start_row, max_col=10, values_only=True):
            if row[0]:
                rows.append(row[:10])
        return rows
    except Exception:
        return []


def ensure_workbook():
    """Create the premium workbook or upgrade a legacy workbook while preserving history."""
    with WORKBOOK_LOCK:
        if EXCEL_PATH.exists():
            try:
                wb = load_workbook(EXCEL_PATH, read_only=True)
                is_premium = all(name in wb.sheetnames for name in PREMIUM_SHEETS) and wb["Portada"]["B2"].value == "GeoPotencial 6"
                wb.close()
                if is_premium:
                    return
            except Exception:
                pass
        history = _legacy_history_rows(EXCEL_PATH)
        wb = _build_premium_workbook(history)
        wb.save(EXCEL_PATH)


def _update_latest_result_sheets(wb, r: Result):
    """Refresh Portada, Resumen and Desarrollo matemático with the latest calculation."""
    payload = result_payload(r)

    # Portada
    portada = wb["Portada"]
    portada["B21"] = f"V = {r.V_total:.4f} J/kg\nφ = {r.phi_deg:.4f}°   ·   λ = {r.lam_deg:.4f}°   ·   h = {r.h:.4f} m"
    portada["B21"].font = BIG_RESULT_FONT
    portada["B21"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Resumen
    ws = wb["Resumen"]
    ws["B5"] = r.phi_deg; ws["B6"] = r.lam_deg; ws["B7"] = r.h
    ws["E5"] = f"V = {r.V_total:.4f} J/kg"
    ws["E5"].font = BIG_RESULT_FONT
    ws["E5"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B11"] = r.N; ws["B12"] = r.r; ws["B13"] = r.theta_deg; ws["B14"] = r.KM_over_r
    ws["F11"] = r.X; ws["F12"] = r.Y; ws["F13"] = r.Z
    for cell in ["B5", "B6", "B7", "B11", "B12", "B13", "B14", "F11", "F12", "F13"]:
        ws[cell].number_format = "0.0000"
    values = [
        (19, r.CE, r.V_CE),
        (20, r.CA, r.V_CA),
        (21, r.CAE, r.V_CAE),
        (22, None, r.V_total),
    ]
    for rr, factor, pot in values:
        ws.cell(rr, 2, factor if factor is not None else "—")
        ws.cell(rr, 3, pot)
        ws.cell(rr, 3).number_format = "0.0000"
        if factor is not None:
            ws.cell(rr, 2).number_format = "0.0000000000E+00"
    ws["G19"] = r.V_CE - r.KM_over_r
    ws["G20"] = r.V_CA
    ws["G21"] = r.V_CAE
    for c in ("G19", "G20", "G21"):
        ws[c].number_format = "0.0000"

    # Desarrollo matemático
    dev = wb["Desarrollo matemático"]
    dev["B24"] = payload["substitutions"]["CE"] + f"\nV_CE = (kM/r)·CE = {r.V_CE:.4f} J/kg"
    dev["B26"] = payload["substitutions"]["CA"] + f"\nV_CA = (kM/r)·CA = {r.V_CA:.4f} J/kg"
    dev["B28"] = payload["substitutions"]["CAE"] + f"\nV_CAE = (kM/r)·CAE = {r.V_CAE:.4f} J/kg"
    dev["B30"] = payload["substitutions"]["TOTAL"]
    for cell in ("B24", "B26", "B28", "B30"):
        dev[cell].font = MATH_FONT
        dev[cell].alignment = Alignment(wrap_text=True, vertical="center")
    dev["A33"] = (
        f"V = {r.V_total:.4f} J/kg\n"
        f"C.E = {r.V_CE:.4f} J/kg   ·   C.A = {r.V_CA:.4f} J/kg   ·   C.A.E = {r.V_CAE:.4f} J/kg"
    )
    dev["A33"].font = BIG_RESULT_FONT
    dev["A33"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def register_result(r: Result) -> int:
    ensure_workbook()
    with WORKBOOK_LOCK:
        wb = load_workbook(EXCEL_PATH)
        hist = wb["Historial"]
        # Header is on row 4 in the premium workbook.
        existing = max(0, hist.max_row - 4)
        n = existing + 1
        row = hist.max_row + 1
        hist.cell(row, 1, f"Resultado {n}")
        hist.cell(row, 2, datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        hist.cell(row, 3, round(r.phi_deg, 4))
        hist.cell(row, 4, round(r.lam_deg, 4))
        hist.cell(row, 5, round(r.h, 4))
        hist.cell(row, 6, round(r.r, 4))
        hist.cell(row, 7, r.CE)
        hist.cell(row, 8, r.CA)
        hist.cell(row, 9, r.CAE)
        hist.cell(row, 10, round(r.V_total, 4))
        _apply_history_row_style(hist, row)
        hist.auto_filter.ref = f"A4:J{hist.max_row}"
        _update_latest_result_sheets(wb, r)
        wb.save(EXCEL_PATH)
    return n


def read_history(limit=100):
    ensure_workbook()
    with WORKBOOK_LOCK:
        wb = load_workbook(EXCEL_PATH, data_only=True)
        hist = wb["Historial"]
        rows = []
        for row in hist.iter_rows(min_row=5, values_only=True):
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


ACCENT = "#2F6F8F"      # azul petróleo para líneas de referencia
G_REF = 9.80665          # gravedad de referencia, solo para expresar J/kg como metros


# ---------- utilidades comunes de las gráficas ----------

def _v(x: float, dec: int = 2) -> str:
    """Número con separador de miles por espacio (62 440 001.43)."""
    return f"{x:,.{dec}f}".replace(",", " ").replace("-", "−")


def _pct(p: float) -> str:
    """Porcentaje legible sin notación científica (0.000074 %)."""
    return np.format_float_positional(p, precision=3 if p < 1 else 4, fractional=p >= 1, trim="-") + " %"


def _padding_y(fig, ys, fila=None, abajo=.12, arriba=.22):
    lo, hi = float(np.min(ys)), float(np.max(ys))
    span = (hi - lo) or abs(hi) or 1.0
    kw = dict(row=fila, col=1) if fila else {}
    fig.update_yaxes(range=[lo - abajo * span, hi + arriba * span], **kw)


def _lat_txt(phi: float) -> str:
    return f"{abs(phi):.4f}° {'N' if phi >= 0 else 'S'}"


def _wrap(texto: str, ancho: int = 140) -> str:
    """Parte un párrafo en líneas para Plotly (que no ajusta texto solo)."""
    lineas = []
    for parrafo in texto.split("\n"):
        lineas.extend(textwrap.wrap(parrafo, ancho) or [""])
    return "<br>".join(lineas)


def _bloque_explicacion(texto: str, ancho: int = 140) -> tuple[str, int]:
    """Devuelve el texto listo para anotación y la altura en px que ocupa."""
    cuerpo = _wrap(texto, ancho)
    n = cuerpo.count("<br>") + 2
    return "<b>¿Qué muestra esta gráfica?</b><br>" + cuerpo, 18 * n + 26


def _agregar_explicacion(fig, texto: str, alto_figura: int, espacio_ejes: int = 58, ancho: int = 140):
    """Agrega la explicación como un recuadro bajo la gráfica, dentro de la misma figura."""
    contenido, alto_txt = _bloque_explicacion(texto, ancho)
    fig.add_annotation(
        text=contenido, xref="paper", yref="paper", x=0, y=0,
        xanchor="left", yanchor="top", yshift=-espacio_ejes, showarrow=False, align="left",
        font=dict(size=12.5, color=INK), bgcolor="#FFFDFC", bordercolor=GRID, borderwidth=1, borderpad=10,
    )
    fig.update_layout(height=alto_figura + alto_txt, margin=dict(b=espacio_ejes + alto_txt + 12))
    return fig


def _cruces_cero(xs, ys):
    """Latitudes (interpoladas) donde una curva cambia de signo."""
    out = []
    for i in range(len(xs) - 1):
        if ys[i] == 0:
            out.append(float(xs[i]))
        elif ys[i] * ys[i + 1] < 0:
            out.append(float(xs[i] - ys[i] * (xs[i + 1] - xs[i]) / (ys[i + 1] - ys[i])))
    return out


def _barrido_latitud(r, n=361):
    """Evalúa el modelo de −90° a 90° con la misma λ y h del punto P."""
    xs = np.linspace(-90, 90, n)
    filas = [calculate_model(float(x), r.lam_deg, r.h) for x in xs]
    return xs, filas


def _gradiente_vertical(phi, lam, h, dh=1.0):
    """dV/dh numérico (J/kg por metro = m/s²)."""
    h1 = h - dh if h - dh >= -1000 else h
    h2 = h + dh if h + dh <= 1_000_000 else h
    v1 = calculate_model(phi, lam, h1).V_total
    v2 = calculate_model(phi, lam, h2).V_total
    return (v2 - v1) / (h2 - h1)


# ---------- 1. Perfil del potencial total por latitud ----------

def graph_profile(r):
    xs, filas = _barrido_latitud(r)
    ys = np.array([f.V_total for f in filas])
    i_eq = int(np.argmin(np.abs(xs)))
    v_eq, v_n, v_s = ys[i_eq], ys[-1], ys[0]
    dif_p = r.V_total - v_eq

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=xs, y=ys, mode="lines", line=dict(color=MOSS, width=3), name="V total a la altura de P",
        fill="tozeroy", fillcolor="rgba(89,98,78,.10)",
        hovertemplate="φ = %{x:.1f}°<br>V = %{y:,.2f} J/kg<extra></extra>",
    ))
    fig.add_hline(y=r.V_total, line_dash="dot", line_color=ROSE, opacity=.9)
    fig.add_trace(go.Scatter(
        x=[0, 90, -90], y=[v_eq, v_n, v_s], mode="markers+text", name="Ecuador y polos",
        marker=dict(size=8, color=SLATE), text=["Ecuador", "Polo N", "Polo S"],
        textposition=["top center", "bottom left", "bottom right"], textfont=dict(size=11, color=SLATE),
        hovertemplate="%{text}<br>V = %{y:,.2f} J/kg<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=[r.phi_deg], y=[r.V_total], mode="markers", name="Punto P",
        marker=dict(size=13, color=CLAY, line=dict(color="white", width=2)),
        hovertemplate=f"Punto P<br>φ = {_lat_txt(r.phi_deg)}<br>V = %{{y:,.3f}} J/kg<extra></extra>",
    ))
    rango = ys.max() - ys.min()
    fig.update_layout(**base_layout(f"¿Cómo cambia V con la latitud?  ·  h = {_v(r.h, 1)} m, λ = {_v(r.lam_deg, 4)}°"))
    fig.update_layout(separators=". ", legend=dict(orientation="h", x=.5, xanchor="center", y=.99, yanchor="top",
                                                   bgcolor="rgba(255,253,252,.85)"))
    fig.update_xaxes(title="Latitud φ [°]  (negativo = Sur)", range=[-90, 90],
                     tickvals=[-90, -60, -30, 0, 30, 60, 90])
    fig.update_yaxes(title="V [J/kg]", tickformat=",.0f",
                     range=[ys.min() - .12 * rango, ys.max() + .08 * rango])

    texto = (
        f"La curva da el potencial total V que tendría un punto con la misma altura (h = {_v(r.h, 1)} m) "
        f"si se desplazara de polo a polo. V es mínimo en el ecuador ({_v(v_eq)} J/kg) y máximo en los polos "
        f"(N: {_v(v_n)} J/kg; S: {_v(v_s)} J/kg): una diferencia de unos {_v(max(v_n, v_s) - v_eq, 0)} J/kg. "
        f"La causa es el achatamiento: en los polos la superficie está más cerca del centro de la Tierra, r es menor "
        f"y el término kM/r crece.\n"
        f"El punto P ({_lat_txt(r.phi_deg)}) está {_v(dif_p)} J/kg por encima del valor en el ecuador, lo que equivale "
        f"a unos {_v(dif_p / G_REF, 1)} m de altura (ΔV/g). La línea punteada une las latitudes con el mismo V que P."
    )
    return _agregar_explicacion(fig, texto, alto_figura=560)


# ---------- 2. Peso de cada aporte y su variación con la latitud ----------

def graph_contributions(r):
    xs, filas = _barrido_latitud(r)
    ce_corr = np.array([f.V_CE - f.KM_over_r for f in filas])
    ca_vals = np.array([f.V_CA for f in filas])
    cae_vals = np.array([f.V_CAE for f in filas])
    corr_p = r.V_CE - r.KM_over_r

    fig = make_subplots(
        rows=4, cols=1, vertical_spacing=.085, row_heights=[.25, .25, .25, .25],
        subplot_titles=(
            "1 · Peso de cada término en V para el punto P (escala logarítmica)",
            "2 · Corrección A₂ de la contribución esférica C.E (grado 2)",
            "3 · Achatamiento C.A (grados 3 y 5)",
            "4 · Asimetría ecuatorial C.A.E (grado 4)",
        ),
    )

    # Panel 1: tamaño de cada término (el signo va en la etiqueta)
    nombres = ["kM/r (Tierra esférica)", "A₂ de C.E", "C.A", "C.A.E"]
    valores = [r.KM_over_r, corr_p, r.V_CA, r.V_CAE]
    colores = [SLATE, MOSS, CLAY, GOLD]
    etiquetas = [f"{'+' if v >= 0 else '−'}{_v(abs(v), 3 if abs(v) < 1e3 else 1)} J/kg  ·  "
                 f"{_pct(abs(v) / r.V_total * 100)} de V" for v in valores]
    fig.add_trace(go.Bar(
        y=nombres, x=[max(abs(v), 1e-6) for v in valores], orientation="h", marker_color=colores,
        text=etiquetas, textposition="outside", cliponaxis=False, showlegend=False,
        customdata=valores, hovertemplate="%{y}<br>%{customdata:,.4f} J/kg<extra></extra>",
    ), row=1, col=1)
    fig.update_xaxes(type="log", range=[0, 10.3], title="|aporte| [J/kg]  (cada división = ×10)",
                     gridcolor=GRID, row=1, col=1)
    fig.update_yaxes(autorange="reversed", row=1, col=1)
    fig.update_xaxes(zeroline=False, row=1, col=1)

    series = [(ce_corr, MOSS, corr_p), (ca_vals, CLAY, r.V_CA), (cae_vals, GOLD, r.V_CAE)]
    for fila, (ys, color, actual) in enumerate(series, start=2):
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="lines", line=dict(color=color, width=2.8), showlegend=False,
            hovertemplate="φ = %{x:.1f}°<br>%{y:,.3f} J/kg<extra></extra>",
        ), row=fila, col=1)
        for x0 in _cruces_cero(xs, ys):
            fig.add_vline(x=x0, line_color=SLATE, line_width=1, opacity=.35, row=fila, col=1)
        fig.add_vline(x=r.phi_deg, line_dash="dot", line_color=ROSE, opacity=.8, row=fila, col=1)
        fig.add_trace(go.Scatter(
            x=[r.phi_deg], y=[actual], mode="markers+text", showlegend=False,
            text=[f"P: {_v(actual, 3)}"], textposition="top right", textfont=dict(size=11, color=INK),
            marker=dict(size=10, color=ROSE, line=dict(color="white", width=1.8)),
            hovertemplate="Punto P<br>%{y:,.4f} J/kg<extra></extra>",
        ), row=fila, col=1)
        fig.update_xaxes(range=[-90, 90], tickvals=[-90, -60, -30, 0, 30, 60, 90], gridcolor=GRID,
                         zeroline=False, row=fila, col=1)
        _padding_y(fig, ys, fila)
        fig.update_yaxes(title="J/kg", gridcolor=GRID, zeroline=True, zerolinecolor="#9C928A",
                         zerolinewidth=1.5, tickformat=",.0f" if fila == 2 else ",.1f", row=fila, col=1)
    fig.update_xaxes(title="Latitud φ [°]", row=4, col=1)

    fig.update_layout(
        paper_bgcolor=PAPER, plot_bgcolor=PAPER, separators=". ",
        font=dict(family="Segoe UI, Arial", color=INK),
        title=dict(text="¿De qué está hecho V? Aportes del modelo en el punto P y según la latitud",
                   x=.02, xanchor="left", font=dict(size=19)),
        margin=dict(l=175, r=40, t=95),
        showlegend=False,
    )
    for anot in fig.layout.annotations:  # títulos de los paneles alineados a la izquierda
        anot.update(x=0, xanchor="left", font=dict(size=13.5, color=INK))

    ceros_ce = [abs(x) for x in _cruces_cero(xs, ce_corr)]
    ceros_cae = sorted({round(abs(x), 2) for x in _cruces_cero(xs, cae_vals)})
    texto = (
        f"V se construye como kM/r (una Tierra esférica) más tres correcciones por su forma. El panel 1 compara su "
        f"tamaño en P: kM/r aporta prácticamente todo el potencial ({r.KM_over_r / r.V_total * 100:.4f} % de V); "
        f"la corrección A₂ es del orden de decenas de miles de J/kg y C.A y C.A.E son de decenas o centenas de J/kg. "
        f"Por eso la escala es logarítmica: en escala normal solo se vería la barra de kM/r.\n"
        f"Los paneles 2 a 4 muestran cómo cambia cada corrección si P se mueve en latitud (misma h y λ); el punto rosado "
        f"es P y las líneas grises verticales marcan dónde la corrección vale cero. La corrección A₂ refleja el "
        f"abultamiento ecuatorial: suma potencial cerca del ecuador y resta hacia los polos, y cambia de signo en "
        f"±{ceros_ce[0] if ceros_ce else 35.26:.2f}° (donde sen²φ = 1/3). C.A reúne los grados 3 y 5, los términos que "
        f"distinguen el hemisferio norte del sur. C.A.E (grado 4) forma bandas y cambia de signo en "
        f"±{' y ±'.join(f'{c:.2f}°' for c in ceros_cae) if ceros_cae else '—'}."
    )
    return _agregar_explicacion(fig, texto, alto_figura=1080, espacio_ejes=60)


# ---------- 3. Potencial y gravedad en función de la altura ----------

def graph_height(r):
    h_max = min(max(100_000.0, 1.3 * r.h), 1_000_000.0)
    hs = np.linspace(0.0, h_max, 161)
    v_tot, central, grad = [], [], []
    for h in hs:
        rr = calculate_model(r.phi_deg, r.lam_deg, float(h))
        v_tot.append(rr.V_total)
        central.append(rr.KM_over_r)
        grad.append(-_gradiente_vertical(r.phi_deg, r.lam_deg, float(h)))
    g_p = -_gradiente_vertical(r.phi_deg, r.lam_deg, r.h)
    km = hs / 1000

    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=.11, row_heights=[.6, .4],
                        subplot_titles=("Potencial V al subir sobre P",
                                        "Pendiente de V: la atracción gravitacional (−dV/dh)"))
    fig.add_trace(go.Scatter(x=km, y=v_tot, mode="lines", line=dict(color=CLAY, width=3), name="V total",
                             hovertemplate="h = %{x:.1f} km<br>V = %{y:,.2f} J/kg<extra></extra>"), row=1, col=1)
    fig.add_trace(go.Scatter(x=km, y=central, mode="lines", line=dict(color=SLATE, width=1.6, dash="dash"),
                             name="kM/r (Tierra esférica)",
                             hovertemplate="h = %{x:.1f} km<br>kM/r = %{y:,.2f} J/kg<extra></extra>"), row=1, col=1)
    fig.add_trace(go.Scatter(x=[r.h / 1000], y=[r.V_total], mode="markers", name="Punto P",
                             marker=dict(size=12, color=MOSS, line=dict(color="white", width=2)),
                             hovertemplate="Punto P<br>V = %{y:,.3f} J/kg<extra></extra>"), row=1, col=1)
    fig.add_trace(go.Scatter(x=km, y=grad, mode="lines", line=dict(color=ACCENT, width=2.6), showlegend=False,
                             hovertemplate="h = %{x:.1f} km<br>−dV/dh = %{y:.4f} m/s²<extra></extra>"), row=2, col=1)
    fig.add_trace(go.Scatter(x=[r.h / 1000], y=[g_p], mode="markers+text", showlegend=False,
                             text=[f"P: {g_p:.4f} m/s²"], textposition="top right",
                             marker=dict(size=11, color=MOSS, line=dict(color="white", width=2)),
                             hovertemplate="Punto P<br>%{y:.5f} m/s²<extra></extra>"), row=2, col=1)

    fig.update_layout(
        paper_bgcolor=PAPER, plot_bgcolor=PAPER, separators=". ",
        font=dict(family="Segoe UI, Arial", color=INK),
        title=dict(text=f"¿Qué pasa con V al alejarse de la Tierra?  ·  φ = {_lat_txt(r.phi_deg)}",
                   x=.02, xanchor="left", font=dict(size=18)),
        legend=dict(x=.99, xanchor="right", y=.99, yanchor="top", bgcolor="rgba(255,253,252,.85)",
                    bordercolor=GRID, borderwidth=1),
        margin=dict(l=90, r=35, t=90),
    )
    for anot in fig.layout.annotations:
        anot.update(x=0, xanchor="left", font=dict(size=13.5, color=INK))
    fig.update_xaxes(gridcolor=GRID, zerolinecolor=GRID)
    fig.update_xaxes(title="Altura elipsoidal h [km]", row=2, col=1)
    fig.update_yaxes(title="V [J/kg]", tickformat=",.0f", gridcolor=GRID, row=1, col=1)
    fig.update_yaxes(title="m/s²", tickformat=".2f", gridcolor=GRID, row=2, col=1)
    fig.update_xaxes(zeroline=False)
    _padding_y(fig, grad, 2, arriba=.3)

    v_top = v_tot[-1]
    texto = (
        f"Arriba: V disminuye a medida que el punto se aleja de la Tierra, porque r aumenta. Entre h = 0 y "
        f"h = {h_max / 1000:.0f} km, V baja {_v(v_tot[0] - v_top, 0)} J/kg. La línea discontinua (kM/r) casi coincide "
        f"con V: la forma de la Tierra cambia el potencial muy poco frente al término de masa puntual.\n"
        f"Abajo: la pendiente de la curva de arriba, con el signo cambiado, es la aceleración de la gravitación. En P "
        f"vale {g_p:.4f} m/s²: subir 1 m reduce V en unos {g_p:.2f} J/kg. Esa relación es la que permite pasar de "
        f"diferencias de potencial a diferencias de altura. La gravitación también disminuye con la altura "
        f"(aproximadamente como 1/r²): entre 0 y {h_max / 1000:.0f} km pasa de {grad[0]:.3f} a {grad[-1]:.3f} m/s², "
        f"así que V no baja exactamente en línea recta."
    )
    return _agregar_explicacion(fig, texto, alto_figura=720)


# ---------- 4. Elipsoide 3D: dónde está P y cómo se reparte V ----------

def _superficie_elipsoide(h, n_lat=73, n_lon=121, lon_ini=-180.0, lon_fin=180.0):
    lat = np.radians(np.linspace(-90, 90, n_lat))
    lon = np.radians(np.linspace(lon_ini, lon_fin, n_lon))
    LAT, LON = np.meshgrid(lat, lon, indexing="ij")
    s, c = np.sin(LAT), np.cos(LAT)
    N = MODEL.a / np.sqrt(1 - MODEL.e2 * s**2)
    X = (N + h) * c * np.cos(LON) / 1000
    Y = (N + h) * c * np.sin(LON) / 1000
    Z = ((1 - MODEL.e2) * N + h) * s / 1000
    return np.degrees(lat), LAT, X, Y, Z


def _paralelo(phi_deg, h, n=181, lon_ini=-180.0, lon_fin=180.0):
    lon = np.radians(np.linspace(lon_ini, lon_fin, n))
    p = math.radians(phi_deg)
    N = MODEL.a / math.sqrt(1 - MODEL.e2 * math.sin(p) ** 2)
    rad = (N + h) * math.cos(p) / 1000
    z = ((1 - MODEL.e2) * N + h) * math.sin(p) / 1000
    return rad * np.cos(lon), rad * np.sin(lon), np.full(n, z)


def _punto_elevado(r, factor=1.012):
    """Posición de P ligeramente por fuera de la superficie para que el marcador no quede enterrado."""
    return r.X / 1000 * factor, r.Y / 1000 * factor, r.Z / 1000 * factor


def _escena(lon_camara_deg, elev=0.55, dist=1.5):
    lam = math.radians(lon_camara_deg)
    oculto = dict(visible=False, showspikes=False)
    return dict(
        bgcolor=PAPER, xaxis=oculto, yaxis=oculto, zaxis=oculto, aspectmode="data",
        camera=dict(eye=dict(x=dist * math.cos(lam), y=dist * math.sin(lam), z=elev), up=dict(x=0, y=0, z=1)),
        dragmode="turntable",
    )


LUZ = dict(ambient=.78, diffuse=.45, specular=.05, roughness=.9, fresnel=.05)


def _marcas_polos(h):
    zp = (MODEL.b + h) / 1000 * 1.12
    return go.Scatter3d(
        x=[0, 0], y=[0, 0], z=[zp, -zp], mode="text", text=["<b>Polo Norte</b>", "<b>Polo Sur</b>"],
        textfont=dict(size=12, color=INK), hoverinfo="skip", showlegend=False,
    )


def graph_ellipsoid(r):
    # Se retira una cuña de 90° al oeste de P para ver el centro, el eje y el radio r.
    lam = r.lam_deg
    lon_ini, lon_fin = lam, lam + 270.0
    lats, LAT, X, Y, Z = _superficie_elipsoide(r.h, n_lon=136, lon_ini=lon_ini, lon_fin=lon_fin)
    v_lat = np.array([calculate_model(float(p), lam, r.h).V_total for p in lats])
    V = np.repeat(v_lat[:, None], X.shape[1], axis=1)  # V solo depende de la latitud

    fig = go.Figure()
    fig.add_trace(go.Surface(
        x=X, y=Y, z=Z, surfacecolor=V, lighting=LUZ,
        colorscale=[[0.0, "#27332C"], [0.35, "#5E6B55"], [0.7, "#B7A77F"], [1.0, "#F1E3C4"]],
        colorbar=dict(title=dict(text="V [J/kg]", side="right"), len=.7, thickness=18, x=1.0,
                      tickformat=",.0f"),
        hovertemplate="V = %{surfacecolor:,.2f} J/kg<extra></extra>", showlegend=False,
    ))
    # caras del corte (planos meridianos en λ y λ − 90°)
    t = np.linspace(0, 1, 14)
    for mu in (lon_ini, lon_fin):
        _, _, bx, by, bz = _superficie_elipsoide(r.h, n_lon=2, lon_ini=mu, lon_fin=mu)
        T, BX = np.meshgrid(t, bx[:, 0], indexing="ij")
        _, BY = np.meshgrid(t, by[:, 0], indexing="ij")
        _, BZ = np.meshgrid(t, bz[:, 0], indexing="ij")
        fig.add_trace(go.Surface(
            x=T * BX, y=T * BY, z=T * BZ, surfacecolor=np.zeros_like(T), showscale=False,
            colorscale=[[0, "#E6DACB"], [1, "#E6DACB"]], lighting=LUZ, hoverinfo="skip", showlegend=False,
        ))
        fig.add_trace(go.Scatter3d(x=bx[:, 0], y=by[:, 0], z=bz[:, 0], mode="lines",
                                   line=dict(color="#8C8279", width=3), hoverinfo="skip", showlegend=False))

    # líneas dibujadas sobre la cara del corte que contiene a P (desplazadas un poco hacia el observador)
    off = math.radians(lam - 90)
    dx, dy = 40 * math.cos(off), 40 * math.sin(off)
    zp = (MODEL.b + r.h) / 1000
    fig.add_trace(go.Scatter3d(x=[dx, dx], y=[dy, dy], z=[-zp * 1.06, zp * 1.06], mode="lines",
                               line=dict(color=SLATE, width=4, dash="dash"), name="Eje de rotación",
                               hoverinfo="skip"))
    lr = math.radians(lam)
    ra = (MODEL.a + r.h) / 1000
    fig.add_trace(go.Scatter3d(x=[dx, ra * math.cos(lr) + dx], y=[dy, ra * math.sin(lr) + dy], z=[0, 0],
                               mode="lines", line=dict(color=INK, width=4, dash="dot"),
                               name="Radio ecuatorial (a)", hoverinfo="skip"))
    px, py, pz = r.X / 1000, r.Y / 1000, r.Z / 1000
    fig.add_trace(go.Scatter3d(x=[dx, px + dx], y=[dy, py + dy], z=[0, pz], mode="lines",
                               line=dict(color=ACCENT, width=8),
                               name=f"Radio geocéntrico r = {_v(r.r / 1000, 3)} km",
                               hovertemplate=f"r = {_v(r.r, 3)} m<extra></extra>"))
    ex, ey, ez = _paralelo(0.0, r.h + 12000, lon_ini=lon_ini, lon_fin=lon_fin)
    fig.add_trace(go.Scatter3d(x=ex, y=ey, z=ez, mode="lines", line=dict(color=INK, width=3, dash="dash"),
                               name="Ecuador", hoverinfo="skip"))
    qx, qy, qz = _paralelo(r.phi_deg, r.h + 15000, lon_ini=lon_ini, lon_fin=lon_fin)
    fig.add_trace(go.Scatter3d(x=qx, y=qy, z=qz, mode="lines", line=dict(color=ROSE, width=8),
                               name=f"Paralelo de P: V = {_v(r.V_total)} J/kg",
                               hovertemplate=f"Paralelo de P<br>V = {_v(r.V_total)} J/kg<extra></extra>"))
    mx, my, mz = _punto_elevado(r)
    fig.add_trace(go.Scatter3d(
        x=[mx], y=[my], z=[mz], mode="markers+text", text=["P  "], textposition="middle left",
        textfont=dict(size=16, color=INK), marker=dict(size=8, color=CLAY, line=dict(color="white", width=2)),
        name="Punto P", hovertemplate=(f"Punto P<br>φ = {_lat_txt(r.phi_deg)}<br>λ = {lam:.4f}°<br>"
                                        f"h = {_v(r.h)} m<br>V = {_v(r.V_total, 3)} J/kg<extra></extra>"),
    ))
    fig.add_trace(go.Scatter3d(x=[dx], y=[dy], z=[0], mode="markers+text", text=["Centro de masas  "],
                               textposition="middle left", textfont=dict(size=12, color=INK),
                               marker=dict(size=4, color=INK), hoverinfo="skip", showlegend=False))
    fig.add_trace(_marcas_polos(r.h))

    fig.update_layout(
        title=dict(text="¿Dónde está P y cómo se reparte V sobre la Tierra?", x=.02, xanchor="left",
                   font=dict(size=19)),
        paper_bgcolor=PAPER, separators=". ", font=dict(family="Segoe UI, Arial", color=INK),
        scene=_escena(lam - 32, elev=.95, dist=1.3), margin=dict(l=0, r=10, t=70),
        legend=dict(x=0, y=1, yanchor="top", bgcolor="rgba(255,253,252,.85)", bordercolor=GRID, borderwidth=1),
        uirevision="elipsoide",
    )
    rango = v_lat.max() - v_lat.min()
    psi = math.degrees(math.atan2(r.Z, math.hypot(r.X, r.Y)))
    texto = (
        f"La superficie es el elipsoide WGS84 a la altura de P (h = {_v(r.h, 1)} m) y su color es el valor de V en "
        f"cada lugar. Los colores forman franjas paralelas al ecuador: en este modelo V depende de la latitud y de r, "
        f"pero no de la longitud, así que cambiar λ solo mueve P a lo largo de su paralelo (línea rosada), donde "
        f"V = {_v(r.V_total)} J/kg en todos los puntos. V aumenta del ecuador (oscuro) a los polos (claro) en "
        f"{_v(rango, 0)} J/kg porque la Tierra está achatada y los polos están más cerca del centro.\n"
        f"Se quitó una cuña para ver el interior. La línea azul es el radio geocéntrico r = {_v(r.r / 1000, 3)} km, "
        f"la distancia que entra en kM/r; forma con el plano del ecuador la latitud geocéntrica ψ = {psi:.4f}°, que "
        f"difiere {abs(r.phi_deg - psi):.4f}° de la geodésica φ = {r.phi_deg:.4f}° (solo coinciden en el ecuador y en "
        f"los polos). El achatamiento real (unos 21 km) no se alcanza a ver a esta "
        f"escala. Arrastre para girar."
    )
    return _agregar_explicacion(fig, texto, alto_figura=630, espacio_ejes=8)


# ---------- 5. Globo 3D de cada corrección (valores reales en J/kg) ----------

def graph_harmonic_shapes(r):
    lats, LAT, X, Y, Z = _superficie_elipsoide(r.h, n_lat=73, n_lon=81)
    filas = [calculate_model(float(p), r.lam_deg, r.h) for p in lats]
    capas = [
        ("Corrección A₂ de C.E", np.array([f.V_CE - f.KM_over_r for f in filas]), r.V_CE - r.KM_over_r,
         "Grado 2, efecto directo del achatamiento. Es positiva (rojo) en la franja ecuatorial, donde el abultamiento "
         "del ecuador acerca más masa, y negativa (azul) hacia los polos. Es, con diferencia, la corrección más grande."),
        ("Achatamiento C.A", np.array([f.V_CA for f in filas]), r.V_CA,
         "Grados 3 y 5 (A₃ y A₅). Son los términos que hacen distinto el hemisferio norte del sur (la llamada «forma de "
         "pera» de la Tierra): compare los colores a uno y otro lado del ecuador."),
        ("Asimetría ecuatorial C.A.E", np.array([f.V_CAE for f in filas]), r.V_CAE,
         "Grado 4 (A₄). Forma bandas: suma potencial en el ecuador y en los polos y lo resta en latitudes medias."),
        ("Suma de las tres correcciones", np.array([f.V_total - f.KM_over_r for f in filas]), r.V_total - r.KM_over_r,
         "Es V − kM/r: todo lo que la forma real de la Tierra añade o quita al potencial de una Tierra esférica. "
         "Está dominada por la corrección A₂."),
    ]

    fig = go.Figure()
    trazas_por_capa = []
    for i, (nombre, vals, _, _) in enumerate(capas):
        lim = float(np.max(np.abs(vals))) or 1.0
        C = np.repeat(vals[:, None], X.shape[1], axis=1)
        idx = [len(fig.data)]
        fig.add_trace(go.Surface(
            x=X, y=Y, z=Z, surfacecolor=C, cmin=-lim, cmax=lim, visible=(i == 0), lighting=LUZ,
            colorscale=[[0, "#2E5A87"], [0.25, "#8FB3D1"], [0.5, "#F4EFE8"], [0.75, "#E3A184"], [1, "#9E3B26"]],
            colorbar=dict(title=dict(text="J/kg", side="right"), len=.7, thickness=18, x=1.0, tickformat=",.1f"),
            hovertemplate=nombre + "<br>%{surfacecolor:,.3f} J/kg<extra></extra>", showlegend=False,
        ))
        # paralelos donde la corrección vale cero (separan rojo y azul)
        for k, x0 in enumerate(_cruces_cero(lats, vals)):
            cx, cy, cz = _paralelo(x0, r.h + 3000)
            idx.append(len(fig.data))
            fig.add_trace(go.Scatter3d(
                x=cx, y=cy, z=cz, mode="lines", line=dict(color=INK, width=4), visible=(i == 0),
                name="Paralelo donde la corrección vale 0", showlegend=(k == 0),
                hovertemplate=f"Corrección = 0 en φ = {x0:.2f}°<extra></extra>",
            ))
        trazas_por_capa.append(idx)

    i_fijas = len(fig.data)
    px, py, pz = _punto_elevado(r)
    fig.add_trace(go.Scatter3d(
        x=[px], y=[py], z=[pz], mode="markers+text", text=["  P"], textposition="middle right",
        textfont=dict(size=15, color=INK), marker=dict(size=7, color=CLAY, line=dict(color="white", width=2)),
        name="Punto P", hovertemplate=(
            f"Punto P ({_lat_txt(r.phi_deg)})<br>A₂ de C.E = {_v(r.V_CE - r.KM_over_r, 3)} J/kg<br>"
            f"C.A = {_v(r.V_CA, 3)} J/kg<br>C.A.E = {_v(r.V_CAE, 3)} J/kg<extra></extra>"),
    ))
    fig.add_trace(_marcas_polos(r.h))
    n_total = len(fig.data)

    def textos(i):
        nombre, vals, en_p, desc = capas[i]
        cuerpo = (f"{desc} En P vale {_v(en_p, 3)} J/kg (≈ {_v(en_p / G_REF, 3)} m como altura equivalente); "
                  f"sobre toda la Tierra va de {_v(vals.min(), 1)} a {_v(vals.max(), 1)} J/kg.\n"
                  f"Rojo = la corrección aumenta V; azul = lo disminuye; blanco = casi cero. Las líneas negras son los "
                  f"paralelos donde vale exactamente cero. La forma del globo es la real; lo que cambia es el color. "
                  f"Use el menú para cambiar de corrección y arrastre para girar.")
        return nombre, cuerpo

    alto_base = 680
    botones, anotaciones, alto_max = [], [], 0
    for i in range(len(capas)):
        nombre, cuerpo = textos(i)
        contenido, alto_txt = _bloque_explicacion(cuerpo)
        alto_max = max(alto_max, alto_txt)
        anot = dict(text=contenido, xref="paper", yref="paper", x=0, y=0, xanchor="left", yanchor="top", yshift=-8,
                    showarrow=False, align="left", font=dict(size=12.5, color=INK), bgcolor="#FFFDFC",
                    bordercolor=GRID, borderwidth=1, borderpad=10)
        anotaciones.append(anot)
        visibles = [False] * n_total
        for j in trazas_por_capa[i]:
            visibles[j] = True
        for j in range(i_fijas, n_total):
            visibles[j] = True
        botones.append(dict(label=nombre, method="update", args=[
            {"visible": visibles},
            {"title.text": f"¿Dónde suma y dónde resta cada corrección?  ·  {nombre}", "annotations": [anot]},
        ]))

    fig.update_layout(
        title=dict(text=f"¿Dónde suma y dónde resta cada corrección?  ·  {capas[0][0]}", x=.02, xanchor="left",
                   font=dict(size=19)),
        paper_bgcolor=PAPER, separators=". ", font=dict(family="Segoe UI, Arial", color=INK),
        scene=_escena(r.lam_deg, elev=.35 + .5 * math.sin(math.radians(r.phi_deg))), annotations=[anotaciones[0]],
        updatemenus=[dict(type="dropdown", direction="down", x=0, y=1.0, xanchor="left", yanchor="top",
                          buttons=botones, bgcolor="#FFFDFC", bordercolor=GRID, font=dict(color=INK), active=0)],
        legend=dict(x=0, y=.9, yanchor="top", bgcolor="rgba(255,253,252,.85)", bordercolor=GRID, borderwidth=1),
        height=alto_base + alto_max, margin=dict(l=0, r=10, t=70, b=alto_max + 20),
        uirevision="correcciones",
    )
    return fig


# Textos breves para que la interfaz pueda mostrar qué es cada gráfica.
GRAFICAS_INFO = {
    "profile": "Potencial total a la altura de P de polo a polo: muestra por qué V es mayor en los polos.",
    "contributions": "Cuánto pesa cada término en V y cómo cambia cada corrección con la latitud.",
    "height": "Cómo disminuye V al subir y cómo su pendiente da la gravedad.",
    "ellipsoid": "Ubicación de P en la Tierra, su radio geocéntrico r y el reparto de V por latitudes.",
    "shapes": "Globo con el valor real de cada corrección: dónde suma y dónde resta potencial.",
}


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
<div class="hero"><div class="eyebrow">GEODESIA FÍSICA · WGS84 · INFORME CIENTÍFICO</div><h1>GeoPotencial 6</h1><div>Potencial gravitacional de la Tierra por aportes</div><p>{' · '.join(AUTHORS)}</p>
<div style="margin-top:8px;font-size:13px;opacity:.88">
  Laura Vargas · 20241025003 · Lista 27 &nbsp;|&nbsp;
  Michael Ramirez · 20232025071 · Lista 29 &nbsp;|&nbsp;
  Stephany Vega · 20241025033 · Lista 28
</div></div>
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
        elif kind == "shapes":
            fig = graph_harmonic_shapes(r)
        elif kind == "surface":
            raise ValueError("La superficie 3D V(φ, h) se retiró; use «Potencial vs altura» o «Perfil».")
        else:
            raise ValueError("Gráfica no reconocida.")
        return jsonify({"ok": True, "figure": pio.to_json(fig, pretty=False),
                        "explanation": GRAFICAS_INFO.get(kind, "")})
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
