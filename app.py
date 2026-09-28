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
    Visualización comparativa de las correcciones armónicas del modelo.
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
        height=940,
        paper_bgcolor=PAPER,
        plot_bgcolor=PAPER,
        font=dict(family="Segoe UI, Arial", color=INK),
        # Más espacio arriba y abajo para evitar solapamientos de textos.
        margin=dict(l=68, r=40, t=145, b=105),
        title=dict(
            text="Correcciones armónicas por latitud",
            x=.02, xanchor="left",
            y=.985, yanchor="top",
            font=dict(size=20)
        ),
        bargap=.48,
        showlegend=False,
    )

    # Los títulos automáticos de los subgráficos quedan demasiado cerca
    # del título principal en Plotly. Bajamos únicamente el primero.
    if fig.layout.annotations:
        fig.layout.annotations[0].update(yshift=-34)

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
        text=(
            "Nota de lectura: en C.E se representa solo la corrección A₂. "
            "El término central kM/r se deja fuera de esta comparación para que "
            "C.A y C.A.E puedan apreciarse pese a la diferencia de escala."
        ),
        xref="paper", yref="paper",
        x=0.0, y=-0.095,
        xanchor="left", yanchor="top",
        showarrow=False,
        align="left",
        font=dict(size=10.5, color=SLATE),
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
    """Elipsoide WGS84 con Norte/Sur explícitos y vista fija."""
    lat = np.linspace(-np.pi/2, np.pi/2, 70)
    lon = np.linspace(-np.pi, np.pi, 100)
    LAT, LON = np.meshgrid(lat, lon, indexing="ij")

    sin_lat = np.sin(LAT)
    cos_lat = np.cos(LAT)
    N = MODEL.a / np.sqrt(1 - MODEL.e2 * sin_lat**2)

    X = (N * cos_lat * np.cos(LON)) / 1000.0
    Y = (N * cos_lat * np.sin(LON)) / 1000.0
    Z = ((1 - MODEL.e2) * N * sin_lat) / 1000.0

    V = np.empty_like(LAT, dtype=float)
    for i in range(LAT.shape[0]):
        phi_deg = np.degrees(LAT[i, 0])
        rr = calculate_model(float(phi_deg), r.lam_deg, 0.0)
        V[i, :] = rr.V_total

    px, py, pz = r.X / 1000.0, r.Y / 1000.0, r.Z / 1000.0

    fig = go.Figure()

    fig.add_trace(go.Surface(
        x=X, y=Y, z=Z,
        surfacecolor=V,
        colorscale=[
            [0.0, "#3F493D"],
            [0.45, "#8E8A68"],
            [0.72, "#C39A72"],
            [1.0, "#875744"],
        ],
        colorbar=dict(title="V [J/kg]", len=.68, thickness=20, x=1.03),
        hovertemplate="X=%{x:.1f} km<br>Y=%{y:.1f} km<br>Z=%{z:.1f} km<br>V=%{surfacecolor:.2f} J/kg<extra></extra>",
        showscale=True
    ))

    fig.add_trace(go.Scatter3d(
        x=[px], y=[py], z=[pz],
        mode="markers+text",
        marker=dict(size=6, color=ROSE, line=dict(color="white", width=2)),
        text=["P"],
        textposition="top center",
        name="Punto P",
        hovertemplate=(
            "Punto P<br>"
            f"φ={r.phi_deg:.4f}°<br>"
            f"λ={r.lam_deg:.4f}°<br>"
            f"h={r.h:.2f} m<extra></extra>"
        )
    ))

    pole_z = MODEL.b / 1000.0
    fig.add_trace(go.Scatter3d(
        x=[0, 0], y=[0, 0], z=[pole_z * 1.08, -pole_z * 1.08],
        mode="markers+text",
        marker=dict(size=5, color=[MOSS, CLAY]),
        text=["NORTE (+Z)", "SUR (−Z)"],
        textposition=["top center", "bottom center"],
        hoverinfo="skip",
        showlegend=False
    ))

    fig.update_layout(
        title=dict(
            text="Elipsoide WGS84 coloreado por potencial · Norte (+Z) / Sur (−Z)",
            x=.02, xanchor="left", font=dict(size=20)
        ),
        paper_bgcolor=PAPER,
        font=dict(family="Segoe UI, Arial", color=INK),
        scene=dict(
            bgcolor=PAPER,
            xaxis=dict(title="X [km]", gridcolor=GRID, zerolinecolor=GRID, showspikes=False),
            yaxis=dict(title="Y [km]", gridcolor=GRID, zerolinecolor=GRID, showspikes=False),
            zaxis=dict(title="Z [km] · Norte (+) / Sur (−)", gridcolor=GRID, zerolinecolor=GRID, showspikes=False),
            aspectmode="data",
            camera=dict(
                eye=dict(x=1.45, y=1.45, z=1.05),
                up=dict(x=0, y=0, z=1)
            ),
            dragmode=False
        ),
        margin=dict(l=0, r=85, t=80, b=20),
        showlegend=False,
        uirevision="fixed-wgs84-view"
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
