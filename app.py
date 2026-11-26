"""
PRISMA SpA Constructora — Plataforma PMO de Control de Presupuesto y APU.

Características principales:
- Lectura en una sola pasada con openpyxl (data_only=True, read_only=True) y @st.cache_data
  para archivos .xlsx (incluyendo libros Anexo N° 4 multi-hoja con desglose APU) y respaldo .xls.
- Mapeador dinámico de columnas y detector inteligente para planillas de Presto,
  Arquímedes, MS Project, Primavera P6 o plantillas propias.
- Generador realista de datos de ejemplo (Modo Demo) con PV y EV.
- Dashboard ejecutivo en Streamlit con identidad visual PRISMA SpA Constructora,
  gráficos interactivos Plotly (Dona, Top 15, Pareto 80/20 y Treemap) y
  exportación limpia a Excel (.xlsx).
"""

from __future__ import annotations

import datetime as dt
import io
import math
import re
import unicodedata
from pathlib import Path
from typing import Any

import numpy as np
import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st


# ---------------------------------------------------------------------------
# Identidad visual PRISMA SpA Constructora y configuración de página
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="PRISMA SpA | Control de Presupuesto y APU",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

PRISMA_NAVY = "#153352"
PRISMA_NAVY_DARK = "#0E2238"
PRISMA_GOLD = "#C89236"
PRISMA_SLATE = "#5A6570"
PRISMA_BG = "#F3F4F6"
PRISMA_CARD = "#FFFFFF"
PRISMA_PALETTE = [
    "#153352",
    "#C89236",
    "#3E6B89",
    "#64748B",
    "#9A6B21",
    "#2A4D6E",
    "#8C939A",
]

BASE_DIR = Path(__file__).resolve().parent
ARCHIVO_LOCAL_DEFECTO = BASE_DIR / "7.21.5_Anexo_N°_4_Presupuesto.xlsx"
LOGO_PRISMA_PATH = BASE_DIR / "assets" / "prisma_logo.jpg"

# Palabras clave para detección de encabezados y mapeo dinámico
HEADER_KEYWORDS: dict[str, list[str]] = {
    "codigo": ["codigo", "id", "item", "cod", "nro", "numero", "n°"],
    "partida": [
        "partida",
        "designacion",
        "descripcion",
        "actividad",
        "especificacion",
        "concepto",
        "glosa",
    ],
    "capitulo": ["capitulo", "especialidad", "area", "wbs", "grupo", "rubro"],
    "recurso": ["recurso", "insumo", "material", "descripcion recurso", "detalle"],
    "tipo_recurso": ["tipo", "categoria", "tipo de recurso", "naturaleza", "clase"],
    "unidad": ["unidad", "und", "un.", "un", "u", "um", "u.m."],
    "cantidad": ["cantidad", "cant", "cant.", "metrado", "cubicacion", "rendimiento"],
    "precio_unitario": [
        "precio unitario",
        "precio_unitario",
        "pu",
        "p.u.",
        "p.u",
        "valor unitario",
        "costo unitario",
        "valor",
        "precio",
    ],
    "total": [
        "costo total",
        "costo_total",
        "valor total",
        "total ($)",
        "total",
        "parcial",
        "monto",
        "importe",
    ],
    "fecha": ["fecha", "periodo", "mes", "semana", "fecha inicio", "fecha_inicio"],
    "costo_real": ["costo_real", "costo real", "ev", "ejecutado", "avance real", "real"],
}

FOOTER_SUMMARY_TOKENS = {
    "DIRECTO",
    "COSTO DIRECTO",
    "GG",
    "GASTOS GENERALES",
    "UTIL",
    "UTILIDAD",
    "UTILIDADES",
    "NETO",
    "TOTAL NETO",
    "IVA",
    "I.V.A.",
    "TOTAL",
    "TOTAL GENERAL",
    "GRAN TOTAL",
    "TOTAL PRESUPUESTO",
}

SUBTOTAL_PREFIXES = (
    "SUBTOTAL",
    "SUB TOTAL",
    "SUB-TOTAL",
    "TOTAL CAPITULO",
    "TOTAL PARTIDA",
    "TOTAL HOJA",
)


# ---------------------------------------------------------------------------
# Inyección de estilos CSS corporativos (PRISMA SpA Constructora)
# ---------------------------------------------------------------------------
def aplicar_estilos_prisma() -> None:
    """Aplica estilos corporativos consistentes y elimina espacio superior sobrante."""
    st.markdown(
        f"""
<style>
:root {{
    color-scheme: light;
    --primary-color: {PRISMA_GOLD};
    --background-color: {PRISMA_BG};
    --secondary-background-color: #FFFFFF;
    --text-color: #152536;
}}
.stApp, [data-testid="stAppViewContainer"] {{
    background-color: {PRISMA_BG} !important;
    color: #152536 !important;
}}
[data-testid="stHeader"] {{
    background: transparent !important;
    height: 2.25rem !important;
}}
.main .block-container, [data-testid="stMainBlockContainer"] {{
    padding-top: 1.15rem !important;
    padding-bottom: 1.5rem !important;
}}
section[data-testid="stSidebar"], section[data-testid="stSidebar"] > div {{
    background-color: #FFFFFF !important;
    color: #152536 !important;
    border-right: 1px solid #CBD5E1;
}}
section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {{
    padding-top: 0.75rem !important;
}}
/* Tipografía general en zona clara (sin pisar textos blancos del banner ni tags) */
.stApp h2, .stApp h3, .stApp h4,
.stApp p, .stApp label,
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label {{
    color: #152536;
}}
/* Inputs, selects y file uploader siempre claros y legibles */
[data-baseweb="input"], [data-baseweb="base-input"],
[data-baseweb="input"] input,
[data-baseweb="select"] > div {{
    background-color: #FFFFFF !important;
    color: #152536 !important;
    border-color: #CBD5E1 !important;
}}
[data-baseweb="select"] span, [data-baseweb="select"] svg {{
    color: #152536 !important;
}}
[data-testid="stFileUploaderDropzone"] {{
    background-color: #F8FAFC !important;
    color: #152536 !important;
    border: 1px dashed #94A3B8 !important;
}}
[data-testid="stFileUploaderDropzone"] * {{
    color: #152536 !important;
}}
[data-testid="stFileUploaderDropzone"] button {{
    background-color: #FFFFFF !important;
    color: {PRISMA_NAVY} !important;
    border: 1px solid #CBD5E1 !important;
}}
[data-testid="stExpander"] {{
    background-color: #FFFFFF !important;
    border: 1px solid #DCE3EC !important;
    border-radius: 10px !important;
}}
/* Tags de multiselect con identidad PRISMA */
span[data-baseweb="tag"] {{
    background-color: {PRISMA_NAVY} !important;
    color: #FFFFFF !important;
}}
span[data-baseweb="tag"] span, span[data-baseweb="tag"] svg {{
    color: #FFFFFF !important;
}}
/* Banner corporativo PRISMA */
.prisma-header {{
    background: linear-gradient(135deg, {PRISMA_NAVY_DARK} 0%, {PRISMA_NAVY} 68%, #204A73 100%);
    border-bottom: 4px solid {PRISMA_GOLD};
    border-radius: 12px;
    padding: 0.95rem 1.35rem;
    margin-top: 0 !important;
    margin-bottom: 0.95rem;
    color: #FFFFFF !important;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 1.15rem;
    box-shadow: 0 4px 16px rgba(21, 51, 82, 0.14);
}}
.prisma-header-left {{
    display: flex;
    align-items: center;
    gap: 0.95rem;
}}
.prisma-emblem {{
    background: rgba(255, 255, 255, 0.12);
    border: 1px solid rgba(200, 146, 54, 0.6);
    border-radius: 10px;
    padding: 0.5rem 0.65rem;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}}
.stApp .prisma-header .prisma-eyebrow {{
    font-size: 0.72rem !important;
    font-weight: 700 !important;
    letter-spacing: 1.2px;
    text-transform: uppercase;
    color: #F6D89C !important;
    margin-bottom: 0.1rem;
}}
.stApp .prisma-header h1.prisma-title {{
    margin: 0 !important;
    padding: 0 !important;
    font-size: 1.38rem !important;
    font-weight: 700 !important;
    line-height: 1.2 !important;
    color: #FFFFFF !important;
}}
.stApp .prisma-header p.prisma-subtitle,
.stApp .prisma-header p.prisma-subtitle b {{
    margin: 0.25rem 0 0 0 !important;
    font-size: 0.84rem !important;
    color: #E2E8F0 !important;
}}
.stApp .prisma-header .prisma-badge {{
    background: rgba(200, 146, 54, 0.22);
    border: 1px solid {PRISMA_GOLD};
    color: #F6D89C !important;
    padding: 0.35rem 0.8rem;
    border-radius: 999px;
    font-size: 0.76rem;
    font-weight: 700;
    white-space: nowrap;
}}
/* Tarjetas KPI sin truncamiento y delta legible */
div[data-testid="stMetric"] {{
    background-color: {PRISMA_CARD} !important;
    border: 1px solid #DCE3EC;
    border-top: 3px solid {PRISMA_NAVY};
    border-radius: 10px;
    padding: 0.75rem 0.9rem;
    box-shadow: 0 2px 8px rgba(15, 23, 42, 0.05);
}}
div[data-testid="stMetricLabel"] p {{
    font-weight: 600 !important;
    color: {PRISMA_SLATE} !important;
    font-size: 0.72rem !important;
    line-height: 1.12 !important;
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    min-height: 2.2em;
}}
div[data-testid="stMetricValue"] > div {{
    color: {PRISMA_NAVY} !important;
    font-weight: 700 !important;
    font-size: clamp(1.1rem, 1.45vw, 1.46rem) !important;
    line-height: 1.25 !important;
    overflow: visible !important;
    text-overflow: unset !important;
    white-space: nowrap !important;
}}
div[data-testid="stMetricDelta"] {{
    background-color: #EEF2F6 !important;
    color: {PRISMA_NAVY} !important;
    padding: 0.15rem 0.5rem !important;
    border-radius: 999px !important;
}}
div[data-testid="stMetricDelta"] > div,
div[data-testid="stMetricDelta"] svg {{
    color: {PRISMA_NAVY} !important;
}}
button[data-baseweb="tab"] {{
    font-weight: 600 !important;
    color: {PRISMA_SLATE} !important;
    white-space: normal !important;
    text-align: center !important;
}}
button[data-baseweb="tab"][aria-selected="true"] {{
    color: {PRISMA_NAVY} !important;
    border-bottom-color: {PRISMA_GOLD} !important;
}}
</style>
        """.strip(),
        unsafe_allow_html=True,
    )


def render_banner_corporativo(meta_texto: str, modo_badge: str) -> None:
    """Muestra la cabecera corporativa sin indentación HTML (evita bloques de código en Markdown)."""
    svg_emblema = (
        '<div class="prisma-emblem">'
        '<svg width="36" height="36" viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg">'
        '<path d="M6 31V11L18 4L30 11V31H24V15L18 11.5L12 15V31H6Z" fill="#FFFFFF" fill-opacity="0.92"/>'
        '<path d="M18 4L30 11V14.5L18 7.5V4Z" fill="#C89236"/>'
        '<path d="M15 31V19H21V24H17.5V31H15Z" fill="#C89236"/>'
        '</svg>'
        '</div>'
    )
    html_banner = (
        f'<div class="prisma-header">'
        f'<div class="prisma-header-left">'
        f'{svg_emblema}'
        f'<div>'
        f'<div class="prisma-eyebrow">PRISMA SpA Constructora · Plataforma PMO</div>'
        f'<h1 class="prisma-title">Control Ejecutivo · Presupuesto y Análisis de APU</h1>'
        f'<p class="prisma-subtitle">{meta_texto}</p>'
        f'</div>'
        f'</div>'
        f'<div class="prisma-badge">{modo_badge}</div>'
        f'</div>'
    )
    st.markdown(html_banner, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Funciones auxiliares de normalización, unidades y parsing seguro
# ---------------------------------------------------------------------------
def normalizar_texto(valor: Any) -> str:
    """Convierte texto a minúsculas sin tildes ni espacios múltiples."""
    if valor is None:
        return ""
    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", texto)


def normalizar_unidad_estandar(valor: Any) -> str:
    """Normaliza unidades de medida evitando duplicados (m2/m² -> M2, U/UN/UNI -> UN)."""
    if valor is None:
        return ""
    raw = str(valor).strip()
    if not raw or raw.lower() in {"none", "nan", "-"}:
        return ""
    norm = normalizar_texto(raw).upper().replace("²", "2").replace("³", "3").replace(".", "").strip()
    compact = re.sub(r"\s+", "", norm)
    if compact in {"M2", "MT2", "MTS2", "METROCUADRADO", "METROSCUADRADOS"}:
        return "M2"
    if compact in {"M3", "MT3", "MTS3"}:
        return "M3"
    if compact in {"GL", "GLB", "GLOBAL"}:
        return "GL"
    if compact in {"ML", "MTL"}:
        return "ML"
    if compact in {"KG", "KGS", "KILO", "KILOS"}:
        return "KG"
    if compact in {"U", "UN", "UND", "UNI", "UNIDAD", "UNIDADES"}:
        return "UN"
    if compact in {"DIA", "DIAS", "DÍA", "DÍAS"}:
        return "DÍA"
    if compact in {"HH", "HR", "HRS", "HORA", "HORAS"}:
        return "HH"
    return norm


def normalizar_tipo_recurso(valor: Any, codigo_recurso: str = "") -> str:
    """Agrupa los tipos de recurso en categorías estándar de control PMO."""
    norm = normalizar_texto(valor)
    cod = str(codigo_recurso).strip().upper()

    if any(k in norm for k in ("mano de obra", "mano obra", "mo", "jornal", "operario", "maestro", "leyes sociales")):
        return "Mano de Obra"
    if any(k in norm for k in ("material", "insumo", "mat")):
        return "Materiales"
    if any(k in norm for k in ("maquin", "equipo", "herramienta", "maq", "arriendo")):
        return "Maquinaria / Equipos"
    if any(k in norm for k in ("subcontrat", "servicio", "tercero", "sub")):
        return "Subcontratos"

    if cod.startswith("MAT"):
        return "Materiales"
    if cod.startswith("MO"):
        return "Mano de Obra"
    if cod.startswith("MAQ") or cod.startswith("EQ"):
        return "Maquinaria / Equipos"
    if cod.startswith("SUB"):
        return "Subcontratos"

    return str(valor).strip() if valor is not None and str(valor).strip() else "Otros / Partida Directa"


def parsear_numero(valor: Any) -> float | None:
    """Convierte de forma segura celdas numéricas o texto con formato moneda a float."""
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float, np.number)):
        fval = float(valor)
        return None if (math.isnan(fval) or math.isinf(fval)) else fval

    texto = str(valor).strip()
    if not texto:
        return None

    limpio = texto.replace("$", "").replace("CLP", "").replace("USD", "").strip()
    if limpio in {"-", "—", "–", "N/A", "NA", "S/I", "None", "nan", ""}:
        return None

    tiene_moneda = "$" in texto or "CLP" in texto.upper() or "USD" in texto.upper()
    negativo = False
    if limpio.startswith("(") and limpio.endswith(")"):
        negativo = True
        limpio = limpio[1:-1].strip()

    candidato = re.sub(r"[^\d.,\-]", "", limpio)
    if not candidato or candidato in {"-", ".", ","}:
        return None
    if candidato.startswith("-"):
        negativo = True
        candidato = candidato[1:]
    candidato = candidato.replace("-", "")
    if not candidato:
        return None

    if "." in candidato and "," in candidato:
        if candidato.rfind(",") > candidato.rfind("."):
            candidato = candidato.replace(".", "").replace(",", ".")
        else:
            candidato = candidato.replace(",", "")
    elif "," in candidato:
        candidato = candidato.replace(",", "") if candidato.count(",") > 1 else candidato.replace(",", ".")
    elif "." in candidato:
        partes = candidato.split(".")
        if len(partes) > 2 or (len(partes) == 2 and len(partes[1]) == 3 and tiene_moneda):
            candidato = "".join(partes)

    try:
        numero = float(candidato)
        return -numero if negativo else numero
    except ValueError:
        return None


def normalizar_item(valor: Any, capitulo_actual: str = "") -> str:
    """Normaliza códigos de ítem (1.0 -> '1') y recupera códigos convertidos a fecha por Excel (10.1)."""
    if valor is None:
        return ""
    if isinstance(valor, (dt.datetime, dt.date, pd.Timestamp)):
        cap_num = capitulo_actual.split(" - ")[0].strip() if capitulo_actual else ""
        dia_mes = f"{valor.day}.{valor.month}"
        mes_dia = f"{valor.month}.{valor.day}"
        if cap_num and str(valor.month) == cap_num and str(valor.day) != cap_num:
            return mes_dia
        return dia_mes
    if isinstance(valor, (int, np.integer)):
        return str(int(valor))
    if isinstance(valor, (float, np.floating)):
        if math.isnan(float(valor)) or math.isinf(float(valor)):
            return ""
        return str(int(valor)) if float(valor).is_integer() else f"{float(valor):g}"
    texto = str(valor).strip()
    return texto.split(".")[0] if re.fullmatch(r"\d+\.0+", texto) else texto


def formato_moneda(valor: float | None, decimales: int = 0) -> str:
    """Formatea un valor numérico como moneda legible con separador de miles."""
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return "$ 0"
    texto = f"{valor:,.{decimales}f}"
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"$ {texto}"


def formato_numero(valor: float | None, decimales: int = 2) -> str:
    """Formatea números con separador de miles '.' y decimal ','."""
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return "0"
    texto = f"{valor:,.{decimales}f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def acortar_etiqueta(texto: str, max_len: int = 40) -> str:
    """Acorta etiquetas largas manteniendo claridad en gráficos."""
    texto = re.sub(r"\s+", " ", str(texto)).strip()
    return texto if len(texto) <= max_len else texto[: max_len - 1].rstrip() + "…"


# ---------------------------------------------------------------------------
# Lectura optimizada en una sola pasada (@st.cache_data)
# ---------------------------------------------------------------------------
def seleccionar_hoja_por_defecto(sheetnames: list[str]) -> str:
    """Selecciona la hoja principal de presupuesto si existe."""
    prioridad = ["PPTO MODIFICADO", "PRESUPUESTO", "PPTO", "RESUMEN", "ANEXO"]
    for clave in prioridad:
        for sheet in sheetnames:
            if clave in normalizar_texto(sheet).upper():
                return sheet
    return sheetnames[0]


def detectar_hojas_apu(sheetnames: list[str]) -> set[str]:
    """Identifica hojas de Análisis de Precios Unitarios cuyo nombre es un código (ej. '1.1.1', '2.9.1')."""
    patron = re.compile(r"^\d+(\.\d+)+$")
    return {s for s in sheetnames if patron.fullmatch(s.strip())}


@st.cache_data(show_spinner=False)
def cargar_libro_y_hoja_unica(
    file_bytes: bytes, nombre_archivo: str, hoja_objetivo: str | None = None
) -> dict[str, Any]:
    """
    Lee el libro Excel en UNA sola apertura (evitando duplicar tiempos en archivos >10 MB).
    Extrae:
      - sheetnames
      - hoja_activa
      - filas de la hoja seleccionada
      - si existen hojas APU vinculadas, extrae también su desglose de recursos en la misma pasada.
    """
    ext = Path(nombre_archivo).suffix.lower()

    if ext == ".xls":
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="xlrd")
        sheetnames = list(xls.sheet_names)
        if not sheetnames:
            raise ValueError("El archivo .xls no contiene hojas.")
        hoja = hoja_objetivo if hoja_objetivo in sheetnames else seleccionar_hoja_por_defecto(sheetnames)
        df_raw = xls.parse(hoja, header=None)
        filas = [tuple(None if pd.isna(v) else v for v in row) for row in df_raw.itertuples(index=False, name=None)]
        return {
            "sheetnames": sheetnames,
            "hoja": hoja,
            "filas": filas,
            "apu_sheets": {},
            "tiene_estructura_apu": False,
        }

    wb = None
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        sheetnames = list(wb.sheetnames)
        if not sheetnames:
            raise ValueError("El libro Excel no contiene hojas visibles.")

        hoja = hoja_objetivo if hoja_objetivo in sheetnames else seleccionar_hoja_por_defecto(sheetnames)
        ws = wb[hoja]
        filas = [tuple(row) for row in ws.iter_rows(values_only=True)]

        hojas_apu_nombres = detectar_hojas_apu(sheetnames)
        tiene_apu = len(hojas_apu_nombres) >= 5 and hoja == seleccionar_hoja_por_defecto(sheetnames)
        apu_sheets: dict[str, list[tuple[Any, ...]]] = {}

        if tiene_apu:
            for apu_name in hojas_apu_nombres:
                ws_apu = wb[apu_name]
                apu_sheets[apu_name] = [tuple(r) for r in ws_apu.iter_rows(max_row=80, values_only=True)]

        return {
            "sheetnames": sheetnames,
            "hoja": hoja,
            "filas": filas,
            "apu_sheets": apu_sheets,
            "tiene_estructura_apu": tiene_apu,
        }
    finally:
        if wb is not None:
            wb.close()


# ---------------------------------------------------------------------------
# Detección de encabezados, metadatos y mapeo dinámico
# ---------------------------------------------------------------------------
def detectar_fila_encabezado(filas: list[tuple[Any, ...]], max_scan: int = 60) -> int:
    """Detecta la fila (1-based) de encabezados según palabras clave de control de proyectos."""
    mejor_fila = 1
    mejor_puntaje = -1.0

    for idx in range(min(len(filas), max_scan)):
        celdas = [normalizar_texto(c) for c in filas[idx]]
        no_vacias = [c for c in celdas if c]
        if len(no_vacias) < 3:
            continue

        roles: set[str] = set()
        for celda in no_vacias:
            for rol, keywords in HEADER_KEYWORDS.items():
                if any(celda == kw or (len(kw) >= 4 and kw in celda) for kw in keywords):
                    roles.add(rol)

        puntaje = float(len(roles))
        if "partida" in roles:
            puntaje += 0.6
        if "total" in roles:
            puntaje += 0.5
        if "codigo" in roles:
            puntaje += 0.4
        if "unidad" in roles:
            puntaje += 0.3

        if puntaje > mejor_puntaje and len(roles) >= 3:
            mejor_puntaje = puntaje
            mejor_fila = idx + 1

    return mejor_fila


def extraer_metadatos_proyecto(filas: list[tuple[Any, ...]], header_row_1based: int) -> dict[str, str]:
    """Extrae metadatos de cabecera institucional sobre la tabla."""
    metadatos: dict[str, str] = {}
    claves = {
        "proyecto": "Proyecto",
        "ubicacion": "Ubicación",
        "contratista": "Contratista",
        "mandante": "Mandante",
        "obra": "Obra",
        "fecha": "Fecha",
    }
    for fila in filas[: max(0, header_row_1based - 1)]:
        valores = [v for v in fila if v is not None and str(v).strip() != ""]
        if len(valores) < 2:
            continue
        etiqueta = normalizar_texto(valores[0]).rstrip(":").strip()
        for k, label in claves.items():
            if etiqueta == k or etiqueta.startswith(k):
                val = valores[1]
                metadatos[label] = val.strftime("%d-%m-%Y") if isinstance(val, (dt.datetime, dt.date)) else str(val).strip()
                break
    return metadatos


def construir_opciones_columnas(
    fila_encabezado: tuple[Any, ...], max_cols: int
) -> tuple[list[str], dict[str, int]]:
    """Genera nombres únicos por columna y su índice 0-based."""
    etiquetas: list[str] = []
    mapa_idx: dict[str, int] = {}
    conteo: dict[str, int] = {}

    for col_idx in range(max_cols):
        val = fila_encabezado[col_idx] if col_idx < len(fila_encabezado) else None
        nombre = str(val).strip() if val is not None and str(val).strip() else ""
        if not nombre:
            nombre = f"Columna {openpyxl.utils.get_column_letter(col_idx + 1)}"

        if nombre in conteo:
            conteo[nombre] += 1
            nombre_unico = f"{nombre} ({conteo[nombre]})"
        else:
            conteo[nombre] = 1
            nombre_unico = nombre
        etiquetas.append(nombre_unico)
        mapa_idx[nombre_unico] = col_idx

    return etiquetas, mapa_idx


def sugerir_mapeo_columnas(etiquetas: list[str]) -> dict[str, str | None]:
    """Sugiere automáticamente las columnas clave del presupuesto o programación."""
    sugerencias: dict[str, str | None] = {rol: None for rol in HEADER_KEYWORDS}
    usadas: set[str] = set()

    for rol, keywords in HEADER_KEYWORDS.items():
        mejor_col = None
        mejor_score = 0
        for etiqueta in etiquetas:
            if etiqueta in usadas:
                continue
            norm = normalizar_texto(etiqueta)
            if not norm or norm.startswith("columna "):
                continue
            for prioridad, kw in enumerate(keywords):
                score = 0
                if norm == kw:
                    score = 100 - prioridad
                elif len(kw) >= 4 and kw in norm:
                    score = 60 - prioridad
                if score > mejor_score:
                    mejor_score = score
                    mejor_col = etiqueta
        if mejor_col is not None:
            sugerencias[rol] = mejor_col
            usadas.add(mejor_col)

    return sugerencias


def detectar_columnas_periodos_matriciales(etiquetas: list[str], usadas: set[str]) -> list[str]:
    """Detecta columnas matriciales de periodos (ej. 'Mes 1', 'Ene-26', 'Sem 01', '2026-03')."""
    patron_periodo = re.compile(
        r"^(mes\s*\d+|sem(ana)?\s*\d+|p\d+|(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)[-/\s]?\d{2,4}|\d{4}[-/]\d{2})$",
        re.IGNORECASE,
    )
    cols_periodo: list[str] = []
    for et in etiquetas:
        if et in usadas:
            continue
        norm = normalizar_texto(et)
        if patron_periodo.match(norm):
            cols_periodo.append(et)
    return cols_periodo


# ---------------------------------------------------------------------------
# Extracción de recursos desde hojas APU (Anexo N° 4) y tabla plana
# ---------------------------------------------------------------------------
def parsear_hoja_apu(
    codigo_partida: str,
    nombre_partida: str,
    capitulo: str,
    unidad_partida: str,
    cantidad_partida: float,
    total_partida: float,
    filas_apu: list[tuple[Any, ...]],
) -> list[dict[str, Any]]:
    """
    Extrae los insumos (A- Materiales, B- Maquinas, C- Mano de Obra, D- Subcontratos)
    desde una hoja APU y escala sus cantidades y montos a la cantidad total de la partida.
    """
    recursos: list[dict[str, Any]] = []
    tipo_actual = "Materiales"
    en_seccion_recursos = False
    total_apu_acumulado = 0.0

    for fila in filas_apu:
        celda_cod = fila[1] if len(fila) > 1 else None
        celda_desc = fila[2] if len(fila) > 2 else None
        celda_rend = fila[3] if len(fila) > 3 else None
        celda_unit = fila[4] if len(fila) > 4 else None
        celda_perd = fila[5] if len(fila) > 5 else None
        celda_pu = fila[6] if len(fila) > 6 else None
        celda_tot = fila[7] if len(fila) > 7 else None

        texto_cod = str(celda_cod).strip() if celda_cod is not None else ""
        norm_cod = normalizar_texto(texto_cod)

        # Cambio de sección APU (solo tras entrar a A- Materiales / B- / C- / D-)
        if norm_cod.startswith(("a-", "a -")):
            tipo_actual = "Materiales"
            en_seccion_recursos = True
            continue
        if norm_cod.startswith(("b-", "b -")):
            tipo_actual = "Maquinaria / Equipos"
            en_seccion_recursos = True
            continue
        if norm_cod.startswith(("c-", "c -")):
            tipo_actual = "Mano de Obra"
            en_seccion_recursos = True
            continue
        if norm_cod.startswith(("d-", "d -")):
            tipo_actual = "Subcontratos"
            en_seccion_recursos = True
            continue

        if not en_seccion_recursos:
            continue

        # Saltar encabezados internos, subtotales y pie de firma
        if norm_cod in {"codigo", "item"}:
            continue
        if celda_pu and "subtotal" in normalizar_texto(celda_pu):
            continue
        if "total costo unitario" in norm_cod or "firma" in norm_cod:
            break

        tot_unit_apu = parsear_numero(celda_tot)
        if tot_unit_apu is None or tot_unit_apu == 0:
            continue

        desc_rec = re.sub(r"\s+", " ", str(celda_desc)).strip() if celda_desc is not None else ""
        # Caso especial: Leyes Sociales en Mano de Obra (viene en celda_perd cuando desc está vacía)
        if not desc_rec and celda_perd is not None and isinstance(celda_perd, str):
            desc_rec = celda_perd.strip()
        if not desc_rec:
            continue

        rend = parsear_numero(celda_rend)
        perd = parsear_numero(celda_perd)
        pu_insumo = parsear_numero(celda_pu)
        unit_rec = normalizar_unidad_estandar(celda_unit) or ("%" if "leyes" in normalizar_texto(desc_rec) else "UN")

        if rend is not None and rend > 0:
            factor_perd = perd if (perd is not None and perd > 0) else 1.0
            cant_total_recurso = rend * factor_perd * cantidad_partida
        else:
            cant_total_recurso = cantidad_partida

        monto_total_recurso = tot_unit_apu * cantidad_partida
        pu_efectivo = (
            pu_insumo
            if (pu_insumo is not None and pu_insumo > 1 and "leyes" not in normalizar_texto(desc_rec))
            else (monto_total_recurso / cant_total_recurso if cant_total_recurso > 0 else tot_unit_apu)
        )

        total_apu_acumulado += monto_total_recurso
        recursos.append(
            {
                "Código": codigo_partida,
                "Partida": nombre_partida,
                "Capítulo": capitulo,
                "Unidad_Partida": unidad_partida,
                "Cantidad_Partida": cantidad_partida,
                "Código_Recurso": texto_cod,
                "Recurso": desc_rec,
                "Tipo_Recurso": tipo_actual,
                "Unidad": unit_rec,
                "Cantidad": cant_total_recurso,
                "Precio_Unitario": pu_efectivo,
                "Total": monto_total_recurso,
                "Fecha": pd.NaT,
                "Costo_Real": np.nan,
            }
        )

    # Ajuste proporcional mínimo si hubiera diferencia de redondeo con el total oficial de la partida
    if recursos and total_apu_acumulado > 0 and abs(total_apu_acumulado - total_partida) > 1.0:
        factor_ajuste = total_partida / total_apu_acumulado
        # Solo reescalar si la discrepancia es por redondeo o discrepancia menor (<15%), si es mayor mantener respaldo
        if 0.85 <= factor_ajuste <= 1.15:
            for r in recursos:
                r["Total"] *= factor_ajuste
        else:
            return []

    return recursos


def es_fila_resumen_pie(
    item_raw: Any, desc_raw: Any, cant_raw: Any, unit_raw: Any, pu_raw: Any
) -> tuple[bool, str]:
    """Detecta filas resumen de pie de presupuesto evitando falsos positivos sobre partidas válidas."""
    tiene_unidad = unit_raw is not None and str(unit_raw).strip() != ""
    cant_num = parsear_numero(cant_raw)
    pu_num = parsear_numero(pu_raw)

    # Si tiene unidad y cantidad numérica, es una partida real aunque su nombre contenga 'Total'
    if tiene_unidad and (cant_num is not None or pu_num is not None):
        return False, ""

    for candidato in (desc_raw, cant_raw, item_raw):
        if candidato is None:
            continue
        norm = normalizar_texto(candidato).upper().rstrip(":").strip()
        if norm in FOOTER_SUMMARY_TOKENS or any(norm.startswith(pref) for pref in SUBTOTAL_PREFIXES):
            return True, norm

    if item_raw is None and desc_raw is None and unit_raw is None and isinstance(cant_raw, str):
        norm_cant = normalizar_texto(cant_raw).upper().strip()
        if norm_cant and parsear_numero(norm_cant) is None:
            return True, norm_cant

    return False, ""


def procesar_datos_excel(
    filas: list[tuple[Any, ...]],
    header_row_1based: int,
    col_indices: dict[str, int | None],
    etiquetas_cols: list[str],
    mapa_col_idx: dict[str, int],
    apu_sheets: dict[str, list[tuple[Any, ...]]],
    expandir_apu: bool,
    cols_periodo_matricial: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    """
    Transforma la hoja Excel en un DataFrame estándar de Control de Proyectos:
    [Código, Partida, Capítulo, Recurso, Tipo_Recurso, Unidad, Cantidad,
     Precio_Unitario, Total, Fecha, Costo_Real].
    """
    def celda(fila: tuple[Any, ...], idx: int | None) -> Any:
        return fila[idx] if (idx is not None and 0 <= idx < len(fila)) else None

    idx_cod = col_indices.get("codigo")
    idx_part = col_indices.get("partida")
    idx_cap = col_indices.get("capitulo")
    idx_rec = col_indices.get("recurso")
    idx_tipo = col_indices.get("tipo_recurso")
    idx_unit = col_indices.get("unidad")
    idx_cant = col_indices.get("cantidad")
    idx_pu = col_indices.get("precio_unitario")
    idx_tot = col_indices.get("total")
    idx_fecha = col_indices.get("fecha")
    idx_real = col_indices.get("costo_real")

    filas_partidas: list[dict[str, Any]] = []
    filas_recursos_expandidos: list[dict[str, Any]] = []
    resumen_pie: list[dict[str, Any]] = []
    filas_con_formula_vacia = 0

    capitulo_actual = "General"

    for num_fila, fila in enumerate(filas[header_row_1based:], start=header_row_1based + 1):
        vals_check = [
            celda(fila, i)
            for i in (idx_cod, idx_part, idx_cap, idx_rec, idx_unit, idx_cant, idx_pu, idx_tot)
            if i is not None
        ]
        if not any(v is not None and str(v).strip() != "" for v in vals_check):
            continue

        raw_cod = celda(fila, idx_cod)
        raw_part = celda(fila, idx_part)
        raw_cap = celda(fila, idx_cap)
        raw_rec = celda(fila, idx_rec)
        raw_tipo = celda(fila, idx_tipo)
        raw_unit = celda(fila, idx_unit)
        raw_cant = celda(fila, idx_cant)
        raw_pu = celda(fila, idx_pu)
        raw_tot = celda(fila, idx_tot)
        raw_fecha = celda(fila, idx_fecha)
        raw_real = celda(fila, idx_real)

        es_pie, etiqueta_pie = es_fila_resumen_pie(raw_cod, raw_part, raw_cant, raw_unit, raw_pu)
        if es_pie:
            tot_pie = parsear_numero(raw_tot)
            pu_pie = parsear_numero(raw_pu)
            if tot_pie is not None or pu_pie is not None:
                resumen_pie.append(
                    {
                        "Concepto": etiqueta_pie,
                        "Factor": pu_pie,
                        "Monto": tot_pie if tot_pie is not None else pu_pie,
                    }
                )
            continue

        cod_str = normalizar_item(raw_cod, capitulo_actual)
        part_str = re.sub(r"\s+", " ", str(raw_part)).strip() if raw_part is not None else ""
        cap_val = re.sub(r"\s+", " ", str(raw_cap)).strip() if raw_cap is not None else ""
        rec_str = re.sub(r"\s+", " ", str(raw_rec)).strip() if raw_rec is not None else ""
        unit_str = normalizar_unidad_estandar(raw_unit)

        cant_val = parsear_numero(raw_cant)
        pu_val = parsear_numero(raw_pu)
        tot_val = parsear_numero(raw_tot)
        real_val = parsear_numero(raw_real)

        # Si faltan ambos valores numéricos y código/unidad están vacíos -> nota al pie
        if not cod_str and not unit_str and cant_val is None and pu_val is None and tot_val is None:
            continue

        if cant_val is not None and pu_val is not None and tot_val is None:
            tot_val = cant_val * pu_val
        elif cant_val is not None and pu_val is None and tot_val is None:
            filas_con_formula_vacia += 1

        # Detección de capítulos jerárquicos cuando no existe columna Capítulo explícita
        profundidad = len(cod_str.split(".")) if re.fullmatch(r"\d+(\.\d+)*", cod_str) else 0
        tiene_valores = any(v is not None for v in (cant_val, pu_val, tot_val)) or bool(unit_str)

        if not cap_val:
            if not tiene_valores and part_str:
                if profundidad == 1 or not cod_str:
                    capitulo_actual = f"{cod_str} - {part_str}" if cod_str else part_str
                continue
            elif profundidad == 1 and part_str:
                capitulo_actual = f"{cod_str} - {part_str}"
            cap_efectivo = capitulo_actual
        else:
            cap_efectivo = cap_val
            if not tiene_valores:
                continue

        # Parseo seguro de fecha si está mapeada
        fecha_dt = pd.NaT
        if raw_fecha is not None and str(raw_fecha).strip() != "":
            fecha_dt = pd.to_datetime(raw_fecha, errors="coerce", dayfirst=True)

        cant_f = cant_val if cant_val is not None else 0.0
        pu_f = pu_val if pu_val is not None else 0.0
        tot_f = tot_val if tot_val is not None else 0.0
        nombre_partida = part_str or rec_str or f"Ítem {cod_str}"

        # Guardar registro base a nivel partida
        registro_base = {
            "Código": cod_str or f"F-{num_fila}",
            "Partida": nombre_partida,
            "Capítulo": cap_efectivo,
            "Unidad_Partida": unit_str or "UN",
            "Cantidad_Partida": cant_f,
            "Código_Recurso": cod_str,
            "Recurso": rec_str or nombre_partida,
            "Tipo_Recurso": normalizar_tipo_recurso(raw_tipo, cod_str),
            "Unidad": unit_str or "UN",
            "Cantidad": cant_f,
            "Precio_Unitario": pu_f,
            "Total": tot_f,
            "Fecha": fecha_dt,
            "Costo_Real": real_val if real_val is not None else np.nan,
        }

        # Si hay columnas matriciales de periodos, expandir por periodo
        if cols_periodo_matricial:
            agregado_matricial = False
            for col_p in cols_periodo_matricial:
                val_p = parsear_numero(celda(fila, mapa_col_idx[col_p]))
                if val_p is not None and val_p != 0:
                    reg_p = dict(registro_base)
                    reg_p["Total"] = val_p
                    reg_p["Fecha"] = col_p
                    filas_partidas.append(reg_p)
                    agregado_matricial = True
            if not agregado_matricial:
                filas_partidas.append(registro_base)
        else:
            filas_partidas.append(registro_base)

        # Expansión APU si el libro trae hojas APU por código de partida
        if expandir_apu and cod_str in apu_sheets and cant_f > 0 and tot_f > 0:
            recs_apu = parsear_hoja_apu(
                codigo_partida=cod_str,
                nombre_partida=nombre_partida,
                capitulo=cap_efectivo,
                unidad_partida=unit_str or "UN",
                cantidad_partida=cant_f,
                total_partida=tot_f,
                filas_apu=apu_sheets[cod_str],
            )
            if recs_apu:
                filas_recursos_expandidos.extend(recs_apu)
            else:
                filas_recursos_expandidos.append(registro_base)
        elif expandir_apu:
            filas_recursos_expandidos.append(registro_base)

    datos_finales = filas_recursos_expandidos if (expandir_apu and filas_recursos_expandidos) else filas_partidas
    df = pd.DataFrame(datos_finales)
    df_pie = pd.DataFrame(resumen_pie)

    if df.empty:
        return df, df_pie, {"formulas_vacias": filas_con_formula_vacia, "expandido_apu": False}

    # Filtrar posibles filas padre que son subtotales duplicados (solo en tablas no expandidas)
    if not (expandir_apu and filas_recursos_expandidos):
        codigos = [c for c in df["Código"].tolist() if re.fullmatch(r"\d+(\.\d+)*", str(c))]
        es_padre_sub = []
        for _, r in df.iterrows():
            c = str(r["Código"])
            sin_cant = r["Cantidad"] == 0 and r["Precio_Unitario"] == 0
            es_padre_sub.append(
                bool(c and sin_cant and r["Total"] > 0 and any(x != c and x.startswith(f"{c}.") for x in codigos))
            )
        df = df[~pd.Series(es_padre_sub, index=df.index)].reset_index(drop=True)

    info_extra = {
        "formulas_vacias": filas_con_formula_vacia,
        "expandido_apu": bool(expandir_apu and filas_recursos_expandidos),
        "partidas_apu_vinculadas": len(
            {r["Código"] for r in filas_recursos_expandidos if r["Código"] in apu_sheets}
        )
        if expandir_apu
        else 0,
    }
    return df, df_pie, info_extra


# ---------------------------------------------------------------------------
# Generador de Datos de Prueba (Modo Demo realista)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def generar_datos_demo() -> pd.DataFrame:
    """Genera en memoria un proyecto de edificación completo con fechas, recursos, PV y EV."""
    partidas_demo = [
        # (Código, Capítulo, Partida, Und_Partida, Cant_Part, Fecha, Recurso, Tipo, Und_Rec, Cant_Rec, PU, Factor_Real)
        ("1.1", "01. Obras Preliminares", "Instalación de faenas y cierro", "GL", 1, "2026-03-10", "Cierro perimetral OSB y malla", "Materiales", "ML", 140, 18500, 1.02),
        ("1.1", "01. Obras Preliminares", "Instalación de faenas y cierro", "GL", 1, "2026-03-15", "Cuadrilla obras preliminares", "Mano de Obra", "DÍA", 18, 42000, 1.05),
        ("1.2", "01. Obras Preliminares", "Trazado, niveles y replanteo", "M2", 950, "2026-03-22", "Estación total y topógrafo", "Subcontratos", "GL", 1, 950000, 1.00),
        ("2.1", "02. Obra Gruesa", "Excavación masiva y fundaciones", "M3", 420, "2026-04-05", "Retroexcavadora 90 HP", "Maquinaria / Equipos", "HR", 65, 58000, 1.08),
        ("2.1", "02. Obra Gruesa", "Excavación masiva y fundaciones", "M3", 420, "2026-04-12", "Jornal excavación manual", "Mano de Obra", "DÍA", 45, 34000, 1.12),
        ("2.2", "02. Obra Gruesa", "Hormigón H30 fundaciones y pilares", "M3", 285, "2026-04-25", "Hormigón H30 premezclado", "Materiales", "M3", 295, 118000, 1.03),
        ("2.2", "02. Obra Gruesa", "Hormigón H30 fundaciones y pilares", "M3", 285, "2026-05-08", "Bomba telescópica hormigón", "Maquinaria / Equipos", "HR", 28, 95000, 1.00),
        ("2.2", "02. Obra Gruesa", "Hormigón H30 fundaciones y pilares", "M3", 285, "2026-05-15", "Concretero y vibrador", "Mano de Obra", "DÍA", 60, 41000, 0.96),
        ("2.3", "02. Obra Gruesa", "Enfierradura A630-420H", "KG", 24500, "2026-05-12", "Acero A630-420H cortado y doblado", "Materiales", "KG", 25200, 1190, 1.04),
        ("2.3", "02. Obra Gruesa", "Enfierradura A630-420H", "KG", 24500, "2026-05-24", "Enfierrador especialista", "Mano de Obra", "DÍA", 140, 48000, 1.07),
        ("2.4", "02. Obra Gruesa", "Moldajes de muros y losas", "M2", 1280, "2026-06-05", "Arriendo moldaje metálico", "Maquinaria / Equipos", "M2", 1280, 9800, 1.05),
        ("2.4", "02. Obra Gruesa", "Moldajes de muros y losas", "M2", 1280, "2026-06-18", "Carpintero obra gruesa", "Mano de Obra", "DÍA", 155, 49000, 1.10),
        ("2.5", "02. Obra Gruesa", "Albañilería confinada ladrillo titan", "M2", 520, "2026-06-28", "Ladrillo fiscal / titan reforzado", "Materiales", "UN", 18200, 560, 1.01),
        ("2.5", "02. Obra Gruesa", "Albañilería confinada ladrillo titan", "M2", 520, "2026-07-05", "Albañil + ayudante", "Mano de Obra", "DÍA", 95, 45000, 0.98),
        ("3.1", "03. Terminaciones", "Cubierta metálica ZincAlum + aislación", "M2", 640, "2026-07-18", "Panel ZincAlum PV6 0.5mm", "Materiales", "M2", 672, 16900, 1.00),
        ("3.1", "03. Terminaciones", "Cubierta metálica ZincAlum + aislación", "M2", 640, "2026-07-26", "Subcontrato hojalatería y cubierta", "Subcontratos", "M2", 640, 12500, 1.04),
        ("3.2", "03. Terminaciones", "Revestimiento EIFS y pintura exterior", "M2", 980, "2026-08-10", "Sistema EIFS + malla + pasta", "Materiales", "M2", 1020, 19500, 0.97),
        ("3.2", "03. Terminaciones", "Revestimiento EIFS y pintura exterior", "M2", 980, "2026-08-22", "Pintor fachada y terminaciones", "Mano de Obra", "DÍA", 88, 43000, 0.94),
        ("3.3", "03. Terminaciones", "Porcelanato y cerámica áreas húmedas", "M2", 710, "2026-09-04", "Porcelanato 60x60 alto tráfico", "Materiales", "M2", 760, 21400, 1.06),
        ("3.3", "03. Terminaciones", "Porcelanato y cerámica áreas húmedas", "M2", 710, "2026-09-14", "Ceramista especialista", "Mano de Obra", "DÍA", 75, 48000, 1.02),
        ("3.4", "03. Terminaciones", "Ventanas termopanel aluminio", "M2", 145, "2026-09-24", "Suministro e inst. termopanel", "Subcontratos", "M2", 145, 132000, 0.95),
        ("4.1", "04. Instalaciones", "Red eléctrica, tableros y luminarias LED", "GL", 1, "2026-10-08", "Subcontrato eléctrico SEC", "Subcontratos", "GL", 1, 18500000, 0.92),
        ("4.2", "04. Instalaciones", "Instalación sanitaria agua potable y alcant.", "GL", 1, "2026-10-18", "Cañerías PPR/PVC y artefactos", "Materiales", "GL", 1, 9800000, 0.96),
        ("4.2", "04. Instalaciones", "Instalación sanitaria agua potable y alcant.", "GL", 1, "2026-10-25", "Gásfiter autorizado", "Mano de Obra", "DÍA", 65, 46000, 0.90),
    ]

    registros = []
    for cod, cap, part, u_part, c_part, f_str, rec, tipo, u_rec, c_rec, pu, f_real in partidas_demo:
        total = float(c_rec) * float(pu)
        # Simular avance real hasta septiembre/principios de octubre
        fecha_dt = pd.to_datetime(f_str)
        costo_real = total * f_real if fecha_dt <= pd.Timestamp("2026-09-30") else 0.0
        registros.append(
            {
                "Código": cod,
                "Partida": part,
                "Capítulo": cap,
                "Unidad_Partida": u_part,
                "Cantidad_Partida": float(c_part),
                "Código_Recurso": cod,
                "Recurso": rec,
                "Tipo_Recurso": tipo,
                "Unidad": u_rec,
                "Cantidad": float(c_rec),
                "Precio_Unitario": float(pu),
                "Total": total,
                "Fecha": fecha_dt,
                "Costo_Real": costo_real,
            }
        )
    return pd.DataFrame(registros)


# ---------------------------------------------------------------------------
# Construcción de serie temporal para Curva S
# ---------------------------------------------------------------------------
def construir_serie_curva_s(
    df: pd.DataFrame,
    periodicidad: str,
    fecha_inicio_sim: dt.date,
    fecha_fin_sim: dt.date,
    modelo_sim: str,
) -> tuple[pd.DataFrame, bool, bool]:
    """
    Construye la serie temporal de inversión parcial y acumulada (PV y EV si existe).
    Retorna: (df_curva, es_simulada, tiene_real).
    """
    datos = df.copy()
    tiene_real = "Costo_Real" in datos.columns and datos["Costo_Real"].notna().any() and (datos["Costo_Real"].fillna(0).sum() > 0)

    # Caso 1: Columnas matriciales de periodo tipo texto ('Mes 1', 'Ene-26', etc.)
    if datos["Fecha"].notna().any() and datos["Fecha"].dtype == object:
        # Intentar parsear como fecha real; si no, preservar orden de aparición de periodo
        fechas_parseadas = pd.to_datetime(datos["Fecha"], errors="coerce")
        if fechas_parseadas.notna().sum() >= len(datos) * 0.6:
            datos["Fecha_Dt"] = fechas_parseadas
            es_simulada = False
        else:
            agg_cols = {"Total": "sum"}
            if tiene_real:
                agg_cols["Costo_Real"] = "sum"
            res = datos.groupby("Fecha", sort=False, as_index=False).agg(agg_cols)
            res = res.rename(columns={"Fecha": "Periodo", "Total": "PV_Parcial"})
            res["PV_Acum"] = res["PV_Parcial"].cumsum()
            tot_pv = res["PV_Parcial"].sum()
            res["PV_Pct"] = (res["PV_Acum"] / tot_pv * 100.0) if tot_pv > 0 else 0.0
            if tiene_real:
                res["EV_Parcial"] = res["Costo_Real"]
                res["EV_Acum"] = res["EV_Parcial"].cumsum()
                res["EV_Pct"] = (res["EV_Acum"] / tot_pv * 100.0) if tot_pv > 0 else 0.0
            return res, False, tiene_real
    elif datos["Fecha"].notna().any():
        datos["Fecha_Dt"] = pd.to_datetime(datos["Fecha"], errors="coerce")
        es_simulada = False
    else:
        es_simulada = True
        datos["Fecha_Dt"] = pd.NaT

    # Si no hay fechas válidas en el archivo, asignar fechas simuladas según modelo PMO
    if datos["Fecha_Dt"].isna().all():
        es_simulada = True
        inicio_ts = pd.Timestamp(fecha_inicio_sim)
        fin_ts = pd.Timestamp(max(fecha_fin_sim, fecha_inicio_sim + dt.timedelta(days=14)))
        total_dias = max(1, (fin_ts - inicio_ts).days)

        freq = "MS" if periodicidad == "Mensual" else "W-MON"
        periodos_idx = pd.date_range(start=inicio_ts, end=fin_ts, freq=freq)
        if len(periodos_idx) < 2:
            periodos_idx = pd.date_range(start=inicio_ts, periods=6, freq=freq)

        n_per = len(periodos_idx)
        # Pesos según modelo elegido
        t_norm = np.linspace(0.05, 0.95, n_per)
        if modelo_sim == "Curva S Logística (PMO típica)":
            # Derivada de sigmoide (campana suave)
            pesos_base = np.exp(-((t_norm - 0.5) ** 2) / (2 * 0.21**2))
            pesos_base = pesos_base / pesos_base.sum()
            pv_parcial = datos["Total"].sum() * pesos_base
            res = pd.DataFrame(
                {
                    "Periodo_Dt": periodos_idx,
                    "Periodo": [
                        p.strftime("%b %Y") if periodicidad == "Mensual" else p.strftime("%d-%b-%y")
                        for p in periodos_idx
                    ],
                    "PV_Parcial": pv_parcial,
                }
            )
            res["PV_Acum"] = res["PV_Parcial"].cumsum()
            tot = res["PV_Parcial"].sum()
            res["PV_Pct"] = (res["PV_Acum"] / tot * 100.0) if tot > 0 else 0.0
            return res, True, False
        elif modelo_sim == "Secuencial por Capítulos (WBS)":
            caps = list(dict.fromkeys(datos["Capítulo"].tolist()))
            mapa_cap_idx = {c: i for i, c in enumerate(caps)}
            n_caps = max(1, len(caps))
            fechas_asignadas = []
            for _, row in datos.iterrows():
                c_idx = mapa_cap_idx[row["Capítulo"]]
                frac = (c_idx + 0.5) / n_caps
                fecha_row = inicio_ts + pd.Timedelta(days=int(frac * total_dias))
                fechas_asignadas.append(fecha_row)
            datos["Fecha_Dt"] = fechas_asignadas
        else:
            # Uniforme en el rango de fechas
            fechas_asignadas = [
                inicio_ts + pd.Timedelta(days=int(i / max(1, len(datos) - 1) * total_dias))
                for i in range(len(datos))
            ]
            datos["Fecha_Dt"] = fechas_asignadas

    datos = datos[datos["Fecha_Dt"].notna()].copy()
    if datos.empty:
        return pd.DataFrame(), es_simulada, tiene_real

    code_freq = "M" if periodicidad == "Mensual" else "W-MON"
    datos["Periodo_Dt"] = datos["Fecha_Dt"].dt.to_period(code_freq).dt.start_time

    agg_dict: dict[str, str] = {"Total": "sum"}
    if tiene_real:
        datos["Costo_Real_Clean"] = datos["Costo_Real"].fillna(0.0)
        agg_dict["Costo_Real_Clean"] = "sum"

    res = datos.groupby("Periodo_Dt", as_index=False).agg(agg_dict).sort_values("Periodo_Dt")
    res = res.rename(columns={"Total": "PV_Parcial"})
    res["Periodo"] = res["Periodo_Dt"].dt.strftime("%b %Y" if periodicidad == "Mensual" else "%d-%b-%y")
    res["PV_Acum"] = res["PV_Parcial"].cumsum()
    tot_pv = res["PV_Parcial"].sum()
    res["PV_Pct"] = (res["PV_Acum"] / tot_pv * 100.0) if tot_pv > 0 else 0.0

    if tiene_real:
        res["EV_Parcial"] = res["Costo_Real_Clean"]
        res["EV_Acum"] = res["EV_Parcial"].cumsum()
        res["EV_Pct"] = (res["EV_Acum"] / tot_pv * 100.0) if tot_pv > 0 else 0.0

    return res.reset_index(drop=True), es_simulada, tiene_real


# ---------------------------------------------------------------------------
# Gráficos interactivos con Plotly (Identidad PRISMA)
# ---------------------------------------------------------------------------
def grafico_dona_tipo_recurso(df: pd.DataFrame) -> go.Figure:
    """Gráfico de Dona: Distribución porcentual del costo por Tipo de Recurso."""
    agrupado = df.groupby("Tipo_Recurso", as_index=False)["Total"].sum().sort_values("Total", ascending=False)
    fig = px.pie(
        agrupado,
        names="Tipo_Recurso",
        values="Total",
        hole=0.55,
        color_discrete_sequence=PRISMA_PALETTE,
    )
    fig.update_traces(
        textposition="outside",
        textinfo="percent+label",
        hovertemplate="<b>%{label}</b><br>Costo: $ %{value:,.0f}<br>Participación: %{percent}<extra></extra>",
        marker=dict(line=dict(color="#FFFFFF", width=2)),
    )
    fig.update_layout(
        title=dict(text="Distribución de Costo por Tipo de Recurso", font=dict(color=PRISMA_NAVY, size=15)),
        height=440,
        margin=dict(t=60, b=55, l=35, r=35),
        legend=dict(orientation="h", yanchor="bottom", y=-0.16, xanchor="center", x=0.5),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


def grafico_top15_recursos(df: pd.DataFrame) -> go.Figure:
    """Gráfico de Barras Horizontales: Top 15 insumos más costosos del proyecto."""
    agrupado = (
        df.groupby(["Recurso", "Tipo_Recurso"], as_index=False)["Total"]
        .sum()
        .sort_values("Total", ascending=False)
        .head(15)
        .sort_values("Total", ascending=True)
    )
    total_proyecto = df["Total"].sum()
    agrupado["Incidencia"] = (agrupado["Total"] / total_proyecto * 100.0) if total_proyecto > 0 else 0.0
    agrupado["Etiqueta"] = agrupado["Recurso"].map(lambda s: acortar_etiqueta(s, 36))

    fig = px.bar(
        agrupado,
        x="Total",
        y="Etiqueta",
        color="Tipo_Recurso",
        orientation="h",
        color_discrete_sequence=PRISMA_PALETTE,
        custom_data=["Recurso", "Tipo_Recurso", "Incidencia"],
    )
    fig.update_traces(
        hovertemplate=(
            "<b>%{customdata[0]}</b><br>"
            "Tipo: %{customdata[1]}<br>"
            "Costo: $ %{x:,.0f}<br>"
            "Incidencia: %{customdata[2]:.2f}%<extra></extra>"
        )
    )
    fig.update_layout(
        title=dict(text="Top 15 Recursos / Insumos de Mayor Costo", font=dict(color=PRISMA_NAVY, size=15)),
        height=440,
        xaxis_title="Costo Total ($)",
        yaxis_title="",
        margin=dict(t=60, b=55, l=10, r=20),
        legend=dict(title="", orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
    )
    fig.update_xaxes(gridcolor="#E2E8F0", tickformat="$,.0f")
    return fig


def grafico_pareto_recursos(df: pd.DataFrame, max_items: int = 20) -> go.Figure:
    """Diagrama de Pareto (80/20): Barras de costo por recurso + línea de porcentaje acumulado."""
    agrupado = (
        df[df["Total"] > 0]
        .groupby("Recurso", as_index=False)["Total"]
        .sum()
        .sort_values("Total", ascending=False)
    )
    total_global = agrupado["Total"].sum()
    top = agrupado.head(max_items).copy()
    top["Pct"] = (top["Total"] / total_global * 100.0) if total_global > 0 else 0.0
    top["Pct_Acum"] = top["Pct"].cumsum()
    top["Etiqueta"] = top["Recurso"].map(lambda s: acortar_etiqueta(s, 26))

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Bar(
            x=top["Etiqueta"],
            y=top["Total"],
            name="Costo Recurso ($)",
            marker_color=PRISMA_NAVY,
            customdata=np.stack([top["Recurso"], top["Pct"]], axis=-1),
            hovertemplate="<b>%{customdata[0]}</b><br>Costo: $ %{y:,.0f}<br>Incidencia: %{customdata[1]:.2f}%<extra></extra>",
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=top["Etiqueta"],
            y=top["Pct_Acum"],
            name="% Acumulado",
            mode="lines+markers",
            line=dict(color=PRISMA_GOLD, width=3),
            marker=dict(size=7, color=PRISMA_GOLD, line=dict(width=1.5, color="#FFFFFF")),
            hovertemplate="<b>%{x}</b><br>Acumulado: %{y:.1f}%<extra></extra>",
        ),
        secondary_y=True,
    )
    fig.add_hline(
        y=80,
        line_dash="dot",
        line_color="#DC2626",
        annotation_text="Umbral Pareto 80%",
        annotation_position="top left",
        secondary_y=True,
    )
    fig.update_layout(
        title=dict(
            text=f"Diagrama de Pareto 80/20 — Top {len(top)} Recursos Críticos",
            font=dict(color=PRISMA_NAVY, size=15),
        ),
        hovermode="x unified",
        margin=dict(t=55, b=85, l=20, r=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1.0),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
    )
    fig.update_xaxes(tickangle=-35, tickfont=dict(size=10))
    fig.update_yaxes(title_text="Costo ($)", tickformat="$,.0f", gridcolor="#E2E8F0", secondary_y=False)
    fig.update_yaxes(title_text="% Acumulado", range=[0, 108], ticksuffix="%", showgrid=False, secondary_y=True)
    return fig


def grafico_treemap_jerarquico(df: pd.DataFrame) -> go.Figure:
    """Treemap jerárquico: Capítulo -> Tipo de Recurso -> Recurso."""
    datos = df[df["Total"] > 0].copy()
    datos["Capítulo_C"] = datos["Capítulo"].map(lambda s: acortar_etiqueta(s, 32))
    datos["Recurso_C"] = datos["Recurso"].map(lambda s: acortar_etiqueta(s, 36))
    agrupado = (
        datos.groupby(["Capítulo_C", "Tipo_Recurso", "Recurso_C"], as_index=False)["Total"].sum()
    )

    fig = px.treemap(
        agrupado,
        path=[px.Constant("Presupuesto Directo"), "Capítulo_C", "Tipo_Recurso", "Recurso_C"],
        values="Total",
        color="Tipo_Recurso",
        color_discrete_sequence=PRISMA_PALETTE,
    )
    fig.update_traces(
        hovertemplate="<b>%{label}</b><br>Costo: $ %{value:,.0f}<br>% del Padre: %{percentParent:.1%}<br>% del Total: %{percentRoot:.1%}<extra></extra>",
        textinfo="label+percent parent",
    )
    fig.update_layout(
        title=dict(text="Mapa de Árbol (Treemap): Capítulo → Tipo de Recurso → Recurso", font=dict(color=PRISMA_NAVY, size=15)),
        margin=dict(t=50, b=15, l=10, r=10),
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig


def grafico_curva_s(df_curva: pd.DataFrame, vista_porcentual: bool, tiene_real: bool) -> go.Figure:
    """Gráfico Combinado Curva S: Barras parciales + Curva Acumulada Planificada (PV) y Real (EV)."""
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    tot_pv = df_curva["PV_Parcial"].sum()

    if vista_porcentual:
        y_bar_pv = (df_curva["PV_Parcial"] / tot_pv * 100.0) if tot_pv > 0 else 0.0
        y_line_pv = df_curva["PV_Pct"]
        titulo_y1 = "Inversión Parcial (%)"
        titulo_y2 = "Avance Acumulado (%)"
        fmt_y1 = ".1f"
        suf_y1 = "%"
        hover_bar = "<b>%{x}</b><br>Planificado Parcial: %{y:.2f}%<extra></extra>"
        hover_pv = "<b>%{x}</b><br>Curva S Planificada (PV): %{y:.2f}%<extra></extra>"
    else:
        y_bar_pv = df_curva["PV_Parcial"]
        y_line_pv = df_curva["PV_Acum"]
        titulo_y1 = "Inversión Parcial ($)"
        titulo_y2 = "Costo Acumulado ($)"
        fmt_y1 = "$,.0f"
        suf_y1 = ""
        hover_bar = "<b>%{x}</b><br>Planificado Parcial: $ %{y:,.0f}<extra></extra>"
        hover_pv = "<b>%{x}</b><br>Curva S Planificada (PV): $ %{y:,.0f}<extra></extra>"

    fig.add_trace(
        go.Bar(
            x=df_curva["Periodo"],
            y=y_bar_pv,
            name="Parcial Planificado (PV)",
            marker_color="rgba(21, 51, 82, 0.35)",
            marker_line_color=PRISMA_NAVY,
            marker_line_width=1,
            hovertemplate=hover_bar,
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=df_curva["Periodo"],
            y=y_line_pv,
            name="Curva S Planificada (PV)",
            mode="lines+markers",
            line=dict(color=PRISMA_NAVY, width=3.2),
            marker=dict(size=7, color=PRISMA_NAVY),
            hovertemplate=hover_pv,
        ),
        secondary_y=True,
    )

    if tiene_real and "EV_Acum" in df_curva.columns:
        if vista_porcentual:
            y_bar_ev = (df_curva["EV_Parcial"] / tot_pv * 100.0) if tot_pv > 0 else 0.0
            y_line_ev = df_curva["EV_Pct"]
            hover_bar_ev = "<b>%{x}</b><br>Ejecutado Parcial: %{y:.2f}%<extra></extra>"
            hover_ev = "<b>%{x}</b><br>Curva S Real (EV): %{y:.2f}%<extra></extra>"
        else:
            y_bar_ev = df_curva["EV_Parcial"]
            y_line_ev = df_curva["EV_Acum"]
            hover_bar_ev = "<b>%{x}</b><br>Ejecutado Parcial: $ %{y:,.0f}<extra></extra>"
            hover_ev = "<b>%{x}</b><br>Curva S Real (EV): $ %{y:,.0f}<extra></extra>"

        fig.add_trace(
            go.Bar(
                x=df_curva["Periodo"],
                y=y_bar_ev,
                name="Parcial Ejecutado (EV)",
                marker_color="rgba(200, 146, 54, 0.48)",
                marker_line_color=PRISMA_GOLD,
                marker_line_width=1,
                hovertemplate=hover_bar_ev,
            ),
            secondary_y=False,
        )
        fig.add_trace(
            go.Scatter(
                x=df_curva["Periodo"],
                y=y_line_ev,
                name="Curva S Real (EV)",
                mode="lines+markers",
                line=dict(color=PRISMA_GOLD, width=3.2, dash="solid"),
                marker=dict(size=7, symbol="diamond", color=PRISMA_GOLD),
                hovertemplate=hover_ev,
            ),
            secondary_y=True,
        )

    fig.update_layout(
        title=dict(
            text="Curva S · Avance Físico y Financiero",
            font=dict(color=PRISMA_NAVY, size=15),
        ),
        barmode="group",
        hovermode="x unified",
        margin=dict(t=60, b=45, l=20, r=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1.0),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
    )
    fig.update_yaxes(
        title_text=titulo_y1,
        tickformat=fmt_y1,
        ticksuffix=suf_y1,
        gridcolor="#E2E8F0",
        secondary_y=False,
    )
    fig.update_yaxes(
        title_text=titulo_y2,
        tickformat=".0f" if vista_porcentual else "$,.0f",
        ticksuffix="%" if vista_porcentual else "",
        range=[0, 108] if vista_porcentual else None,
        showgrid=False,
        secondary_y=True,
    )
    return fig


# ---------------------------------------------------------------------------
# Exportación limpia a Excel con openpyxl
# ---------------------------------------------------------------------------
def exportar_excel_estructurado(df_export: pd.DataFrame) -> bytes:
    """Exporta la tabla filtrada y un resumen por capítulo y tipo de recurso a .xlsx."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Detalle_Recursos")

        res_cap = (
            df_export.groupby("Capítulo", as_index=False)["Total"]
            .sum()
            .sort_values("Total", ascending=False)
        )
        res_cap.to_excel(writer, index=False, sheet_name="Resumen_Capitulos")

        res_tipo = (
            df_export.groupby("Tipo_Recurso", as_index=False)["Total"]
            .sum()
            .sort_values("Total", ascending=False)
        )
        res_tipo.to_excel(writer, index=False, sheet_name="Resumen_Tipos")

        # Ajuste básico de ancho de columnas
        for ws in writer.sheets.values():
            for col in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 48)

    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Aplicación Principal (Streamlit)
# ---------------------------------------------------------------------------
def main() -> None:
    aplicar_estilos_prisma()

    # -----------------------------------------------------------------------
    # Barra lateral: fuente de datos, mapeo dinámico y filtros
    # -----------------------------------------------------------------------
    with st.sidebar:
        if LOGO_PRISMA_PATH.exists():
            st.image(str(LOGO_PRISMA_PATH), use_container_width=True)
        else:
            st.markdown(f"### 🏗️ PRISMA SpA Constructora")

        st.header("1. Origen de Datos")
        modo_demo = st.checkbox(
            "Cargar datos de ejemplo de construcción",
            value=False,
            help="Genera un presupuesto de edificación de ejemplo en memoria.",
        )

        archivo_subido = None
        usar_local = False
        if not modo_demo:
            archivo_subido = st.file_uploader(
                "Sube el archivo Excel (.xlsx / .xls)",
                type=["xlsx", "xls"],
                help="Compatible con presupuestos Anexo N° 4, Presto, Arquímedes, MS Project o plantillas propias.",
            )
            if archivo_subido is None and ARCHIVO_LOCAL_DEFECTO.exists():
                usar_local = st.checkbox(
                    f"Usar archivo local ({ARCHIVO_LOCAL_DEFECTO.name})",
                    value=True,
                )

    df_datos = pd.DataFrame()
    df_pie = pd.DataFrame()
    metadatos: dict[str, str] = {}
    info_extra: dict[str, Any] = {"formulas_vacias": 0, "expandido_apu": False}
    nombre_fuente = "Proyecto Demo Edificación PRISMA"
    modo_badge = "MODO DEMO PMO"

    # Rama A: Generador de Datos Demo
    if modo_demo:
        df_datos = generar_datos_demo()
        metadatos = {
            "Proyecto": "Edificio Corporativo PRISMA (Datos Demo)",
            "Ubicación": "Concepción, Región del Biobío",
        }
    else:
        if archivo_subido is None and not usar_local:
            render_banner_corporativo(
                "Sube un archivo Excel en la barra lateral o activa el Modo Demo para iniciar.",
                "ESPERANDO DATOS",
            )
            st.info(
                "👈 Activa **'Cargar datos de ejemplo de construcción'** o sube tu archivo `.xlsx`/`.xls` en la barra lateral."
            )
            st.stop()

        try:
            if archivo_subido is not None:
                file_bytes = archivo_subido.getvalue()
                nombre_fuente = archivo_subido.name
            else:
                file_bytes = ARCHIVO_LOCAL_DEFECTO.read_bytes()
                nombre_fuente = ARCHIVO_LOCAL_DEFECTO.name
        except Exception as exc:
            st.error(f"No se pudo leer el archivo seleccionado: {exc}")
            st.stop()

        # Lectura en una sola pasada
        try:
            with st.spinner("Procesando libro Excel con openpyxl..."):
                paquete = cargar_libro_y_hoja_unica(file_bytes, nombre_fuente, hoja_objetivo=None)
        except Exception as exc:
            st.error(f"Error al abrir el libro Excel: {exc}")
            st.stop()

        sheetnames: list[str] = paquete["sheetnames"]
        with st.sidebar:
            hoja_defecto = paquete["hoja"]
            hoja_sel = st.selectbox(
                "Hoja de cálculo",
                options=sheetnames,
                index=sheetnames.index(hoja_defecto),
            )

        # Si el usuario eligió una hoja distinta a la por defecto, recargar solo esa hoja cacheada
        if hoja_sel != paquete["hoja"]:
            try:
                with st.spinner(f"Cargando hoja '{hoja_sel}'..."):
                    paquete = cargar_libro_y_hoja_unica(file_bytes, nombre_fuente, hoja_objetivo=hoja_sel)
            except Exception as exc:
                st.error(f"No se pudo leer la hoja '{hoja_sel}': {exc}")
                st.stop()

        filas_crudas = paquete["filas"]
        if not filas_crudas:
            st.warning("La hoja seleccionada no contiene datos.")
            st.stop()

        with st.sidebar:
            st.header("2. Asignación de Columnas")
            fila_auto = detectar_fila_encabezado(filas_crudas)
            auto_hdr = st.checkbox("Detectar fila de encabezados automáticamente", value=True)
            fila_hdr = (
                fila_auto
                if auto_hdr
                else int(
                    st.number_input(
                        "Fila de encabezados (1-based)",
                        min_value=1,
                        max_value=max(1, len(filas_crudas)),
                        value=fila_auto,
                    )
                )
            )
            if auto_hdr:
                st.caption(f"Encabezados detectados en **fila {fila_hdr}**.")

            fila_hdr_tuple = filas_crudas[fila_hdr - 1]
            max_cols = max(len(fila_hdr_tuple), max((len(r) for r in filas_crudas[:50]), default=0))
            etiquetas_cols, mapa_col_idx = construir_opciones_columnas(fila_hdr_tuple, max_cols)
            sugerencias = sugerir_mapeo_columnas(etiquetas_cols)

            expandir_apu = False
            if paquete["tiene_estructura_apu"]:
                expandir_apu = st.checkbox(
                    f"Extraer recursos desde hojas APU ({len(paquete['apu_sheets'])} hojas)",
                    value=True,
                    help="Desglosa Materiales, Mano de Obra, Equipos y Subcontratos desde cada hoja APU.",
                )

            opciones_map = ["(Ninguna)"] + etiquetas_cols

            def mapear_campo(label: str, rol: str) -> int | None:
                sug = sugerencias.get(rol)
                idx = opciones_map.index(sug) if sug in opciones_map else 0
                sel = st.selectbox(label, options=opciones_map, index=idx, key=f"map_{rol}_{hoja_sel}")
                return None if sel == "(Ninguna)" else mapa_col_idx[sel]

            with st.expander("Mapeador dinámico de columnas", expanded=False):
                col_indices = {
                    "codigo": mapear_campo("Código / ID Partida", "codigo"),
                    "partida": mapear_campo("Nombre Partida / Actividad", "partida"),
                    "capitulo": mapear_campo("Capítulo / Especialidad / WBS", "capitulo"),
                    "recurso": mapear_campo("Recurso / Insumo", "recurso"),
                    "tipo_recurso": mapear_campo("Tipo de Recurso", "tipo_recurso"),
                    "unidad": mapear_campo("Unidad de Medida", "unidad"),
                    "cantidad": mapear_campo("Cantidad / Metrado", "cantidad"),
                    "precio_unitario": mapear_campo("Precio Unitario (PU)", "precio_unitario"),
                    "total": mapear_campo("Costo Total / Parcial", "total"),
                    "fecha": None,
                    "costo_real": None,
                }

        usadas_set = {v for v in sugerencias.values() if v is not None}
        cols_matriciales = (
            detectar_columnas_periodos_matriciales(etiquetas_cols, usadas_set)
            if col_indices["fecha"] is None
            else []
        )

        try:
            df_datos, df_pie, info_extra = procesar_datos_excel(
                filas=filas_crudas,
                header_row_1based=fila_hdr,
                col_indices=col_indices,
                etiquetas_cols=etiquetas_cols,
                mapa_col_idx=mapa_col_idx,
                apu_sheets=paquete["apu_sheets"],
                expandir_apu=expandir_apu,
                cols_periodo_matricial=cols_matriciales,
            )
            metadatos = extraer_metadatos_proyecto(filas_crudas, fila_hdr)
        except Exception as exc:
            st.error(f"Error al estructurar el presupuesto: {exc}")
            st.stop()

        modo_badge = (
            f"ANEXO N° 4 + APU ({info_extra['partidas_apu_vinculadas']} APU)"
            if info_extra.get("expandido_apu")
            else f"HOJA: {hoja_sel}"
        )

    if df_datos.empty:
        st.warning("No se encontraron registros válidos. Ajusta la fila de encabezado o el mapeo de columnas.")
        st.stop()

    # -----------------------------------------------------------------------
    # Filtros globales en sidebar
    # -----------------------------------------------------------------------
    with st.sidebar:
        st.header("3. Filtros")
        caps_disp = list(dict.fromkeys(df_datos["Capítulo"].tolist()))
        caps_sel = st.multiselect("Capítulo / Especialidad", options=caps_disp, default=caps_disp)

        tipos_disp = sorted(df_datos["Tipo_Recurso"].unique().tolist())
        tipos_sel = st.multiselect("Tipo de Recurso", options=tipos_disp, default=tipos_disp)

        unidades_disp = sorted(df_datos["Unidad"].unique().tolist(), key=lambda x: (0 if x == "M2" else 1, x))
        unidades_sel = st.multiselect("Unidad de Medida", options=unidades_disp, default=unidades_disp)

        buscador = st.text_input("Buscar partida o recurso", placeholder="Ej. Hormigón, Carpintero, M2…").strip()

    # Aplicar filtros globales
    mask = (
        df_datos["Capítulo"].isin(caps_sel)
        & df_datos["Tipo_Recurso"].isin(tipos_sel)
        & df_datos["Unidad"].isin(unidades_sel)
    )
    if buscador:
        pat = re.escape(normalizar_texto(buscador))
        mask = mask & (
            df_datos["Partida"].map(normalizar_texto).str.contains(pat, regex=True)
            | df_datos["Recurso"].map(normalizar_texto).str.contains(pat, regex=True)
            | df_datos["Código"].map(normalizar_texto).str.contains(pat, regex=True)
        )

    df_filt = df_datos[mask].copy()

    # Banner corporativo
    meta_str = f"<b>Archivo:</b> {nombre_fuente}"
    for k, v in metadatos.items():
        meta_str += f" &nbsp;|&nbsp; <b>{k}:</b> {v}"
    render_banner_corporativo(meta_str, modo_badge)

    if info_extra.get("formulas_vacias", 0) > 0:
        st.warning(
            f"Se detectaron {info_extra['formulas_vacias']} filas con fórmulas sin valor cacheado en `data_only=True`. "
            "Guarda el Excel desde Microsoft Excel antes de subirlo para ver todos los montos calculados."
        )

    if df_filt.empty:
        st.warning("Los filtros actuales no devuelven registros. Amplía la selección en la barra lateral.")
        st.stop()

    # -----------------------------------------------------------------------
    # A. Panel Superior de KPIs
    # -----------------------------------------------------------------------
    total_presupuesto = df_filt["Total"].sum()
    partidas_activas = df_filt[["Código", "Partida"]].drop_duplicates().shape[0]
    recursos_unicos = df_filt["Recurso"].nunique()

    rank_recursos = df_filt.groupby("Recurso", as_index=False)["Total"].sum().sort_values("Total", ascending=False)
    if not rank_recursos.empty and total_presupuesto > 0:
        top_rec_nombre = str(rank_recursos.iloc[0]["Recurso"])
        top_rec_pct = rank_recursos.iloc[0]["Total"] / total_presupuesto * 100.0
    else:
        top_rec_nombre = "N/A"
        top_rec_pct = 0.0

    por_tipo = df_filt.groupby("Tipo_Recurso")["Total"].sum()
    pct_mat = (por_tipo.get("Materiales", 0.0) / total_presupuesto * 100.0) if total_presupuesto > 0 else 0.0
    pct_mo = (por_tipo.get("Mano de Obra", 0.0) / total_presupuesto * 100.0) if total_presupuesto > 0 else 0.0

    k1, k2, k3, k4, k5 = st.columns([1.45, 0.82, 0.82, 1.0, 1.1])
    with k1:
        st.metric("Presupuesto", formato_moneda(total_presupuesto, 0))
    with k2:
        st.metric("Partidas", f"{partidas_activas:,}".replace(",", "."))
    with k3:
        st.metric("Recursos", f"{recursos_unicos:,}".replace(",", "."))
    with k4:
        st.metric(
            "Recurso Líder",
            f"{top_rec_pct:.1f}%",
            delta=acortar_etiqueta(top_rec_nombre, 26),
            delta_color="off",
        )
    with k5:
        st.metric(
            "Mats. / Mano de Obra",
            f"{pct_mat:.0f}% / {pct_mo:.0f}%",
            delta="Materiales / MO",
            delta_color="off",
        )

    if not df_pie.empty:
        with st.expander("💼 Resumen Económico del Presupuesto", expanded=False):
            cols_pie = st.columns(min(len(df_pie), 6))
            for i, (_, r_pie) in enumerate(df_pie.iterrows()):
                cols_pie[i % len(cols_pie)].metric(
                    label=str(r_pie["Concepto"]),
                    value=formato_moneda(r_pie["Monto"], 0),
                    delta=f"Factor: {r_pie['Factor'] * 100:.2f}%" if pd.notna(r_pie["Factor"]) and r_pie["Factor"] < 1 else None,
                )

    # -----------------------------------------------------------------------
    # Pestañas Principales del Dashboard
    # -----------------------------------------------------------------------
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        [
            "📋 Presupuesto y trazabilidad",
            "🧱 Analizador de APU",
            "📦 Recursos y abastecimiento",
            "📊 Críticos y Pareto",
            "💱 Escenarios de precios",
            "🛡️ Calidad del presupuesto",
        ]
    )

    # 1. Presupuesto y trazabilidad
    with tab1:
        insumo_rastrear = st.text_input(
            "Rastreador de partida o recurso",
            value="",
            key="rastreador_insumo",
            placeholder="Ej. Hormigón, Carpintero, OSB",
        ).strip()

        df_tab1 = df_filt.copy()
        if insumo_rastrear:
            pat_ins = re.escape(normalizar_texto(insumo_rastrear))
            df_tab1 = df_tab1[
                df_tab1["Recurso"].map(normalizar_texto).str.contains(pat_ins, regex=True)
                | df_tab1["Partida"].map(normalizar_texto).str.contains(pat_ins, regex=True)
            ]
            if df_tab1.empty:
                st.warning(f"No se encontraron partidas o recursos relacionados con '{insumo_rastrear}'.")

        if not df_tab1.empty:
            total_t1 = df_tab1["Total"].sum()
            df_vista = df_tab1[
                ["Capítulo", "Código", "Partida", "Tipo_Recurso", "Recurso", "Unidad", "Cantidad", "Precio_Unitario", "Total"]
            ].copy()
            df_vista["Incidencia (%)"] = df_vista["Total"] / total_t1 * 100.0 if total_t1 > 0 else 0.0

            st.dataframe(
                df_vista,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Capítulo": st.column_config.TextColumn("Capítulo / WBS", width="medium"),
                    "Código": st.column_config.TextColumn("Código", width="small"),
                    "Partida": st.column_config.TextColumn("Partida / Actividad", width="large"),
                    "Tipo_Recurso": st.column_config.TextColumn("Tipo", width="small"),
                    "Recurso": st.column_config.TextColumn("Recurso / Insumo", width="large"),
                    "Unidad": st.column_config.TextColumn("Unidad", width="small"),
                    "Cantidad": st.column_config.NumberColumn("Cantidad", format="%.2f"),
                    "Precio_Unitario": st.column_config.NumberColumn("Precio Unitario", format="$ %.0f"),
                    "Total": st.column_config.NumberColumn("Costo Total", format="$ %.0f"),
                    "Incidencia (%)": st.column_config.ProgressColumn("Incidencia", format="%.2f%%", min_value=0.0, max_value=max(float(df_vista["Incidencia (%)"].max()), 1.0)),
                },
            )

            with st.expander("Resumen por capítulo y partida", expanded=False):
                res_jer = (
                    df_tab1.groupby(["Capítulo", "Código", "Partida"], as_index=False)
                    .agg(Recursos=("Recurso", "nunique"), Costo_Total=("Total", "sum"))
                    .sort_values(["Capítulo", "Código"])
                )
                res_jer["% Incidencia"] = res_jer["Costo_Total"] / total_t1 * 100.0 if total_t1 > 0 else 0.0
                st.dataframe(res_jer, use_container_width=True, hide_index=True)

            excel_bytes = exportar_excel_estructurado(df_vista)
            st.download_button(
                "Exportar detalle a Excel",
                data=excel_bytes,
                file_name=f"{Path(nombre_fuente).stem}_PRISMA_APU.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # 2. Analizador de APU
    with tab2:
        partidas_disp = (
            df_filt[["Código", "Partida"]]
            .drop_duplicates()
            .sort_values(["Código", "Partida"])
        )
        partidas_disp["_Etiqueta"] = partidas_disp["Código"].astype(str) + " — " + partidas_disp["Partida"].astype(str)
        partida_sel = st.selectbox("Seleccione una partida", options=partidas_disp["_Etiqueta"].tolist())
        codigo_sel = partida_sel.split(" — ", 1)[0]
        df_apu = df_filt[df_filt["Código"].astype(str) == codigo_sel].copy()

        costo_apu = df_apu["Total"].sum()
        recursos_apu = df_apu["Recurso"].nunique()
        por_tipo_apu = df_apu.groupby("Tipo_Recurso", as_index=False)["Total"].sum().sort_values("Total", ascending=False)
        por_tipo_apu["Incidencia"] = por_tipo_apu["Total"] / costo_apu * 100.0 if costo_apu > 0 else 0.0

        c1, c2, c3 = st.columns(3)
        c1.metric("Costo directo del APU", formato_moneda(costo_apu, 0))
        c2.metric("Recursos incluidos", f"{recursos_apu:,}".replace(",", "."))
        c3.metric("Tipos de recurso", str(df_apu["Tipo_Recurso"].nunique()))

        cc1, cc2 = st.columns([1.35, 1])
        with cc1:
            st.dataframe(
                df_apu[
                    ["Tipo_Recurso", "Recurso", "Unidad", "Cantidad", "Precio_Unitario", "Total"]
                ].sort_values("Total", ascending=False),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Precio_Unitario": st.column_config.NumberColumn("Precio Unitario", format="$ %.0f"),
                    "Total": st.column_config.NumberColumn("Costo Total", format="$ %.0f"),
                },
            )
        with cc2:
            fig_apu = px.pie(por_tipo_apu, names="Tipo_Recurso", values="Total", hole=0.55, color_discrete_sequence=PRISMA_PALETTE)
            fig_apu.update_layout(title="Composición del APU", height=400, margin=dict(t=50, b=20))
            st.plotly_chart(fig_apu, use_container_width=True)

    # 3. Recursos y abastecimiento
    with tab3:
        res_abast = (
            df_filt.groupby(["Recurso", "Tipo_Recurso", "Unidad"], as_index=False)
            .agg(Cantidad_Total=("Cantidad", "sum"), Costo_Total=("Total", "sum"), Partidas=("Código", "nunique"))
            .sort_values("Costo_Total", ascending=False)
        )
        res_abast["Precio_Ponderado"] = res_abast["Costo_Total"] / res_abast["Cantidad_Total"].where(res_abast["Cantidad_Total"] != 0)
        res_abast["Incidencia"] = res_abast["Costo_Total"] / total_presupuesto * 100.0 if total_presupuesto > 0 else 0.0
        st.dataframe(
            res_abast,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Cantidad_Total": st.column_config.NumberColumn("Cantidad Total", format="%.2f"),
                "Costo_Total": st.column_config.NumberColumn("Costo Total", format="$ %.0f"),
                "Precio_Ponderado": st.column_config.NumberColumn("Costo / Unidad", format="$ %.2f"),
                "Incidencia": st.column_config.ProgressColumn("Incidencia", format="%.2f%%", min_value=0.0, max_value=max(float(res_abast["Incidencia"].max()), 1.0)),
            },
        )
        st.download_button(
            "Exportar listado de abastecimiento",
            data=res_abast.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"{Path(nombre_fuente).stem}_abastecimiento.csv",
            mime="text/csv",
        )

    # 4. Críticos y Pareto
    with tab4:
        c1, c2 = st.columns([1, 1.35])
        with c1:
            st.plotly_chart(grafico_dona_tipo_recurso(df_filt), use_container_width=True)
        with c2:
            st.plotly_chart(grafico_top15_recursos(df_filt), use_container_width=True)
        st.plotly_chart(grafico_pareto_recursos(df_filt, max_items=20), use_container_width=True)
        st.plotly_chart(grafico_treemap_jerarquico(df_filt), use_container_width=True)

    # 5. Escenarios económicos
    with tab5:
        st.caption("Calcule el impacto de variaciones de mercado sobre el presupuesto filtrado.")
        tipos_escenario = [
            "Materiales",
            "Mano de Obra",
            "Maquinaria / Equipos",
            "Subcontratos",
            "Otros / Partida Directa",
        ]
        variaciones: dict[str, float] = {}
        cols_esc = st.columns(len(tipos_escenario))
        for i, tipo in enumerate(tipos_escenario):
            if tipo in df_filt["Tipo_Recurso"].unique():
                variaciones[tipo] = cols_esc[i].slider(f"{tipo}", -30, 50, 0, 1, format="%d%%")

        df_esc = df_filt.copy()
        df_esc["Variación"] = df_esc["Tipo_Recurso"].map(lambda x: variaciones.get(str(x), 0.0) / 100.0)
        df_esc["Costo_Escenario"] = df_esc["Total"] * (1.0 + df_esc["Variación"])
        costo_escenario = df_esc["Costo_Escenario"].sum()

        m1, m2, m3 = st.columns(3)
        m1.metric("Presupuesto base", formato_moneda(total_presupuesto, 0))
        m2.metric("Presupuesto escenario", formato_moneda(costo_escenario, 0))
        m3.metric("Variación", f"{((costo_escenario / total_presupuesto - 1) * 100 if total_presupuesto else 0):.2f}%")

        resumen_esc = df_esc.groupby("Tipo_Recurso", as_index=False).agg(Base=("Total", "sum"), Escenario=("Costo_Escenario", "sum"))
        resumen_esc["Variación $"] = resumen_esc["Escenario"] - resumen_esc["Base"]
        fig_esc = px.bar(resumen_esc, x="Tipo_Recurso", y=["Base", "Escenario"], barmode="group", labels={"value": "Costo", "Tipo_Recurso": "Tipo de recurso"}, color_discrete_sequence=[PRISMA_NAVY, PRISMA_GOLD])
        fig_esc.update_layout(title="Comparación por tipo de recurso", margin=dict(t=50, b=50))
        st.plotly_chart(fig_esc, use_container_width=True)
        st.dataframe(resumen_esc, use_container_width=True, hide_index=True)

    # 6. Diagnóstico de calidad
    with tab6:
        sin_codigo = df_filt[df_filt["Código"].isna() | df_filt["Código"].astype(str).str.strip().isin(["", "None"])]
        sin_recurso = df_filt[df_filt["Recurso"].isna() | df_filt["Recurso"].astype(str).str.strip().isin(["", "None"])]
        costos_invalidos = df_filt[(df_filt["Total"].isna()) | (df_filt["Total"] <= 0)]
        cantidades_invalidas = df_filt[(df_filt["Cantidad"].isna()) | (df_filt["Cantidad"] <= 0)]
        duplicados = df_filt.duplicated(subset=["Código", "Partida", "Recurso", "Unidad", "Cantidad", "Precio_Unitario"], keep=False)

        resumen_diag = [
            ("Filas sin código", len(sin_codigo)),
            ("Filas sin recurso", len(sin_recurso)),
            ("Costos cero o negativos", len(costos_invalidos)),
            ("Cantidades cero o negativas", len(cantidades_invalidas)),
            ("Filas duplicadas", int(duplicados.sum())),
        ]
        diag_df = pd.DataFrame(resumen_diag, columns=["Control", "Cantidad"])
        st.dataframe(diag_df, use_container_width=True, hide_index=True)

        if sin_recurso.empty and costos_invalidos.empty and cantidades_invalidas.empty and not duplicados.any():
            st.success("No se detectaron inconsistencias básicas en los registros filtrados.")

        with st.expander("Partidas sin recursos detectables", expanded=False):
            agrupadas = df_filt.groupby(["Código", "Partida"], as_index=False)["Recurso"].nunique()
            st.dataframe(agrupadas[agrupadas["Recurso"] == 0], use_container_width=True, hide_index=True)

        with st.expander("Revisar registros con costo no válido", expanded=False):
            st.dataframe(costos_invalidos, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
