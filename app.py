"""
Aplicación Streamlit para Análisis Inteligente de Presupuestos de Obra (Anexo N° 4).

Características principales:
- Lectura de libros .xlsx con openpyxl (data_only=True) para obtener valores evaluados de fórmulas.
- Detección automática o manual de encabezados y mapeo flexible de columnas.
- Limpieza jerárquica (Capítulos, Subcapítulos/Partidas principales y Subpartidas/Recursos),
  eliminación de subtotales duplicados y separación de filas resumen de pie de presupuesto.
- Conversión segura de formatos numéricos y de moneda, y detección normalizada de m².
- Panel de métricas clave, 3 gráficos profesionales con Matplotlib y descarga en Excel / CSV.
"""

from __future__ import annotations

import datetime as dt
import io
import math
import re
import unicodedata
from pathlib import Path
from typing import Any

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import openpyxl
import pandas as pd
import streamlit as st


# ---------------------------------------------------------------------------
# Configuración de página y estilos
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Análisis de Presupuesto de Obra",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARCHIVO_LOCAL_DEFECTO = Path(__file__).resolve().parent / "7.21.5_Anexo_N°_4_Presupuesto.xlsx"

# Palabras clave normalizadas para detección automática de encabezados
HEADER_KEYWORDS: dict[str, list[str]] = {
    "item": [
        "item",
        "codigo",
        "cod",
        "nro",
        "numero",
        "n°",
        "no.",
    ],
    "descripcion": [
        "designacion",
        "descripcion",
        "descripcion recurso",
        "actividad",
        "partida",
        "especificacion",
        "detalle",
        "recurso",
        "glosa",
        "concepto",
    ],
    "unidad": [
        "unidad",
        "un.",
        "un",
        "u",
        "und",
        "um",
        "u.m.",
    ],
    "cantidad": [
        "cantidad",
        "cant",
        "cant.",
        "cubicacion",
        "metrado",
        "rendimiento",
    ],
    "precio_unitario": [
        "precio unitario",
        "costo unitario directo ($)",
        "costo unitario directo",
        "costo unitario",
        "valor unitario",
        "p.u.",
        "p.u",
        "pu",
        "valor",
        "precio",
        "sueldo",
    ],
    "total": [
        "costo total",
        "valor total",
        "total ($)",
        "total",
        "monto",
        "importe",
        "subtotal",
    ],
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

M2_TOKENS = {
    "M2",
    "M²",
    "MT2",
    "MTS2",
    "METRO CUADRADO",
    "METROS CUADRADOS",
}


# ---------------------------------------------------------------------------
# Funciones auxiliares de normalización y parsing seguro
# ---------------------------------------------------------------------------
def normalizar_texto(valor: Any) -> str:
    """Convierte texto a minúsculas sin tildes ni espacios redundantes."""
    if valor is None:
        return ""
    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"\s+", " ", texto)
    return texto


def normalizar_unidad_estandar(valor: Any) -> str:
    """Normaliza una unidad de medida para filtrado uniforme (ej. m2, m², M2 -> M2)."""
    if valor is None:
        return ""
    raw = str(valor).strip()
    if not raw or raw.lower() in {"none", "nan", "-"}:
        return ""
    norm = normalizar_texto(raw).upper().replace("²", "2").replace(".", "").strip()
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
        return norm  # Mantiene U/UN/UNI legible o agrupable en mayúsculas
    return norm


def es_metro_cuadrado(valor: Any) -> bool:
    """Determina si la unidad corresponde a metros cuadrados (m2, m², M2, etc.)."""
    if valor is None:
        return False
    norm = normalizar_texto(valor).upper().replace("²", "2").replace(".", "")
    compact = re.sub(r"\s+", "", norm)
    if compact in {"M2", "MT2", "MTS2", "METROCUADRADO", "METROSCUADRADOS"}:
        return True
    return norm.strip() in M2_TOKENS


def parsear_numero(valor: Any) -> float | None:
    """
    Convierte de forma segura celdas numéricas o cadenas con formato de moneda a float.
    Soporta formatos chilenos/europeos ($ 4.100.000 / 1.234,56) y US (1,234.56).
    """
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        if math.isnan(valor) or math.isinf(valor):
            return None
        return float(valor)

    texto = str(valor).strip()
    if not texto:
        return None

    # Valores simbólicos que representan cero o vacío
    limpio_basico = texto.replace("$", "").replace("CLP", "").replace("USD", "").strip()
    if limpio_basico in {"-", "—", "–", "N/A", "NA", "S/I", "None", "nan"}:
        return None

    tiene_moneda = "$" in texto or "CLP" in texto.upper() or "USD" in texto.upper()
    negativo = False
    if limpio_basico.startswith("(") and limpio_basico.endswith(")"):
        negativo = True
        limpio_basico = limpio_basico[1:-1].strip()

    # Conservar dígitos, signos, puntos y comas
    candidato = re.sub(r"[^\d.,\-]", "", limpio_basico)
    if not candidato or candidato in {"-", ".", ","}:
        return None

    if candidato.startswith("-"):
        negativo = True
        candidato = candidato[1:]
    candidato = candidato.replace("-", "")
    if not candidato:
        return None

    # Resolución de separadores decimales y de miles
    if "." in candidato and "," in candidato:
        # El último separador define el decimal
        if candidato.rfind(",") > candidato.rfind("."):
            candidato = candidato.replace(".", "").replace(",", ".")
        else:
            candidato = candidato.replace(",", "")
    elif "," in candidato:
        # Múltiples comas -> separador de miles; una coma -> decimal
        if candidato.count(",") > 1:
            candidato = candidato.replace(",", "")
        else:
            candidato = candidato.replace(",", ".")
    elif "." in candidato:
        partes = candidato.split(".")
        if len(partes) > 2:
            # Ej. 4.100.000 -> miles
            candidato = "".join(partes)
        elif len(partes) == 2 and len(partes[1]) == 3 and tiene_moneda:
            # Ej. $ 400.000 -> miles en moneda
            candidato = "".join(partes)
        # En otro caso (ej. "23.446"), se mantiene como decimal

    try:
        numero = float(candidato)
        return -numero if negativo else numero
    except ValueError:
        return None


def normalizar_item(valor: Any, capitulo_actual: str = "") -> str:
    """
    Normaliza el código de ítem:
    - Enteros almacenados como float (1.0 -> '1').
    - Códigos convertidos accidentalmente en fecha por Excel (ej. '10.1' -> 2026-01-10 -> '10.1').
    """
    if valor is None:
        return ""

    if isinstance(valor, (dt.datetime, dt.date)):
        # En configuración regional DD-MM, "5.4" se convierte en día=5, mes=4
        cap_num = capitulo_actual.split(" - ")[0].strip() if capitulo_actual else ""
        dia_mes = f"{valor.day}.{valor.month}"
        mes_dia = f"{valor.month}.{valor.day}"
        if cap_num and str(valor.month) == cap_num and str(valor.day) != cap_num:
            return mes_dia
        return dia_mes

    if isinstance(valor, int):
        return str(valor)

    if isinstance(valor, float):
        if math.isnan(valor) or math.isinf(valor):
            return ""
        if valor.is_integer():
            return str(int(valor))
        return f"{valor:g}"

    texto = str(valor).strip()
    if re.fullmatch(r"\d+\.0+", texto):
        return texto.split(".")[0]
    return texto


def formato_moneda(valor: float, decimales: int = 0) -> str:
    """Formatea un valor numérico como moneda legible con separador de miles '.'."""
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return "$ 0"
    if decimales == 0:
        texto = f"{valor:,.0f}"
    else:
        texto = f"{valor:,.{decimales}f}"
    # Estilo chileno/latino: punto para miles, coma para decimales
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"$ {texto}"


def formato_numero(valor: float, decimales: int = 2) -> str:
    """Formatea cantidades numéricas con separador de miles '.' y decimal ','."""
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return "0"
    texto = f"{valor:,.{decimales}f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def acortar_etiqueta(texto: str, max_len: int = 44) -> str:
    """Acorta etiquetas largas para gráficos sin perder legibilidad."""
    texto = re.sub(r"\s+", " ", str(texto)).strip()
    if len(texto) <= max_len:
        return texto
    return texto[: max_len - 1].rstrip() + "…"


# ---------------------------------------------------------------------------
# Lectura con openpyxl y detección inteligente de encabezados
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def obtener_nombres_hojas(file_bytes: bytes) -> list[str]:
    """Lee los nombres de hojas del archivo .xlsx usando openpyxl con data_only=True."""
    wb = None
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        return list(wb.sheetnames)
    finally:
        if wb is not None:
            wb.close()


@st.cache_data(show_spinner=False)
def leer_hoja_cruda(file_bytes: bytes, sheet_name: str) -> list[tuple[Any, ...]]:
    """Extrae todas las filas evaluadas (data_only=True) de la hoja seleccionada."""
    wb = None
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"La hoja '{sheet_name}' no existe en el libro.")
        ws = wb[sheet_name]
        filas: list[tuple[Any, ...]] = []
        for fila in ws.iter_rows(values_only=True):
            filas.append(tuple(fila))
        return filas
    finally:
        if wb is not None:
            wb.close()


def seleccionar_hoja_por_defecto(sheetnames: list[str]) -> int:
    """Elige automáticamente la hoja principal de presupuesto si existe."""
    prioridad = ["PPTO MODIFICADO", "PRESUPUESTO", "PPTO", "RESUMEN", "ANEXO"]
    norm_sheets = [normalizar_texto(s).upper() for s in sheetnames]
    for clave in prioridad:
        for idx, s in enumerate(norm_sheets):
            if clave in s:
                return idx
    return 0


def detectar_fila_encabezado(filas: list[tuple[Any, ...]], max_scan: int = 60) -> int:
    """
    Detecta la fila (1-based) donde comienzan los encabezados de la tabla
    buscando palabras clave típicas de presupuestos de construcción.
    """
    mejor_fila = 1
    mejor_puntaje = -1.0

    limite = min(len(filas), max_scan)
    for idx in range(limite):
        celdas = [normalizar_texto(c) for c in filas[idx]]
        no_vacias = [c for c in celdas if c]
        if len(no_vacias) < 3:
            continue

        roles_encontrados: set[str] = set()
        for celda in no_vacias:
            for rol, keywords in HEADER_KEYWORDS.items():
                if any(celda == kw or (len(kw) >= 4 and kw in celda) for kw in keywords):
                    roles_encontrados.add(rol)

        # Ponderación adicional si aparecen descripción + total o cantidad
        puntaje = float(len(roles_encontrados))
        if "descripcion" in roles_encontrados:
            puntaje += 0.6
        if "total" in roles_encontrados:
            puntaje += 0.5
        if "item" in roles_encontrados:
            puntaje += 0.4
        if "unidad" in roles_encontrados:
            puntaje += 0.3

        if puntaje > mejor_puntaje and len(roles_encontrados) >= 3:
            mejor_puntaje = puntaje
            mejor_fila = idx + 1

    return mejor_fila


def extraer_metadatos_proyecto(
    filas: list[tuple[Any, ...]], header_row_1based: int
) -> dict[str, str]:
    """Extrae metadatos institucionales (Proyecto, Ubicación, Contratista, etc.) sobre el encabezado."""
    metadatos: dict[str, str] = {}
    claves_interes = {
        "proyecto": "Proyecto",
        "ubicacion": "Ubicación",
        "contratista": "Contratista",
        "obra": "Obra",
        "mandante": "Mandante",
        "fecha": "Fecha",
    }

    for fila in filas[: max(0, header_row_1based - 1)]:
        valores = [v for v in fila if v is not None and str(v).strip() != ""]
        if len(valores) < 2:
            continue
        etiqueta_norm = normalizar_texto(valores[0]).rstrip(":").strip()
        for clave, nombre_mostrar in claves_interes.items():
            if etiqueta_norm == clave or etiqueta_norm.startswith(clave):
                valor = valores[1]
                if isinstance(valor, (dt.datetime, dt.date)):
                    metadatos[nombre_mostrar] = valor.strftime("%d-%m-%Y")
                else:
                    metadatos[nombre_mostrar] = str(valor).strip()
                break
    return metadatos


def construir_opciones_columnas(
    fila_encabezado: tuple[Any, ...], max_cols: int
) -> tuple[list[str], dict[str, int]]:
    """Construye etiquetas únicas para cada columna del encabezado y su índice 0-based."""
    etiquetas: list[str] = []
    mapa_indices: dict[str, int] = {}
    conteo: dict[str, int] = {}

    for col_idx in range(max_cols):
        val = fila_encabezado[col_idx] if col_idx < len(fila_encabezado) else None
        nombre = str(val).strip() if val is not None and str(val).strip() != "" else ""
        if not nombre:
            letra_col = openpyxl.utils.get_column_letter(col_idx + 1)
            nombre = f"Columna {letra_col}"

        if nombre in conteo:
            conteo[nombre] += 1
            nombre_unico = f"{nombre} ({conteo[nombre]})"
        else:
            conteo[nombre] = 1
            nombre_unico = nombre

        etiquetas.append(nombre_unico)
        mapa_indices[nombre_unico] = col_idx

    return etiquetas, mapa_indices


def sugerir_mapeo_columnas(etiquetas: list[str]) -> dict[str, str | None]:
    """Sugiere automáticamente la mejor columna para cada rol presupuestario."""
    sugerencias: dict[str, str | None] = {
        "item": None,
        "descripcion": None,
        "unidad": None,
        "cantidad": None,
        "precio_unitario": None,
        "total": None,
    }
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
            for prioridad_idx, kw in enumerate(keywords):
                score = 0
                if norm == kw:
                    score = 100 - prioridad_idx
                elif len(kw) >= 4 and kw in norm:
                    score = 60 - prioridad_idx
                if score > mejor_score:
                    mejor_score = score
                    mejor_col = etiqueta
        if mejor_col is not None:
            sugerencias[rol] = mejor_col
            usadas.add(mejor_col)

    return sugerencias


# ---------------------------------------------------------------------------
# Limpieza inteligente y clasificación jerárquica
# ---------------------------------------------------------------------------
def es_fila_resumen_pie(
    item_raw: Any, desc_raw: Any, cant_raw: Any, unit_raw: Any
) -> tuple[bool, str]:
    """Identifica filas resumen al pie del presupuesto (DIRECTO, GG, UTIL, NETO, IVA, TOTAL)."""
    for candidato in (desc_raw, cant_raw, item_raw):
        if candidato is None:
            continue
        norm = normalizar_texto(candidato).upper().rstrip(":").strip()
        if norm in FOOTER_SUMMARY_TOKENS:
            return True, norm
        if any(norm.startswith(pref) for pref in SUBTOTAL_PREFIXES):
            return True, norm
        if norm.startswith("$ SUBTOTAL"):
            return True, norm

    # Si no hay código ni descripción pero sí una etiqueta de texto en cantidad o unidad
    if item_raw is None and desc_raw is None and isinstance(cant_raw, str):
        norm_cant = normalizar_texto(cant_raw).upper().strip()
        if norm_cant and parsear_numero(norm_cant) is None:
            return True, norm_cant
    if item_raw is None and desc_raw is None and unit_raw is None and cant_raw is not None:
        if isinstance(cant_raw, str):
            return True, cant_raw.strip()

    return False, ""


def clasificar_profundidad_item(item_str: str) -> int:
    """Devuelve la profundidad jerárquica de un código tipo '1', '1.2', '2.8.1.1'."""
    if not item_str:
        return 0
    limpio = item_str.strip().rstrip(".")
    if re.fullmatch(r"\d+(\.\d+)*", limpio):
        return len(limpio.split("."))
    return 0


def procesar_presupuesto(
    filas: list[tuple[Any, ...]],
    header_row_1based: int,
    col_indices: dict[str, int | None],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Procesa y limpia la tabla presupuestaria desde la fila posterior al encabezado.

    Retorna:
    - df_partidas: Ítems ejecutables/valorizados (sin capítulos agrupadores ni subtotales duplicados).
    - df_jerarquia: Tabla completa incluyendo capítulos y subcapítulos detectados.
    - df_resumen_pie: Filas resumen detectadas al pie (DIRECTO, GG, UTIL, NETO, IVA, TOTAL).
    """
    idx_item = col_indices.get("item")
    idx_desc = col_indices.get("descripcion")
    idx_unit = col_indices.get("unidad")
    idx_cant = col_indices.get("cantidad")
    idx_pu = col_indices.get("precio_unitario")
    idx_tot = col_indices.get("total")

    def obtener_celda(fila: tuple[Any, ...], idx: int | None) -> Any:
        if idx is None or idx < 0 or idx >= len(fila):
            return None
        return fila[idx]

    registros_raw: list[dict[str, Any]] = []
    resumen_pie: list[dict[str, Any]] = []

    capitulo_actual = "Sin Capítulo"
    subcapitulo_actual = "General"

    for num_fila, fila in enumerate(filas[header_row_1based:], start=header_row_1based + 1):
        # Ignorar filas completamente vacías en las columnas mapeadas
        vals_mapeados = [
            obtener_celda(fila, i)
            for i in (idx_item, idx_desc, idx_unit, idx_cant, idx_pu, idx_tot)
            if i is not None
        ]
        if not any(v is not None and str(v).strip() != "" for v in vals_mapeados):
            continue

        raw_item = obtener_celda(fila, idx_item)
        raw_desc = obtener_celda(fila, idx_desc)
        raw_unit = obtener_celda(fila, idx_unit)
        raw_cant = obtener_celda(fila, idx_cant)
        raw_pu = obtener_celda(fila, idx_pu)
        raw_tot = obtener_celda(fila, idx_tot)

        # Detectar filas resumen o subtotales al pie
        es_pie, etiqueta_pie = es_fila_resumen_pie(raw_item, raw_desc, raw_cant, raw_unit)
        if es_pie:
            tot_pie = parsear_numero(raw_tot)
            pu_pie = parsear_numero(raw_pu)
            if tot_pie is not None or pu_pie is not None:
                resumen_pie.append(
                    {
                        "Concepto": etiqueta_pie,
                        "Factor / Tasa": pu_pie,
                        "Monto ($)": tot_pie if tot_pie is not None else pu_pie,
                        "Fila Excel": num_fila,
                    }
                )
            continue

        item_str = normalizar_item(raw_item, capitulo_actual)
        desc_str = re.sub(r"\s+", " ", str(raw_desc)).strip() if raw_desc is not None else ""
        unit_str = str(raw_unit).strip() if raw_unit is not None else ""
        unit_estandar = normalizar_unidad_estandar(unit_str)

        cant_val = parsear_numero(raw_cant)
        pu_val = parsear_numero(raw_pu)
        tot_val = parsear_numero(raw_tot)

        # Descartar filas plantilla vacías de APU (ej. "MAT-", "MAQ-" sin descripción y total 0)
        if not desc_str and (tot_val is None or tot_val == 0) and (cant_val is None or cant_val == 0):
            continue

        # Descartar notas al pie puramente textuales sin código ni valores numéricos
        if not item_str and not unit_str and cant_val is None and pu_val is None and tot_val is None:
            continue

        # Si falta el total pero existen cantidad y precio unitario, calcularlo de forma segura
        if tot_val is None and cant_val is not None and pu_val is not None:
            tot_val = cant_val * pu_val

        profundidad = clasificar_profundidad_item(item_str)
        tiene_valores = any(v is not None for v in (cant_val, pu_val, tot_val)) or bool(unit_str)

        # Diferenciar Capítulos / Partidas principales agrupadoras vs Partidas ejecutables / Recursos
        if not tiene_valores and desc_str:
            if profundidad == 1 or (item_str.upper().startswith(("A-", "B-", "C-", "D-")) and not unit_str):
                capitulo_actual = f"{item_str} - {desc_str}" if item_str else desc_str
                subcapitulo_actual = desc_str
                tipo_fila = "Capítulo"
            else:
                if profundidad <= 2:
                    subcapitulo_actual = f"{item_str} - {desc_str}" if item_str else desc_str
                tipo_fila = "Subcapítulo / Partida Principal"
        else:
            # Si viene un ítem de nivel 1 con valores propios, actualizar también el capítulo
            if profundidad == 1 and desc_str:
                capitulo_actual = f"{item_str} - {desc_str}"
                subcapitulo_actual = desc_str
            elif profundidad == 2 and desc_str:
                subcapitulo_actual = f"{item_str} - {desc_str}"
            tipo_fila = "Partida / Recurso"

        registros_raw.append(
            {
                "Fila Excel": num_fila,
                "Ítem": item_str,
                "Descripción": desc_str,
                "Unidad": unit_str,
                "Unidad_Estandar": unit_estandar,
                "Cantidad": cant_val if cant_val is not None else 0.0,
                "Precio Unitario": pu_val if pu_val is not None else 0.0,
                "Total": tot_val if tot_val is not None else 0.0,
                "Tiene_Valor": tiene_valores,
                "Tiene_Cantidad_o_PU": (cant_val is not None) or (pu_val is not None) or bool(unit_str),
                "Profundidad": profundidad,
                "Capítulo": capitulo_actual,
                "Partida Principal": subcapitulo_actual,
                "Tipo": tipo_fila,
                "Es_m2": es_metro_cuadrado(unit_str),
            }
        )

    df_jerarquia = pd.DataFrame(registros_raw)
    df_resumen_pie = pd.DataFrame(resumen_pie)

    if df_jerarquia.empty:
        return df_jerarquia.copy(), df_jerarquia, df_resumen_pie

    # Detección adicional de subtotales duplicados en filas padre:
    # Si una fila tiene código padre (ej. '1' o '1.1'), sin unidad/cantidad/PU, pero sí un Total,
    # y existen filas hijas con prefijo '1.' o '1.1.', se marca como Subtotal Padre para no duplicar costos.
    codigos_existentes = [c for c in df_jerarquia["Ítem"].tolist() if c]
    es_subtotal_padre: list[bool] = []
    for _, row in df_jerarquia.iterrows():
        codigo = row["Ítem"]
        if (
            codigo
            and not row["Tiene_Cantidad_o_PU"]
            and row["Total"] > 0
            and any(c != codigo and c.startswith(f"{codigo}.") for c in codigos_existentes)
        ):
            es_subtotal_padre.append(True)
        else:
            es_subtotal_padre.append(False)

    df_jerarquia["Es_Subtotal_Padre"] = es_subtotal_padre
    df_jerarquia.loc[df_jerarquia["Es_Subtotal_Padre"], "Tipo"] = "Subtotal Capítulo/Partida"

    df_partidas = df_jerarquia[
        (df_jerarquia["Tipo"] == "Partida / Recurso") & (~df_jerarquia["Es_Subtotal_Padre"])
    ].copy()

    # Etiqueta única para gráficos ([Ítem] Descripción) evitando colisiones por nombres repetidos
    df_partidas["Etiqueta"] = df_partidas.apply(
        lambda r: f"{r['Ítem']} · {r['Descripción']}" if r["Ítem"] else r["Descripción"],
        axis=1,
    )

    return df_partidas.reset_index(drop=True), df_jerarquia.reset_index(drop=True), df_resumen_pie


# ---------------------------------------------------------------------------
# Gráficos profesionales con Matplotlib
# ---------------------------------------------------------------------------
def configurar_estilo_matplotlib() -> None:
    """Aplica estilo limpio y legible para todos los gráficos Matplotlib."""
    plt.rcParams.update(
        {
            "font.sans-serif": ["DejaVu Sans", "Segoe UI", "Arial"],
            "axes.edgecolor": "#CBD5E1",
            "axes.linewidth": 0.8,
            "axes.titlesize": 12,
            "axes.titleweight": "bold",
            "axes.labelsize": 10,
            "xtick.labelsize": 9,
            "ytick.labelsize": 9,
        }
    )


def grafico_distribucion_capitulos(df: pd.DataFrame, agrupar_por: str = "Capítulo") -> plt.Figure:
    """
    Gráfico 1: Barras horizontales ordenadas de mayor a menor con monto total ($)
    y porcentaje de incidencia por Capítulo o Partida Principal.
    """
    configurar_estilo_matplotlib()
    agrupado = (
        df.groupby(agrupar_por, as_index=False)["Total"]
        .sum()
        .sort_values("Total", ascending=False)
    )
    total_general = agrupado["Total"].sum()
    agrupado["Pct"] = (agrupado["Total"] / total_general * 100.0) if total_general > 0 else 0.0

    # Si hay muchas subcategorías, limitar visualmente a Top 20 para legibilidad
    if len(agrupado) > 20:
        agrupado = agrupado.head(20)

    etiquetas = [acortar_etiqueta(x, 42) for x in agrupado[agrupar_por]]
    montos = agrupado["Total"].values
    pcts = agrupado["Pct"].values

    alto = max(4.2, 0.42 * len(agrupado) + 1.4)
    fig, ax = plt.subplots(figsize=(10.5, alto), dpi=130)

    colores = plt.cm.Blues([0.45 + 0.45 * (1 - i / max(1, len(agrupado))) for i in range(len(agrupado))])
    y_pos = range(len(agrupado))
    barras = ax.barh(y_pos, montos, color=colores, edgecolor="#1E3A8A", linewidth=0.5, height=0.68)

    ax.set_yticks(list(y_pos))
    ax.set_yticklabels(etiquetas)
    ax.invert_yaxis()  # Mayor incidencia arriba

    max_monto = max(montos) if len(montos) > 0 else 1.0
    ax.set_xlim(0, max_monto * 1.28)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: formato_moneda(x, 0)))
    ax.grid(axis="x", linestyle="--", alpha=0.45, color="#94A3B8")
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    ax.set_title(f"Distribución del Presupuesto por {agrupar_por}", pad=12, loc="left", color="#0F172A")
    ax.set_xlabel("Costo Total ($)")

    for barra, monto, pct in zip(barras, montos, pcts):
        ax.text(
            barra.get_width() + max_monto * 0.015,
            barra.get_y() + barra.get_height() / 2,
            f"{formato_moneda(monto, 0)}  ({pct:.1f}%)",
            va="center",
            ha="left",
            fontsize=8.5,
            color="#1E293B",
            fontweight="medium",
        )

    fig.tight_layout()
    return fig


def grafico_metros_cuadrados(df_m2: pd.DataFrame, max_barras: int = 18) -> plt.Figure:
    """
    Gráfico 2: Barras verticales dedicadas exclusivamente a las partidas medidas en m²,
    mostrando la cantidad de m² con etiquetas de valor sobre cada barra.
    """
    configurar_estilo_matplotlib()
    datos = df_m2[df_m2["Cantidad"] > 0].sort_values("Cantidad", ascending=False).head(max_barras)

    etiquetas = [acortar_etiqueta(x, 28) for x in datos["Etiqueta"]]
    cantidades = datos["Cantidad"].values

    ancho = max(10.0, 0.58 * len(datos) + 2.0)
    fig, ax = plt.subplots(figsize=(min(ancho, 15.0), 5.6), dpi=130)

    x_pos = range(len(datos))
    barras = ax.bar(x_pos, cantidades, color="#0D9488", edgecolor="#115E59", linewidth=0.6, width=0.65)

    ax.set_xticks(list(x_pos))
    ax.set_xticklabels(etiquetas, rotation=38, ha="right", fontsize=8.5)
    max_q = max(cantidades) if len(cantidades) > 0 else 1.0
    ax.set_ylim(0, max_q * 1.18)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda y, _: f"{formato_numero(y, 0)} m²"))
    ax.grid(axis="y", linestyle="--", alpha=0.45, color="#94A3B8")
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    ax.set_title(
        f"Metros Cuadrados (m²) por Partida de Obra (Top {len(datos)} actividades)",
        pad=12,
        loc="left",
        color="#0F172A",
    )
    ax.set_ylabel("Superficie (m²)")

    for barra, cant in zip(barras, cantidades):
        dec = 0 if cant >= 100 else 1
        ax.text(
            barra.get_x() + barra.get_width() / 2,
            barra.get_height() + max_q * 0.018,
            f"{formato_numero(cant, dec)} m²",
            ha="center",
            va="bottom",
            fontsize=8,
            rotation=0 if len(datos) <= 12 else 30,
            color="#0F172A",
            fontweight="bold",
        )

    fig.tight_layout()
    return fig


def grafico_pareto_top10(df: pd.DataFrame) -> plt.Figure:
    """
    Gráfico 3: Top 10 recursos / partidas de mayor incidencia en costo y curva
    de porcentaje acumulado sobre el presupuesto filtrado.
    """
    configurar_estilo_matplotlib()
    datos_validos = df[df["Total"] > 0].sort_values("Total", ascending=False)
    total_presupuesto = datos_validos["Total"].sum()
    top10 = datos_validos.head(10).copy()

    top10["Pct_Individual"] = (
        (top10["Total"] / total_presupuesto * 100.0) if total_presupuesto > 0 else 0.0
    )
    top10["Pct_Acumulado"] = top10["Pct_Individual"].cumsum()

    etiquetas = [acortar_etiqueta(x, 30) for x in top10["Etiqueta"]]
    montos = top10["Total"].values
    acumulado = top10["Pct_Acumulado"].values

    fig, ax1 = plt.subplots(figsize=(11.5, 5.6), dpi=130)
    x_pos = list(range(len(top10)))

    barras = ax1.bar(
        x_pos,
        montos,
        color="#2563EB",
        alpha=0.88,
        edgecolor="#1E3A8A",
        linewidth=0.6,
        width=0.62,
        label="Costo Partida ($)",
    )
    ax1.set_xticks(x_pos)
    ax1.set_xticklabels(etiquetas, rotation=32, ha="right", fontsize=8.5)
    max_m = max(montos) if len(montos) > 0 else 1.0
    ax1.set_ylim(0, max_m * 1.2)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda y, _: formato_moneda(y, 0)))
    ax1.set_ylabel("Monto ($)", color="#1E3A8A", fontweight="bold")
    ax1.grid(axis="y", linestyle="--", alpha=0.4, color="#94A3B8")
    ax1.set_axisbelow(True)
    ax1.spines["top"].set_visible(False)

    for barra, monto in zip(barras, montos):
        ax1.text(
            barra.get_x() + barra.get_width() / 2,
            barra.get_height() + max_m * 0.015,
            formato_moneda(monto, 0),
            ha="center",
            va="bottom",
            fontsize=7.8,
            color="#1E293B",
        )

    # Eje secundario para curva acumulada (%)
    ax2 = ax1.twinx()
    ax2.plot(
        x_pos,
        acumulado,
        color="#DC2626",
        marker="o",
        linewidth=2.2,
        markersize=6,
        label="% Acumulado del Presupuesto",
    )
    ax2.axhline(80, color="#F59E0B", linestyle=":", linewidth=1.3, label="Referencia 80%")
    ax2.set_ylim(0, 108)
    ax2.set_ylabel("Incidencia Acumulada (%)", color="#B91C1C", fontweight="bold")
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda y, _: f"{y:.0f}%"))
    ax2.spines["top"].set_visible(False)

    for x, y_pct in zip(x_pos, acumulado):
        ax2.text(
            x,
            min(y_pct + 3.0, 103),
            f"{y_pct:.1f}%",
            ha="center",
            va="bottom",
            fontsize=8,
            color="#991B1B",
            fontweight="bold",
        )

    ax1.set_title(
        "Top 10 Partidas / Recursos de Mayor Incidencia (Análisis de Pareto)",
        pad=14,
        loc="left",
        color="#0F172A",
    )

    # Leyenda combinada de ambos ejes
    h1, l1 = ax1.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax1.legend(h1 + h2, l1 + l2, loc="upper left", frameon=True, facecolor="white", framealpha=0.9, fontsize=8.5)

    fig.tight_layout()
    return fig


# ---------------------------------------------------------------------------
# Exportación limpia a Excel y CSV
# ---------------------------------------------------------------------------
def exportar_excel_bytes(df: pd.DataFrame) -> bytes:
    """Genera un archivo Excel (.xlsx) limpio en memoria para descarga."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Presupuesto_Procesado")
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Interfaz Principal (Streamlit)
# ---------------------------------------------------------------------------
def main() -> None:
    st.title("🏗️ Analizador de Presupuestos de Obra (Anexo N° 4)")
    st.caption(
        "Carga un presupuesto en Excel (.xlsx), extrae partidas, costos y cubicaciones en m² "
        "mediante `openpyxl (data_only=True)` y visualiza su incidencia con `matplotlib`."
    )

    # -----------------------------------------------------------------------
    # Barra Lateral: Carga de archivo y configuración de extracción
    # -----------------------------------------------------------------------
    with st.sidebar:
        st.header("1. Origen de Datos")
        archivo_subido = st.file_uploader(
            "Sube el archivo Excel de presupuesto (.xlsx)",
            type=["xlsx"],
            help="Se lee con openpyxl(..., data_only=True) para extraer fórmulas ya evaluadas.",
        )

        usar_ejemplo_local = False
        if archivo_subido is None and ARCHIVO_LOCAL_DEFECTO.exists():
            usar_ejemplo_local = st.checkbox(
                f"Cargar archivo de la carpeta ({ARCHIVO_LOCAL_DEFECTO.name})",
                value=True,
            )

        if archivo_subido is None and not usar_ejemplo_local:
            st.info("Sube un archivo `.xlsx` para iniciar el análisis.")
            st.stop()

        try:
            if archivo_subido is not None:
                file_bytes = archivo_subido.getvalue()
                nombre_archivo = archivo_subido.name
            else:
                file_bytes = ARCHIVO_LOCAL_DEFECTO.read_bytes()
                nombre_archivo = ARCHIVO_LOCAL_DEFECTO.name
        except Exception as exc:
            st.error(f"No se pudo leer el archivo seleccionado: {exc}")
            st.stop()

        # Lectura de hojas del libro
        try:
            with st.spinner("Leyendo estructura del libro Excel..."):
                sheetnames = obtener_nombres_hojas(file_bytes)
        except Exception as exc:
            st.error(f"Error al abrir el archivo Excel con openpyxl: {exc}")
            st.stop()

        if not sheetnames:
            st.error("El archivo Excel no contiene hojas visibles.")
            st.stop()

        idx_defecto = seleccionar_hoja_por_defecto(sheetnames)
        hoja_seleccionada = st.selectbox(
            "Hoja de cálculo a analizar",
            options=sheetnames,
            index=idx_defecto,
        )

        # Lectura de filas de la hoja
        try:
            with st.spinner(f"Extrayendo datos evaluados de '{hoja_seleccionada}'..."):
                filas_crudas = leer_hoja_cruda(file_bytes, hoja_seleccionada)
        except Exception as exc:
            st.error(f"No se pudo leer la hoja '{hoja_seleccionada}': {exc}")
            st.stop()

        if not filas_crudas:
            st.warning("La hoja seleccionada está vacía.")
            st.stop()

        # Detección automática o manual de fila de encabezados
        st.header("2. Encabezados y Columnas")
        fila_auto = detectar_fila_encabezado(filas_crudas)
        deteccion_auto = st.checkbox("Detectar fila de encabezados automáticamente", value=True)
        if deteccion_auto:
            fila_encabezado = fila_auto
            st.caption(f"Encabezado detectado en la **fila {fila_encabezado}**.")
        else:
            fila_encabezado = int(
                st.number_input(
                    "Número de fila de encabezado (1-based)",
                    min_value=1,
                    max_value=max(1, len(filas_crudas)),
                    value=fila_auto,
                    step=1,
                )
            )

        if fila_encabezado > len(filas_crudas):
            st.error("La fila de encabezado indicada excede la cantidad de filas de la hoja.")
            st.stop()

        fila_hdr_tuple = filas_crudas[fila_encabezado - 1]
        max_cols = max(len(fila_hdr_tuple), max((len(r) for r in filas_crudas[:50]), default=0))
        etiquetas_cols, mapa_col_idx = construir_opciones_columnas(fila_hdr_tuple, max_cols)
        sugerencias = sugerir_mapeo_columnas(etiquetas_cols)

        opciones_con_ninguna = ["(Ninguna)"] + etiquetas_cols

        def select_col(etiqueta_ui: str, rol: str) -> int | None:
            sugerida = sugerencias.get(rol)
            idx_sug = opciones_con_ninguna.index(sugerida) if sugerida in opciones_con_ninguna else 0
            elegida = st.selectbox(etiqueta_ui, options=opciones_con_ninguna, index=idx_sug, key=f"map_{rol}_{hoja_seleccionada}")
            return None if elegida == "(Ninguna)" else mapa_col_idx[elegida]

        with st.expander("Mapeo flexible de columnas", expanded=False):
            col_indices = {
                "item": select_col("Columna de Ítem / Código", "item"),
                "descripcion": select_col("Columna de Descripción / Recurso", "descripcion"),
                "unidad": select_col("Columna de Unidad", "unidad"),
                "cantidad": select_col("Columna de Cantidad", "cantidad"),
                "precio_unitario": select_col("Columna de Precio Unitario", "precio_unitario"),
                "total": select_col("Columna de Costo Total", "total"),
            }

    # -----------------------------------------------------------------------
    # Procesamiento y limpieza del presupuesto
    # -----------------------------------------------------------------------
    try:
        df_partidas, df_jerarquia, df_resumen_pie = procesar_presupuesto(
            filas_crudas, fila_encabezado, col_indices
        )
        metadatos = extraer_metadatos_proyecto(filas_crudas, fila_encabezado)
    except Exception as exc:
        st.error(f"Ocurrió un error al procesar la estructura del presupuesto: {exc}")
        st.stop()

    if df_partidas.empty:
        st.warning(
            "No se detectaron partidas valorizadas con el mapeo actual. "
            "Ajusta la fila de encabezado o el mapeo de columnas en la barra lateral."
        )
        st.stop()

    # -----------------------------------------------------------------------
    # Filtros en barra lateral (Capítulo y Unidad de Medida)
    # -----------------------------------------------------------------------
    with st.sidebar:
        st.header("3. Filtros de Análisis")
        capitulos_disponibles = list(dict.fromkeys(df_partidas["Capítulo"].tolist()))
        capitulos_sel = st.multiselect(
            "Filtrar por Capítulo / Partida",
            options=capitulos_disponibles,
            default=capitulos_disponibles,
        )

        unidades_disponibles = sorted(
            {u for u in df_partidas["Unidad_Estandar"].tolist() if u},
            key=lambda x: (0 if x == "M2" else 1, x),
        )
        incluir_sin_unidad = "" in df_partidas["Unidad_Estandar"].values
        opciones_unidad = unidades_disponibles + (["(Sin unidad)"] if incluir_sin_unidad else [])

        unidades_sel = st.multiselect(
            "Filtrar por Unidad de Medida",
            options=opciones_unidad,
            default=opciones_unidad,
            help="Las variantes de metros cuadrados (m2, m², M2) se agrupan bajo 'M2'.",
        )

        busqueda_texto = st.text_input(
            "Buscar por código o descripción",
            value="",
            placeholder="Ej. hormigón, pintura, moldaje…",
        ).strip()

    # Aplicar filtros
    unidad_filtrada = df_partidas["Unidad_Estandar"].replace("", "(Sin unidad)")
    mask = df_partidas["Capítulo"].isin(capitulos_sel) & unidad_filtrada.isin(unidades_sel)

    if busqueda_texto:
        patron = re.escape(normalizar_texto(busqueda_texto))
        mask = mask & (
            df_partidas["Descripción"].map(normalizar_texto).str.contains(patron, regex=True)
            | df_partidas["Ítem"].map(normalizar_texto).str.contains(patron, regex=True)
        )

    df_filtrado = df_partidas[mask].copy()

    # -----------------------------------------------------------------------
    # Banner informativo del proyecto
    # -----------------------------------------------------------------------
    info_parts = [f"📄 **Archivo:** `{nombre_archivo}`", f"📑 **Hoja:** `{hoja_seleccionada}`"]
    for clave, valor in metadatos.items():
        info_parts.append(f"**{clave}:** {valor}")
    st.info("  |  ".join(info_parts))

    if df_filtrado.empty:
        st.warning("Los filtros actuales no devuelven partidas. Amplía la selección en la barra lateral.")
        st.stop()

    # -----------------------------------------------------------------------
    # Métricas Clave (KPIs)
    # -----------------------------------------------------------------------
    presupuesto_total_base = df_partidas["Total"].sum()
    presupuesto_filtrado = df_filtrado["Total"].sum()
    df_m2 = df_filtrado[df_filtrado["Es_m2"]].copy()
    total_m2 = df_m2["Cantidad"].sum()
    costo_prom_m2 = (presupuesto_filtrado / total_m2) if total_m2 > 0 else None
    cantidad_partidas = len(df_filtrado)

    pct_del_total = (
        (presupuesto_filtrado / presupuesto_total_base * 100.0) if presupuesto_total_base > 0 else 0.0
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric(
            label="Presupuesto Total (Filtrado)",
            value=formato_moneda(presupuesto_filtrado, 0),
            delta=f"{pct_del_total:.1f}% del costo directo" if len(df_filtrado) < len(df_partidas) else "100% costo directo",
        )
    with col2:
        st.metric(
            label="Total Metros Cuadrados (m²)",
            value=f"{formato_numero(total_m2, 2)} m²",
            delta=f"{len(df_m2)} partidas en m²",
        )
    with col3:
        st.metric(
            label="Costo Promedio por m²",
            value=f"{formato_moneda(costo_prom_m2, 0)} / m²" if costo_prom_m2 is not None else "N/A",
            help="Relación entre el presupuesto filtrado y el total de m² detectados.",
        )
    with col4:
        st.metric(
            label="Partidas / Recursos",
            value=f"{cantidad_partidas}",
            delta=f"de {len(df_partidas)} ítems totales",
        )

    # Mostrar resumen oficial del pie de presupuesto si existe (Directo, GG, Utilidad, Neto, IVA, Total)
    if not df_resumen_pie.empty:
        with st.expander("💼 Resumen General del Pie de Presupuesto detectado en la hoja", expanded=False):
            cols_pie = st.columns(min(len(df_resumen_pie), 6))
            for idx, (_, fila_pie) in enumerate(df_resumen_pie.iterrows()):
                col_destino = cols_pie[idx % len(cols_pie)]
                monto_pie = fila_pie["Monto ($)"]
                tasa = fila_pie["Factor / Tasa"]
                subtexto = f"Factor: {tasa * 100:.2f}%" if pd.notna(tasa) and tasa < 1 else None
                col_destino.metric(
                    label=str(fila_pie["Concepto"]),
                    value=formato_moneda(monto_pie, 0) if pd.notna(monto_pie) else "—",
                    delta=subtexto,
                )

    st.divider()

    # -----------------------------------------------------------------------
    # Visualizaciones con Matplotlib
    # -----------------------------------------------------------------------
    st.subheader("📊 Análisis Gráfico del Presupuesto")

    # Gráfico 1: Distribución por Capítulo / Partida Principal
    col_g1_ctrl, _ = st.columns([2, 3])
    with col_g1_ctrl:
        nivel_agrupacion = st.radio(
            "Nivel de agrupación para Gráfico 1:",
            options=["Capítulo", "Partida Principal"],
            horizontal=True,
        )

    try:
        fig1 = grafico_distribucion_capitulos(df_filtrado, agrupar_por=nivel_agrupacion)
        st.pyplot(fig1, use_container_width=True)
        plt.close(fig1)
    except Exception as exc:
        st.error(f"No se pudo generar el Gráfico 1: {exc}")

    # Gráfico 2: Metros cuadrados (m²) por partida de obra
    st.markdown("#### 📐 Metros Cuadrados ($m^2$) por Partida de Obra")
    if df_m2.empty or total_m2 <= 0:
        st.info("No se encontraron partidas con unidad en metros cuadrados ($m^2$) dentro de los filtros actuales.")
    else:
        total_partidas_m2 = len(df_m2[df_m2["Cantidad"] > 0])
        if total_partidas_m2 > 10:
            max_m2_mostrar = st.slider(
                "Cantidad de partidas en m² a mostrar (ordenadas de mayor a menor):",
                min_value=5,
                max_value=total_partidas_m2,
                value=min(18, total_partidas_m2),
                step=1,
            )
        else:
            max_m2_mostrar = total_partidas_m2

        try:
            fig2 = grafico_metros_cuadrados(df_m2, max_barras=max_m2_mostrar)
            st.pyplot(fig2, use_container_width=True)
            plt.close(fig2)
        except Exception as exc:
            st.error(f"No se pudo generar el Gráfico 2: {exc}")

    # Gráfico 3: Pareto Top 10 Partidas / Recursos
    st.markdown("#### 📈 Top 10 Partidas / Recursos de Mayor Incidencia (Pareto)")
    if df_filtrado["Total"].sum() <= 0:
        st.info("No hay montos positivos para construir el gráfico de Pareto.")
    else:
        try:
            fig3 = grafico_pareto_top10(df_filtrado)
            st.pyplot(fig3, use_container_width=True)
            plt.close(fig3)
        except Exception as exc:
            st.error(f"No se pudo generar el Gráfico 3: {exc}")

    st.divider()

    # -----------------------------------------------------------------------
    # Tabla interactiva y exportación
    # -----------------------------------------------------------------------
    st.subheader("📋 Detalle Limpio de Partidas y Recursos")

    df_tabla = df_filtrado[
        [
            "Capítulo",
            "Partida Principal",
            "Ítem",
            "Descripción",
            "Unidad",
            "Cantidad",
            "Precio Unitario",
            "Total",
        ]
    ].copy()
    df_tabla["% Incidencia"] = (
        (df_tabla["Total"] / presupuesto_filtrado * 100.0) if presupuesto_filtrado > 0 else 0.0
    )

    st.dataframe(
        df_tabla,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Capítulo": st.column_config.TextColumn("Capítulo", width="medium"),
            "Partida Principal": st.column_config.TextColumn("Subcapítulo / Partida", width="medium"),
            "Ítem": st.column_config.TextColumn("Ítem", width="small"),
            "Descripción": st.column_config.TextColumn("Descripción / Recurso", width="large"),
            "Unidad": st.column_config.TextColumn("Unidad", width="small"),
            "Cantidad": st.column_config.NumberColumn("Cantidad", format="%.2f"),
            "Precio Unitario": st.column_config.NumberColumn("Precio Unitario ($)", format="$ %.0f"),
            "Total": st.column_config.NumberColumn("Costo Total ($)", format="$ %.0f"),
            "% Incidencia": st.column_config.NumberColumn("Incidencia (%)", format="%.2f %%"),
        },
    )

    # Botones de descarga (CSV y Excel)
    col_dl1, col_dl2, _ = st.columns([1.2, 1.2, 3])
    nombre_base = Path(nombre_archivo).stem
    with col_dl1:
        csv_bytes = df_tabla.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="⬇️ Descargar CSV limpio",
            data=csv_bytes,
            file_name=f"{nombre_base}_procesado.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with col_dl2:
        try:
            xlsx_bytes = exportar_excel_bytes(df_tabla)
            st.download_button(
                label="⬇️ Descargar Excel (.xlsx)",
                data=xlsx_bytes,
                file_name=f"{nombre_base}_procesado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as exc:
            st.warning(f"No se pudo preparar la descarga Excel: {exc}")

    # Vista opcional de la jerarquía completa (incluyendo encabezados y subtotales descartados)
    with st.expander("🔍 Ver estructura jerárquica completa extraída (Capítulos, Subcapítulos y Partidas)", expanded=False):
        st.dataframe(
            df_jerarquia[["Fila Excel", "Tipo", "Ítem", "Descripción", "Unidad", "Cantidad", "Precio Unitario", "Total"]],
            use_container_width=True,
            hide_index=True,
        )


if __name__ == "__main__":
    main()
