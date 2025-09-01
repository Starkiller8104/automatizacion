
# -*- coding: utf-8 -*-

import io
import os
import json
import pytz
import math
import time
import requests
import datetime as dt
from typing import List, Tuple, Optional

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

import streamlit as st

# ===============================
# Configuración de página
# ===============================
st.set_page_config(page_title=\"Indicadores IMEMSA - app.py\", layout=\"wide\")
st.title(\"📊 Indicadores IMEMSA – Generador y Corrector de Excel\")
st.caption(\"Incluye FIX (USD/MXN), y corrige Compra/Venta con spread alrededor del FIX.\")

TZ_MX = pytz.timezone(\"America/Mexico_City\")

# ===============================
# Utilidades generales
# ===============================
def now_mx() -> dt.datetime:
    return dt.datetime.now(TZ_MX)

def to_float_safe(x) -> Optional[float]:
    try:
        if x is None or (isinstance(x, float) and math.isnan(x)):
            return None
        return float(str(x).replace(\",\", \"\").strip())
    except Exception:
        return None

def business_days_mx(end_date: dt.date, n_days: int) -> List[dt.date]:
    \"\"\"
    Regresa los últimos n_days días hábiles en México ANTES de end_date (excluyendo fines de semana).
    Incluye end_date si es hábil.
    \"\"\"
    out = []
    d = end_date
    while len(out) < n_days:
        if d.weekday() < 5:  # 0=lunes .. 4=viernes
            out.append(d)
        d = d - dt.timedelta(days=1)
    return list(reversed(out))

# ===============================
# Banxico SIE API
# ===============================
def sie_fetch_series_range(series_id: str, start: dt.date, end: dt.date, banxico_token: str) -> pd.DataFrame:
    \"\"\"
    Obtiene una serie por rango de fechas (YYYY-MM-DD) de Banxico SIE.
    Devuelve DataFrame con columnas: fecha (date), valor (float) para la serie pedida.
    \"\"\"
    url = f\"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series_id}/datos/{start:%Y-%m-%d}/{end:%Y-%m-%d}\"
    headers = {\"Bmx-Token\": banxico_token.strip()}
    r = requests.get(url, headers=headers, timeout=20)
    r.raise_for_status()
    js = r.json()
    # Navegar estructura SIE
    data = js.get(\"bmx\", {}).get(\"series\", [])
    if not data:
        return pd.DataFrame(columns=[\"fecha\", \"valor\"])
    datos = data[0].get(\"datos\", [])
    rows = []
    for d in datos:
        fecha_str = d.get(\"fecha\")
        valor_str = d.get(\"dato\")
        try:
            fecha = dt.datetime.strptime(fecha_str, \"%d/%m/%Y\").date() if \"/\" in fecha_str else dt.datetime.strptime(fecha_str, \"%Y-%m-%d\").date()
        except Exception:
            # fallback por si viene otro formato
            try:
                fecha = dt.datetime.fromisoformat(fecha_str).date()
            except Exception:
                continue
        val = to_float_safe(valor_str)
        rows.append({\"fecha\": fecha, \"valor\": val})
    df = pd.DataFrame(rows).dropna().sort_values(\"fecha\")
    return df

# ===============================
# Cálculo Compra/Venta desde FIX
# ===============================
def calcular_compra_venta_desde_fix(fix_vals: List[float], spread_pct: float = 0.004) -> Tuple[List[Optional[float]], List[Optional[float]]]:
    \"\"\"
    A partir de una lista FIX (USD/MXN), calcula:
      compra = FIX * (1 - spread_pct/2)
      venta  = FIX * (1 + spread_pct/2)

    spread_pct: 0.004 => total 0.40% (compra -0.20%, venta +0.20%).
    \"\"\"
    if not fix_vals:
        return [], []
    half = spread_pct / 2.0
    compra, venta = [], []
    for v in fix_vals:
        vf = to_float_safe(v)
        if vf is None:
            compra.append(None)
            venta.append(None)
        else:
            compra.append(round(vf * (1.0 - half), 5))
            venta.append(round(vf * (1.0 + half), 5))
    return compra, venta

# ===============================
# Excel helpers
# ===============================
THIN = Side(border_style=\"thin\", color=\"808080\")
BORDER_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def ensure_hoja_indicadores(wb: Workbook) -> None:
    if \"Indicadores\" not in wb.sheetnames:
        wb.create_sheet(\"Indicadores\")
    ws = wb[\"Indicadores\"]
    # Encabezados B2..G2: Día 1..Día 5, Día actual
    headers = [\"Día 1\", \"Día 2\", \"Día 3\", \"Día 4\", \"Día 5\", \"Día actual\"]
    ws.cell(row=2, column=1, value=\"Fecha\")  # Col A para etiquetas opcionales
    for i, h in enumerate(headers, start=2):
        c = ws.cell(row=2, column=i, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal=\"center\")
        c.border = BORDER_THIN
        ws.column_dimensions[get_column_letter(i)].width = 16

    # Etiquetas base
    ws.cell(row=7, column=1, value=\"Dólar / Pesos (FIX)\").font = Font(bold=True)
    ws.cell(row=9, column=1, value=\"Dólar Americano (Compra)\").font = Font(bold=True)
    ws.cell(row=10, column=1, value=\"Dólar Americano (Venta)\").font = Font(bold=True)

    # Ancho col A
    ws.column_dimensions[\"A\"].width = 32

def escribir_fechas_encabezado(ws, fechas: List[dt.date]) -> None:
    # Escribe fechas forma YYYY-MM-DD en fila 3 B..G
    for idx, d in enumerate(fechas, start=2):
        ws.cell(row=3, column=idx, value=d.strftime(\"%Y-%m-%d\")).alignment = Alignment(horizontal=\"center\")

def escribir_valores_en_fila(ws, fila: int, valores: List[Optional[float]], col_start: int = 2) -> None:
    for i, v in enumerate(valores):
        ws.cell(row=fila, column=col_start + i, value=v)

def leer_fix_de_excel(ws, row_fix: int = 7, col_start: int = 2, col_end: int = 7) -> List[Optional[float]]:
    vals = []
    for c in range(col_start, col_end + 1):
        vals.append(to_float_safe(ws.cell(row=row_fix, column=c).value))
    return vals

def build_and_download_wb(fechas: List[dt.date], fix_vals: List[Optional[float]], spread_pct_total: float, nombre_archivo: str):
    wb = Workbook()
    # Eliminar la hoja 'Sheet' por defecto si existe
    if \"Sheet\" in wb.sheetnames:
        std = wb[\"Sheet\"]
        wb.remove(std)
    ensure_hoja_indicadores(wb)
    ws = wb[\"Indicadores\"]

    escribir_fechas_encabezado(ws, fechas)

    # FIX en fila 7
    escribir_valores_en_fila(ws, 7, fix_vals, col_start=2)

    # Compra/Venta desde FIX
    compra, venta = calcular_compra_venta_desde_fix(fix_vals, spread_pct=spread_pct_total / 100.0)
    escribir_valores_en_fila(ws, 9, compra, col_start=2)
    escribir_valores_en_fila(ws, 10, venta, col_start=2)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    st.download_button(\"⬇️ Descargar Excel\", data=bio, file_name=nombre_archivo, mime=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\")

# ===============================
# Sidebar: parámetros
# ===============================
with st.sidebar:
    st.header(\"⚙️ Parámetros\")
    modo = st.radio(\"Modo de uso\", [\"Generar Excel desde cero\", \"Corregir un Excel existente\"], index=0)
    spread_total = st.slider(\"Spread total (Compra/Venta) %\", min_value=0.10, max_value=1.50, value=0.40, step=0.05,
                             help=\"Ejemplo: 0.40% => Compra = FIX - 0.20% y Venta = FIX + 0.20%\" )

    st.markdown(\"---\")
    st.subheader(\"Banxico (opcional)\")
    banxico_token = st.text_input(\"Bmx-Token\", type=\"password\", help=\"Ingresa tu token de Banxico si deseas traer FIX automáticamente\")
    usar_banxico = st.checkbox(\"Obtener FIX desde Banxico\", value=False)

# ===============================
# Flujo 1: Generar Excel desde cero
# ===============================
if modo == \"Generar Excel desde cero\":
    st.subheader(\"🧰 Generar Excel\")
    hoy = now_mx().date()
    fechas = business_days_mx(hoy, 6)  # 5 días + hoy (si es hábil). Se mostrará como Día 1..Día 5..Día actual

    if usar_banxico and banxico_token.strip():
        st.info(\"Se intentará obtener el FIX (serie SIE configurable) para las fechas mostradas. Por defecto se usa SF43718 (FIX USD).\" )
        serie_fix = st.text_input(\"Serie SIE del FIX\", value=\"SF43718\")
        if st.button(\"📥 Traer FIX y generar\"):
            try:
                df_fix = sie_fetch_series_range(serie_fix, fechas[0], fechas[-1], banxico_token)
                # Empatar a las 6 fechas; si falta alguna, se deja None
                mapa = {r[\"fecha\"]: r[\"valor\"] for _, r in df_fix.iterrows()}
                fix_vals = [mapa.get(d) for d in fechas]
                st.write(pd.DataFrame({\"Fecha\": [d.strftime(\"%Y-%m-%d\") for d in fechas], \"FIX\": fix_vals}))
                nombre = f\"indicadores_{hoy:%Y-%m-%d}.xlsx\"
                build_and_download_wb(fechas, fix_vals, spread_total, nombre)
            except Exception as e:
                st.error(f\"Fallo al consultar Banxico: {e}\")
    else:
        st.info(\"Captura manual del FIX para las 6 columnas (Día 1..Día actual).\" )
        cols = st.columns(6)
        fix_vals = []
        for i, c in enumerate(cols):
            with c:
                fix_vals.append(st.number_input(f\"FIX {fechas[i].strftime('%Y-%m-%d')}\", value=0.0, step=0.0001, format=\"%.5f\", key=f\"fix_{i}\"))
        if st.button(\"📝 Generar Excel con FIX capturado\"):
            nombre = f\"indicadores_{hoy:%Y-%m-%d}.xlsx\"
            build_and_download_wb(fechas, fix_vals, spread_total, nombre)

# ===============================
# Flujo 2: Corregir un Excel existente
# ===============================
else:
    st.subheader(\"🩹 Corregir Excel existente (Compra/Venta)\")
    archivo = st.file_uploader(\"Sube el Excel a corregir (debe contener la hoja 'Indicadores')\", type=[\"xlsx\"])
    if archivo is not None:
        try:
            wb = load_workbook(io.BytesIO(archivo.read()))
            if \"Indicadores\" not in wb.sheetnames:
                st.error(\"No se encontró la hoja 'Indicadores'.\")
            else:
                ws = wb[\"Indicadores\"]
                # Intentar leer fechas del encabezado fila 3, B..G (opcional)
                fechas = []
                for c in range(2, 8):
                    v = ws.cell(row=3, column=c).value
                    try:
                        # soporta strings YYYY-MM-DD
                        if isinstance(v, str):
                            fechas.append(dt.datetime.fromisoformat(v).date())
                        elif isinstance(v, (dt.date, dt.datetime)):
                            fechas.append(v.date() if isinstance(v, dt.datetime) else v)
                        else:
                            fechas.append(None)
                    except Exception:
                        fechas.append(None)

                fix_vals = leer_fix_de_excel(ws, row_fix=7, col_start=2, col_end=7)
                compra, venta = calcular_compra_venta_desde_fix(fix_vals, spread_pct=spread_total/100.0)

                # Escribir compra (fila 9) y venta (fila 10)
                escribir_valores_en_fila(ws, 9, compra, col_start=2)
                escribir_valores_en_fila(ws, 10, venta, col_start=2)

                # Previa
                prev = pd.DataFrame({
                    \"Col\": [\"B\",\"C\",\"D\",\"E\",\"F\",\"G\"],
                    \"Fecha\": [f.strftime(\"%Y-%m-%d\") if isinstance(f, dt.date) else \"\" for f in fechas],
                    \"FIX\": fix_vals,
                    \"Compra(calc)\": compra,
                    \"Venta(calc)\": venta
                })
                st.write(prev)

                out = io.BytesIO()
                wb.save(out); out.seek(0)
                hoy = now_mx().date()
                st.download_button(\"⬇️ Descargar Excel corregido\", data=out, file_name=f\"indicadores_corregido_{hoy:%Y-%m-%d}.xlsx\", mime=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\")
        except Exception as e:
            st.error(f\"No se pudo abrir o procesar el archivo: {e}\")
