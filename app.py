
# -*- coding: utf-8 -*-


import io
import pytz
import math
import requests
import datetime as dt
from typing import List, Tuple, Optional, Dict

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter

import streamlit as st

# ===============================
# Configuración de página
# ===============================
st.set_page_config(page_title=\"Indicadores IMEMSA\", layout=\"wide\")
st.title(\"📊 Indicadores IMEMSA – Generar/Corregir Excel\")
st.caption(\"Cálculo de FIX, Compra/Venta, UDIS y TIIE (objetivo/28/91/182) con SIE Banxico.\")

TZ_MX = pytz.timezone(\"America/Mexico_City\")

def now_mx() -> dt.datetime:
    return dt.datetime.now(TZ_MX)

# ===============================
# Helpers
# ===============================
def to_float_safe(x) -> Optional[float]:
    try:
        if x is None:
            return None
        if isinstance(x, float) and math.isnan(x):
            return None
        return float(str(x).replace(\",\", \"\").strip())
    except Exception:
        return None

def business_days_mx(end_date: dt.date, n_days: int) -> List[dt.date]:
    out = []
    d = end_date
    while len(out) < n_days:
        if d.weekday() < 5:
            out.append(d)
        d = d - dt.timedelta(days=1)
    return list(reversed(out))

def calendar_days(end_date: dt.date, n_days: int) -> List[dt.date]:
    start = end_date - dt.timedelta(days=n_days - 1)
    return [start + dt.timedelta(days=i) for i in range(n_days)]

# ===============================
# Banxico SIE API
# ===============================
BASE = \"https://www.banxico.org.mx/SieAPIRest/service/v1\"

def sie_fetch_series_range(series_ids: List[str], start: dt.date, end: dt.date, token: str) -> Dict[str, pd.DataFrame]:
    \"\"\"Consulta rango para múltiples series. Regresa dict idSerie -> DataFrame(fecha, valor).\"\"\"
    sid = \",\".join(series_ids)
    url = f\"{BASE}/series/{sid}/datos/{start:%Y-%m-%d}/{end:%Y-%m-%d}\"
    r = requests.get(url, headers={\"Bmx-Token\": token.strip()}, timeout=30)
    r.raise_for_status()
    js = r.json()
    series = js.get(\"bmx\", {}).get(\"series\", [])
    out = {}
    for s in series:
        s_id = s.get(\"idSerie\")
        datos = s.get(\"datos\", [])
        rows = []
        for d in datos:
            f = d.get(\"fecha\")
            v = d.get(\"dato\")
            # fechas pueden venir en dd/mm/aaaa
            try:
                if \"/\" in f:
                    fecha = dt.datetime.strptime(f, \"%d/%m/%Y\").date()
                else:
                    fecha = dt.datetime.fromisoformat(f).date()
            except Exception:
                continue
            val = None if v in (None, \"\", \"N/E\") else to_float_safe(v)
            if val is not None:
                rows.append({\"fecha\": fecha, \"valor\": val})
        out[s_id] = pd.DataFrame(rows).sort_values(\"fecha\").reset_index(drop=True)
    return out

def sie_opportuno(series_ids: List[str], token: str) -> Dict[str, Optional[float]]:
    \"\"\"Último dato publicado por serie. Regresa dict idSerie -> valor (float o None).\"\"\"
    sid = \",\".join(series_ids)
    url = f\"{BASE}/series/{sid}/datos/oportuno\"
    r = requests.get(url, headers={\"Bmx-Token\": token.strip()}, timeout=20)
    r.raise_for_status()
    js = r.json()
    out = {}
    for s in js.get(\"bmx\", {}).get(\"series\", []):
        serie_id = s.get(\"idSerie\")
        datos = s.get(\"datos\", [])
        val = None
        if datos:
            dato = datos[0].get(\"dato\")
            if dato not in (None, \"\", \"N/E\"):
                val = to_float_safe(dato)
        out[serie_id] = val
    return out

# ===============================
# Lógica de negocio
# ===============================
def calcular_compra_venta_desde_fix(fix_vals: List[Optional[float]], spread_total_pct: float = 0.40) -> Tuple[List[Optional[float]], List[Optional[float]]]:
    \"\"\"Compra = FIX*(1 - spread/2); Venta = FIX*(1 + spread/2); spread en % total (p.ej. 0.40).\"\"\"
    half = (spread_total_pct / 100.0) / 2.0
    compra, venta = [], []
    for v in fix_vals:
        vf = to_float_safe(v)
        if vf is None:
            compra.append(None); venta.append(None)
        else:
            compra.append(round(vf * (1.0 - half), 5))
            venta.append(round(vf * (1.0 + half), 5))
    return compra, venta

def align_to_calendar(df: pd.DataFrame, fechas: List[dt.date], forward_fill: bool) -> List[Optional[float]]:
    \"\"\"Alinea DataFrame(fecha,valor) a una lista de fechas calendario; opcionalmente forward-fill.\"\"\"
    if df is None or df.empty:
        return [None] * len(fechas)
    s = df.set_index(\"fecha\")[\"valor\"].sort_index()
    vals = []
    last = None
    for d in fechas:
        if d in s.index:
            last = s.loc[d]
            vals.append(float(last))
        else:
            if forward_fill and last is not None:
                vals.append(float(last))
            else:
                vals.append(None)
    return vals

# ===============================
# Excel helpers
# ===============================
THIN = Side(border_style=\"thin\", color=\"808080\")
BORDER_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def ensure_hoja_indicadores(wb: Workbook) -> None:
    if \"Indicadores\" not in wb.sheetnames:
        wb.create_sheet(\"Indicadores\")
    ws = wb[\"Indicadores\"]
    headers = [\"Día 1\", \"Día 2\", \"Día 3\", \"Día 4\", \"Día 5\", \"Día actual\"]
    ws.cell(row=2, column=1, value=\"Concepto\").font = Font(bold=True)
    for i, h in enumerate(headers, start=2):
        c = ws.cell(row=2, column=i, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal=\"center\")
        c.border = BORDER_THIN
        ws.column_dimensions[get_column_letter(i)].width = 16
    ws.column_dimensions[\"A\"].width = 40

    labels = [
        (5,  \"UDIS (valor)\"),
        (7,  \"Dólar / Pesos (FIX)\"),
        (9,  \"Dólar Americano (Compra)\"),
        (10, \"Dólar Americano (Venta)\"),
        (12, \"Tasa objetivo\"),
        (13, \"TIIE 28 días\"),
        (14, \"TIIE 91 días\"),
        (15, \"TIIE 182 días\"),
    ]
    for row, text in labels:
        ws.cell(row=row, column=1, value=text).font = Font(bold=True)

def escribir_fechas_encabezado(ws, fechas: List[dt.date]) -> None:
    for idx, d in enumerate(fechas, start=2):
        ws.cell(row=3, column=idx, value=d.strftime(\"%Y-%m-%d\")).alignment = Alignment(horizontal=\"center\")

def escribir_series(ws, row: int, vals: List[Optional[float]], col_start: int = 2) -> None:
    for i, v in enumerate(vals):
        ws.cell(row=row, column=col_start + i, value=v)

# ===============================
# Sidebar / Parámetros
# ===============================
with st.sidebar:
    st.header(\"⚙️ Parámetros\")
    modo = st.radio(\"Modo\", [\"Generar Excel\", \"Corregir Excel\"], index=0)

    st.subheader(\"Banxico API\")
    token = st.text_input(\"Bmx-Token\", type=\"password\", help=\"Requerido para consultar SIE\")

    st.subheader(\"Series SIE (editables)\")
    serie_fix = st.text_input(\"FIX (USD/MXN)\", value=\"SF43718\")
    serie_udis = st.text_input(\"UDIS\", value=\"SP68257\")
    serie_obj = st.text_input(\"Tasa objetivo\", value=\"SF61745\")
    serie_tiie_28 = st.text_input(\"TIIE 28 días\", value=\"SF60648\")
    serie_tiie_91 = st.text_input(\"TIIE 91 días\", value=\"SF60649\")
    serie_tiie_182 = st.text_input(\"TIIE 182 días (verifica ID)\", value=\"\", help=\"Déjalo vacío si no lo conoces; puedes llenarlo luego.\")

    st.markdown(\"---\")
    spread_total = st.slider(\"Spread total Compra/Venta (%)\", 0.10, 1.50, 0.40, 0.05,
                             help=\"0.40% ⇒ Compra = FIX - 0.20% y Venta = FIX + 0.20%\" )
    forward_fill_udis = st.checkbox(\"UDIS: usar forward-fill cuando falte el último dato\", value=True)

# ===============================
# Flujo A: Generar Excel
# ===============================
if modo == \"Generar Excel\":
    st.subheader(\"🧰 Generar Excel desde cero\")
    hoy = now_mx().date()
    # Para UDIS usamos calendario; para los demás, el archivo queda en 6 columnas uniformes
    fechas = calendar_days(hoy, 6)

    col1, col2 = st.columns([1,1])
    with col1:
        st.write(pd.DataFrame({\"Fecha\": [f.strftime(\"%Y-%m-%d\") for f in fechas]}))

    if not token:
        st.warning(\"Ingresa tu Bmx-Token en la barra lateral para consultar datos de SIE.\")
        st.stop()

    # Consultar FIX y UDIS en rango para alinear
    ids_rango = [serie_fix, serie_udis]
    ids_rango = [s for s in ids_rango if s]  # no vacíos
    data_rango = sie_fetch_series_range(ids_rango, fechas[0], fechas[-1], token)

    df_fix = data_rango.get(serie_fix, pd.DataFrame(columns=[\"fecha\",\"valor\"])) if serie_fix else None
    df_udis = data_rango.get(serie_udis, pd.DataFrame(columns=[\"fecha\",\"valor\"])) if serie_udis else None

    fix_vals = align_to_calendar(df_fix, fechas, forward_fill=False) if df_fix is not None else [None]*len(fechas)
    udis_vals = align_to_calendar(df_udis, fechas, forward_fill=forward_fill_udis) if df_udis is not None else [None]*len(fechas)

    compra, venta = calcular_compra_venta_desde_fix(fix_vals, spread_total_pct=spread_total)

    # TIIE/Objetivo oportuno (solo últimos valores, se repiten en las 6 columnas para mostrar)
    ids_tiie = [serie_obj, serie_tiie_28, serie_tiie_91] + ([serie_tiie_182] if serie_tiie_182 else [])
    ids_tiie = [s for s in ids_tiie if s]
    map_tiie = sie_opportuno(ids_tiie, token) if ids_tiie else {}

    val_obj = map_tiie.get(serie_obj)
    val_28  = map_tiie.get(serie_tiie_28)
    val_91  = map_tiie.get(serie_tiie_91)
    val_182 = map_tiie.get(serie_tiie_182) if serie_tiie_182 else None

    serie_obj_vals = [val_obj]*len(fechas)
    serie_28_vals  = [val_28]*len(fechas)
    serie_91_vals  = [val_91]*len(fechas)
    serie_182_vals = [val_182]*len(fechas) if serie_tiie_182 else [None]*len(fechas)

    # Construir Excel
    wb = Workbook()
    # remover hoja inicial
    if \"Sheet\" in wb.sheetnames:
        wb.remove(wb[\"Sheet\"])
    ensure_hoja_indicadores(wb)
    ws = wb[\"Indicadores\"]
    escribir_fechas_encabezado(ws, fechas)

    escribir_series(ws, 5,  udis_vals)
    escribir_series(ws, 7,  fix_vals)
    escribir_series(ws, 9,  compra)
    escribir_series(ws, 10, venta)
    escribir_series(ws, 12, serie_obj_vals)
    escribir_series(ws, 13, serie_28_vals)
    escribir_series(ws, 14, serie_91_vals)
    escribir_series(ws, 15, serie_182_vals)

    # Previa
    prev = pd.DataFrame({
        \"Fecha\": [f.strftime(\"%Y-%m-%d\") for f in fechas],
        \"UDIS\": udis_vals,
        \"FIX\": fix_vals,
        \"Compra\": compra,
        \"Venta\": venta,
        \"Obj\": serie_obj_vals,
        \"TIIE28\": serie_28_vals,
        \"TIIE91\": serie_91_vals,
        \"TIIE182\": serie_182_vals,
    })
    st.dataframe(prev, use_container_width=True)

    out = io.BytesIO()
    wb.save(out); out.seek(0)
    st.download_button(\"⬇️ Descargar Excel\", data=out, file_name=f\"indicadores_{hoy:%Y-%m-%d}.xlsx\",
                       mime=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\")

# ===============================
# Flujo B: Corregir Excel existente
# ===============================
else:
    st.subheader(\"🩹 Corregir Excel existente (Compra/Venta, TIIE duplicadas, UDIS vacíos)\")
    archivo = st.file_uploader(\"Sube tu Excel (hoja 'Indicadores')\", type=[\"xlsx\"])

    if archivo is not None:
        try:
            wb = load_workbook(io.BytesIO(archivo.read()))
            if \"Indicadores\" not in wb.sheetnames:
                st.error(\"No se encontró la hoja 'Indicadores'.\")
            else:
                ws = wb[\"Indicadores\"]
                # Leer fechas del encabezado fila 3 B..G; si no están, usamos hoy-5..hoy
                fechas = []
                for c in range(2, 8):
                    v = ws.cell(row=3, column=c).value
                    d = None
                    if isinstance(v, str):
                        try:
                            d = dt.datetime.fromisoformat(v).date()
                        except Exception:
                            pass
                    elif isinstance(v, dt.datetime):
                        d = v.date()
                    elif isinstance(v, dt.date):
                        d = v
                    fechas.append(d)
                if any(f is None for f in fechas):
                    hoy = now_mx().date()
                    fechas = calendar_days(hoy, 6)

                # Recalcular Compra/Venta desde FIX fila 7
                fix_vals = [to_float_safe(ws.cell(row=7, column=2+i).value) for i in range(6)]
                compra, venta = calcular_compra_venta_desde_fix(fix_vals, spread_total_pct=spread_total)
                escribir_series(ws, 9, compra)
                escribir_series(ws, 10, venta)

                # UDIS en fila 5: si faltan últimos días, traer rango y alinear (opcional FF)
                if token and serie_udis:
                    data_rango = sie_fetch_series_range([serie_udis], fechas[0], fechas[-1], token)
                    dfu = data_rango.get(serie_udis, pd.DataFrame(columns=[\"fecha\",\"valor\"]))
                    udis_vals = align_to_calendar(dfu, fechas, forward_fill=forward_fill_udis)
                    escribir_series(ws, 5, udis_vals)

                # TIIE/Objetivo oportuno (evitar copias): cada serie con su id
                if token:
                    ids_tiie = [serie_obj, serie_tiie_28, serie_tiie_91] + ([serie_tiie_182] if serie_tiie_182 else [])
                    ids_tiie = [s for s in ids_tiie if s]
                    m = sie_opportuno(ids_tiie, token) if ids_tiie else {}
                    val_obj = m.get(serie_obj); val_28 = m.get(serie_tiie_28); val_91 = m.get(serie_tiie_91)
                    val_182 = m.get(serie_tiie_182) if serie_tiie_182 else None
                    escribir_series(ws, 12, [val_obj]*6)
                    escribir_series(ws, 13, [val_28]*6)
                    escribir_series(ws, 14, [val_91]*6)
                    if serie_tiie_182:
                        escribir_series(ws, 15, [val_182]*6)

                # Previa
                prev = pd.DataFrame({
                    \"Fecha\": [f.strftime(\"%Y-%m-%d\") for f in fechas],
                    \"FIX\": fix_vals,
                    \"Compra\": compra,
                    \"Venta\": venta,
                })
                st.dataframe(prev, use_container_width=True)

                out = io.BytesIO()
                wb.save(out); out.seek(0)
                hoy = now_mx().date()
                st.download_button(\"⬇️ Descargar Excel corregido\", data=out, file_name=f\"indicadores_corregido_{hoy:%Y-%m-%d}.xlsx\",
                                   mime=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\")
        except Exception as e:
            st.error(f\"No se pudo abrir o procesar el archivo: {e}\")
