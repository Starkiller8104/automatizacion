
# -*- coding: utf-8 -*-
"""
Indicadores IMEMSA (Streamlit)
- Generar o corregir Excel con FIX, Compra/Venta (a partir de FIX con spread), UDIS y TIIE (objetivo, 28, 91, 182).
- Evita que Compra/Venta sean iguales y que las TIIE se repitan por error de mapeo.
"""

import io
import math
import datetime as dt
from typing import List, Tuple, Optional, Dict

import pytz
import requests
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# -----------------------------
# Config
# -----------------------------
st.set_page_config(page_title="Indicadores IMEMSA", layout="wide")
st.title("Indicadores IMEMSA - Generar/Corregir Excel")
st.caption("FIX, Compra/Venta, UDIS y TIIE con SIE Banxico.")

TZ_MX = pytz.timezone("America/Mexico_City")
BASE = "https://www.banxico.org.mx/SieAPIRest/service/v1"

def now_mx() -> dt.datetime:
    return dt.datetime.now(TZ_MX)

# -----------------------------
# Utils
# -----------------------------
def to_float_safe(x) -> Optional[float]:
    try:
        if x is None:
            return None
        if isinstance(x, float) and math.isnan(x):
            return None
        return float(str(x).replace(",", "").strip())
    except Exception:
        return None

def calendar_days(end_date: dt.date, n_days: int) -> List[dt.date]:
    start = end_date - dt.timedelta(days=n_days - 1)
    return [start + dt.timedelta(days=i) for i in range(n_days)]

# -----------------------------
# Banxico SIE
# -----------------------------
def sie_fetch_series_range(series_ids: List[str], start: dt.date, end: dt.date, token: str) -> Dict[str, pd.DataFrame]:
    sid = ",".join(series_ids)
    url = f"{BASE}/series/{sid}/datos/{start:%Y-%m-%d}/{end:%Y-%m-%d}"
    r = requests.get(url, headers={"Bmx-Token": token.strip()}, timeout=30)
    r.raise_for_status()
    js = r.json()
    out = {}
    for s in js.get("bmx", {}).get("series", []):
        s_id = s.get("idSerie")
        rows = []
        for d in s.get("datos", []):
            f = d.get("fecha"); v = d.get("dato")
            if not f:
                continue
            try:
                if "/" in f:
                    fecha = dt.datetime.strptime(f, "%d/%m/%Y").date()
                else:
                    fecha = dt.datetime.fromisoformat(f).date()
            except Exception:
                continue
            val = None if v in (None, "", "N/E") else to_float_safe(v)
            if val is not None:
                rows.append({"fecha": fecha, "valor": val})
        out[s_id] = pd.DataFrame(rows).sort_values("fecha").reset_index(drop=True)
    return out

def sie_opportuno(series_ids: List[str], token: str) -> Dict[str, Optional[float]]:
    sid = ",".join(series_ids)
    url = f"{BASE}/series/{sid}/datos/oportuno"
    r = requests.get(url, headers={"Bmx-Token": token.strip()}, timeout=20)
    r.raise_for_status()
    js = r.json()
    out = {}
    for s in js.get("bmx", {}).get("series", []):
        serie_id = s.get("idSerie")
        dato = None
        datos = s.get("datos", [])
        if datos:
            d0 = datos[0].get("dato")
            if d0 not in (None, "", "N/E"):
                dato = to_float_safe(d0)
        out[serie_id] = dato
    return out

# -----------------------------
# Business logic
# -----------------------------
def calcular_compra_venta_desde_fix(fix_vals: List[Optional[float]], spread_total_pct: float = 0.40) -> Tuple[List[Optional[float]], List[Optional[float]]]:
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
    if df is None or df.empty:
        return [None] * len(fechas)
    s = df.set_index("fecha")["valor"].sort_index()
    vals = []
    last = None
    for d in fechas:
        if d in s.index:
            last = s.loc[d]
            vals.append(float(last))
        else:
            vals.append(float(last) if (forward_fill and last is not None) else None)
    return vals

# -----------------------------
# Excel helpers
# -----------------------------
THIN = Side(border_style="thin", color="808080")
BORDER_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def ensure_hoja_indicadores(wb: Workbook) -> None:
    if "Indicadores" not in wb.sheetnames:
        wb.create_sheet("Indicadores")
    ws = wb["Indicadores"]
    headers = ["Dia 1", "Dia 2", "Dia 3", "Dia 4", "Dia 5", "Dia actual"]
    ws.cell(row=2, column=1, value="Concepto").font = Font(bold=True)
    for i, h in enumerate(headers, start=2):
        c = ws.cell(row=2, column=i, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")
        c.border = BORDER_THIN
        ws.column_dimensions[get_column_letter(i)].width = 16
    ws.column_dimensions["A"].width = 40
    labels = [
        (5,  "UDIS (valor)"),
        (7,  "Dolar / Pesos (FIX)"),
        (9,  "Dolar Americano (Compra)"),
        (10, "Dolar Americano (Venta)"),
        (12, "Tasa objetivo"),
        (13, "TIIE 28 dias"),
        (14, "TIIE 91 dias"),
        (15, "TIIE 182 dias"),
    ]
    for row, text in labels:
        ws.cell(row=row, column=1, value=text).font = Font(bold=True)

def escribir_fechas_encabezado(ws, fechas: List[dt.date]) -> None:
    for idx, d in enumerate(fechas, start=2):
        ws.cell(row=3, column=idx, value=d.strftime("%Y-%m-%d")).alignment = Alignment(horizontal="center")

def escribir_series(ws, row: int, vals: List[Optional[float]], col_start: int = 2) -> None:
    for i, v in enumerate(vals):
        ws.cell(row=row, column=col_start + i, value=v)

# -----------------------------
# Sidebar
# -----------------------------
with st.sidebar:
    st.header("Parametros")
    modo = st.radio("Modo", ["Generar Excel", "Corregir Excel"], index=0)

    st.subheader("Banxico API")
    token = st.text_input("Bmx-Token", type="password", help="Necesario para consultar SIE")

    st.subheader("Series SIE (editables)")
    serie_fix = st.text_input("FIX (USD/MXN)", value="SF43718")
    serie_udis = st.text_input("UDIS", value="SP68257")
    serie_obj = st.text_input("Tasa objetivo", value="SF61745")
    serie_tiie_28 = st.text_input("TIIE 28 dias", value="SF60648")
    serie_tiie_91 = st.text_input("TIIE 91 dias", value="SF60649")
    serie_tiie_182 = st.text_input("TIIE 182 dias (verifica ID)", value="")

    st.markdown("---")
    spread_total = st.slider("Spread total Compra/Venta (%)", 0.10, 1.50, 0.40, 0.05)
    forward_fill_udis = st.checkbox("UDIS: forward-fill si falta ultimo dato", value=True)

# -----------------------------
# Flow A: Generate
# -----------------------------
if modo == "Generar Excel":
    st.subheader("Generar Excel desde cero")
    hoy = now_mx().date()
    fechas = calendar_days(hoy, 6)

    st.write(pd.DataFrame({"Fecha": [f.strftime("%Y-%m-%d") for f in fechas]}))

    if not token:
        st.warning("Ingresa tu Bmx-Token en la barra lateral para consultar SIE.")
        st.stop()

    ids_rango = [s for s in [serie_fix, serie_udis] if s]
    data_rango = sie_fetch_series_range(ids_rango, fechas[0], fechas[-1], token)

    df_fix = data_rango.get(serie_fix, pd.DataFrame(columns=["fecha","valor"])) if serie_fix else None
    df_udis = data_rango.get(serie_udis, pd.DataFrame(columns=["fecha","valor"])) if serie_udis else None

    fix_vals = align_to_calendar(df_fix, fechas, forward_fill=False) if df_fix is not None else [None]*len(fechas)
    udis_vals = align_to_calendar(df_udis, fechas, forward_fill=forward_fill_udis) if df_udis is not None else [None]*len(fechas)

    compra, venta = calcular_compra_venta_desde_fix(fix_vals, spread_total_pct=spread_total)

    ids_tiie = [s for s in [serie_obj, serie_tiie_28, serie_tiie_91, serie_tiie_182] if s]
    map_tiie = sie_opportuno(ids_tiie, token) if ids_tiie else {}

    val_obj = map_tiie.get(serie_obj)
    val_28  = map_tiie.get(serie_tiie_28)
    val_91  = map_tiie.get(serie_tiie_91)
    val_182 = map_tiie.get(serie_tiie_182) if serie_tiie_182 else None

    serie_obj_vals = [val_obj]*len(fechas)
    serie_28_vals  = [val_28]*len(fechas)
    serie_91_vals  = [val_91]*len(fechas)
    serie_182_vals = [val_182]*len(fechas) if serie_tiie_182 else [None]*len(fechas)

    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
    ensure_hoja_indicadores(wb)
    ws = wb["Indicadores"]
    escribir_fechas_encabezado(ws, fechas)

    escribir_series(ws, 5,  udis_vals)
    escribir_series(ws, 7,  fix_vals)
    escribir_series(ws, 9,  compra)
    escribir_series(ws, 10, venta)
    escribir_series(ws, 12, serie_obj_vals)
    escribir_series(ws, 13, serie_28_vals)
    escribir_series(ws, 14, serie_91_vals)
    escribir_series(ws, 15, serie_182_vals)

    out = io.BytesIO()
    wb.save(out); out.seek(0)
    st.download_button("Descargar Excel", data=out, file_name=f"indicadores_{hoy:%Y-%m-%d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -----------------------------
# Flow B: Fix existing
# -----------------------------
else:
    st.subheader("Corregir Excel existente")
    archivo = st.file_uploader("Sube tu Excel (hoja 'Indicadores')", type=["xlsx"])
    if archivo is not None:
        try:
            wb = load_workbook(io.BytesIO(archivo.read()))
            if "Indicadores" not in wb.sheetnames:
                st.error("No se encontro la hoja 'Indicadores'.")
            else:
                ws = wb["Indicadores"]
                fechas = []
                for c in range(2, 8):
                    v = ws.cell(row=3, column=c).value
                    d = None
                    if isinstance(v, str):
                        try:
                            d = dt.datetime.fromisoformat(v).date()
                        except Exception:
                            d = None
                    elif isinstance(v, dt.datetime):
                        d = v.date()
                    elif isinstance(v, dt.date):
                        d = v
                    fechas.append(d)
                if any(f is None for f in fechas):
                    hoy = now_mx().date()
                    fechas = calendar_days(hoy, 6)

                fix_vals = [to_float_safe(ws.cell(row=7, column=2+i).value) for i in range(6)]
                compra, venta = calcular_compra_venta_desde_fix(fix_vals, spread_total_pct=spread_total)
                escribir_series(ws, 9, compra)
                escribir_series(ws, 10, venta)

                if token and serie_udis:
                    data_rango = sie_fetch_series_range([serie_udis], fechas[0], fechas[-1], token)
                    dfu = data_rango.get(serie_udis, pd.DataFrame(columns=["fecha","valor"]))
                    udis_vals = align_to_calendar(dfu, fechas, forward_fill=forward_fill_udis)
                    escribir_series(ws, 5, udis_vals)

                if token:
                    ids_tiie = [s for s in [serie_obj, serie_tiie_28, serie_tiie_91, serie_tiie_182] if s]
                    m = sie_opportuno(ids_tiie, token) if ids_tiie else {}
                    val_obj = m.get(serie_obj); val_28 = m.get(serie_tiie_28); val_91 = m.get(serie_tiie_91)
                    val_182 = m.get(serie_tiie_182) if serie_tiie_182 else None
                    escribir_series(ws, 12, [val_obj]*6)
                    escribir_series(ws, 13, [val_28]*6)
                    escribir_series(ws, 14, [val_91]*6)
                    if serie_tiie_182:
                        escribir_series(ws, 15, [val_182]*6)

                out = io.BytesIO()
                wb.save(out); out.seek(0)
                hoy = now_mx().date()
                st.download_button("Descargar Excel corregido", data=out,
                                   file_name=f"indicadores_corregido_{hoy:%Y-%m-%d}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"No se pudo abrir o procesar el archivo: {e}")
