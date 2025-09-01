# -*- coding: utf-8 -*-
"""
Indicadores IMEMSA (Streamlit) - con LOGIN
- Mantiene una ventana de login (contraseña) antes de mostrar la app.
- Genera o corrige Excel con: FIX, Compra/Venta (desde FIX con spread), UDIS, y TIIE (objetivo/28/91/182).
- Corrige que Compra/Venta salgan iguales y que las TIIE se repitan por mapeo.
"""

import io
import os
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

# ===============================
# Configuración de página
# ===============================
st.set_page_config(page_title="Indicadores IMEMSA", layout="wide")
st.title("Indicadores IMEMSA")

TZ_MX = pytz.timezone("America/Mexico_City")
BASE = "https://www.banxico.org.mx/SieAPIRest/service/v1"

# ===============================
# LOGIN
# ===============================
DEFAULT_APP_PASSWORD = os.environ.get("APP_PASSWORD", None)
if DEFAULT_APP_PASSWORD is None:
    # Si no hay variable de entorno ni secret, usa st.secrets si existe; de lo contrario fija valor demo.
    try:
        DEFAULT_APP_PASSWORD = st.secrets.get("APP_PASSWORD", "demo")
    except Exception:
        DEFAULT_APP_PASSWORD = "demo"

if "auth_ok" not in st.session_state:
    st.session_state["auth_ok"] = False

def login_gate():
    st.subheader("Acceso")
    with st.form("login_form"):
        pwd = st.text_input("Contraseña", type="password")
        ok = st.form_submit_button("Entrar")
        if ok:
            if pwd == DEFAULT_APP_PASSWORD:
                st.session_state["auth_ok"] = True
                st.success("Acceso concedido")
            else:
                st.error("Contraseña incorrecta")

if not st.session_state["auth_ok"]:
    login_gate()
    st.stop()

st.caption("Sesión iniciada ✅")

# ===============================
# Utilidades
# ===============================
def now_mx() -> dt.datetime:
    return dt.datetime.now(TZ_MX)

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

# ===============================
# Banxico SIE
# ===============================
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

# ===============================
# Lógica de negocio
# ===============================
def calcular_compra_venta_desde_fix(fix_vals: List[Optional[float]], spread_total_pct: float = 0.40) -> Tuple[List[Optional[float]], List[Optional[float]]]:
    """
    Compra = FIX*(1 - spread/2); Venta = FIX*(1 + spread/2); spread_total_pct=0.40 => +/-0.20%
    """
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
    """
    Alinea por días de calendario; si forward_fill=True, usa el último valor cuando falte el día más reciente.
    """
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

# ===============================
# Excel helpers
# ===============================
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

# ===============================
# Sidebar (conserva opciones)
# ===============================
with st.sidebar:
    st.header("Parámetros")
    modo = st.radio("Modo", ["Generar Excel", "Corregir Excel"], index=0)

    st.subheader("Banxico API")
    # Permite usar secreto BANXICO_TOKEN si existe
    try:
        default_token = st.secrets.get("BANXICO_TOKEN", "")
    except Exception:
        default_token = ""
    token = st.text_input("Bmx-Token", value=default_token, type="password", help="Requerido para consultar SIE")

    st.subheader("Series SIE (editables)")
    serie_fix = st.text_input("FIX (USD/MXN)", value="SF43718")
    serie_udis = st.text_input("UDIS", value="SP68257")
    serie_obj = st.text_input("Tasa objetivo", value="SF61745")
    serie_tiie_28 = st.text_input("TIIE 28 dias", value="SF60648")
    serie_tiie_91 = st.text_input("TIIE 91 dias", value="SF60649")
    serie_tiie_182 = st.text_input("TIIE 182 dias (verifica ID)", value="")

    st.markdown("---")
    spread_total = st.slider("Spread total Compra/Venta (%)", 0.10, 1.50, 0.40, 0.05,
                             help="0.40% ⇒ Compra = FIX - 0.20% y Venta = FIX + 0.20%")
    forward_fill_udis = st.checkbox("UDIS: forward-fill si falta el último dato", value=True)

# ===============================
# Flujo A: Generar Excel
# ===============================
if modo == "Generar Excel":
    st.subheader("Generar Excel desde cero")
    hoy = now_mx().date()
    fechas = calendar_days(hoy, 6)  # usamos días de calendario para UDIS/FIX

    st.write(pd.DataFrame({"Fecha": [f.strftime("%Y-%m-%d") for f in fechas]}))

    if not token:
        st.warning("Ingresa tu Bmx-Token en la barra lateral para consultar SIE.")
        st.stop()

    # FIX y UDIS en rango para alinear a calendario
    ids_rango = [s for s in [serie_fix, serie_udis] if s]
    data_rango = sie_fetch_series_range(ids_rango, fechas[0], fechas[-1], token)
    df_fix = data_rango.get(serie_fix, pd.DataFrame(columns=["fecha","valor"])) if serie_fix else None
    df_udis = data_rango.get(serie_udis, pd.DataFrame(columns=["fecha","valor"])) if serie_udis else None

    fix_vals = [None]*len(fechas) if df_fix is None else align_to_calendar(df_fix, fechas, forward_fill=False)
    udis_vals = [None]*len(fechas) if df_udis is None else align_to_calendar(df_udis, fechas, forward_fill=forward_fill_udis)

    # Compra/Venta desde FIX
    compra, venta = calcular_compra_venta_desde_fix(fix_vals, spread_total_pct=spread_total)

    # TIIE y objetivo (oportuno)
    ids_tiie = [s for s in [serie_obj, serie_tiie_28, serie_tiie_91, serie_tiie_182] if s]
    map_tiie = sie_opportuno(ids_tiie, token) if ids_tiie else {}
    val_obj = map_tiie.get(serie_obj); val_28 = map_tiie.get(serie_tiie_28)
    val_91 = map_tiie.get(serie_tiie_91); val_182 = map_tiie.get(serie_tiie_182) if serie_tiie_182 else None

    serie_obj_vals = [val_obj]*len(fechas)
    serie_28_vals  = [val_28]*len(fechas)
    serie_91_vals  = [val_91]*len(fechas)
    serie_182_vals = [val_182]*len(fechas) if serie_tiie_182 else [None]*len(fechas)

    # Construcción de Excel
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
    ensure_hoja_indicadores(wb)
    ws = wb["Indicadores"]
    # encabezado de fechas
    for idx, d in enumerate(fechas, start=2):
        ws.cell(row=3, column=idx, value=d.strftime("%Y-%m-%d")).alignment = Alignment(horizontal="center")

    # escribir series
    for row, vals in [(5, udis_vals), (7, fix_vals), (9, compra), (10, venta),
                      (12, serie_obj_vals), (13, serie_28_vals), (14, serie_91_vals), (15, serie_182_vals)]:
        for i, v in enumerate(vals, start=2):
            ws.cell(row=row, column=i, value=v)

    out = io.BytesIO()
    wb.save(out); out.seek(0)
    st.download_button("Descargar Excel", data=out, file_name=f"indicadores_{hoy:%Y-%m-%d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ===============================
# Flujo B: Corregir Excel existente
# ===============================
else:
    st.subheader("Corregir Excel existente")
    archivo = st.file_uploader("Sube tu Excel (hoja 'Indicadores')", type=["xlsx"])

    if archivo is not None:
        try:
            wb = load_workbook(io.BytesIO(archivo.read()))
            if "Indicadores" not in wb.sheetnames:
                st.error("No se encontró la hoja 'Indicadores'.")
            else:
                ws = wb["Indicadores"]
                # fechas en encabezado (fila 3, B..G); si faltan, usa hoy-5..hoy
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

                # Recalcular Compra/Venta desde FIX
                fix_vals = [to_float_safe(ws.cell(row=7, column=2+i).value) for i in range(6)]
                compra, venta = calcular_compra_venta_desde_fix(fix_vals, spread_total_pct=spread_total)
                for i, v in enumerate(compra, start=2):
                    ws.cell(row=9, column=i, value=v)
                for i, v in enumerate(venta, start=2):
                    ws.cell(row=10, column=i, value=v)

                # UDIS: si faltan últimos días, consulta rango y alinea (con forward-fill opcional)
                if token and serie_udis:
                    data_rango = sie_fetch_series_range([serie_udis], fechas[0], fechas[-1], token)
                    dfu = data_rango.get(serie_udis, pd.DataFrame(columns=["fecha","valor"]))
                    udis_vals = align_to_calendar(dfu, fechas, forward_fill=forward_fill_udis)
                    for i, v in enumerate(udis_vals, start=2):
                        ws.cell(row=5, column=i, value=v)

                # TIIE/Objetivo oportuno: cada serie con su id (evita duplicar)
                if token:
                    ids_tiie = [s for s in [serie_obj, serie_tiie_28, serie_tiie_91, serie_tiie_182] if s]
                    m = sie_opportuno(ids_tiie, token) if ids_tiie else {}
                    val_obj = m.get(serie_obj); val_28 = m.get(serie_tiie_28); val_91 = m.get(serie_tiie_91)
                    val_182 = m.get(serie_tiie_182) if serie_tiie_182 else None
                    for i in range(2, 8):
                        ws.cell(row=12, column=i, value=val_obj)
                        ws.cell(row=13, column=i, value=val_28)
                        ws.cell(row=14, column=i, value=val_91)
                        if serie_tiie_182:
                            ws.cell(row=15, column=i, value=val_182)

                out = io.BytesIO()
                wb.save(out); out.seek(0)
                hoy = now_mx().date()
                st.download_button("Descargar Excel corregido", data=out,
                                   file_name=f"indicadores_corregido_{hoy:%Y-%m-%d}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"No se pudo abrir o procesar el archivo: {e}")
