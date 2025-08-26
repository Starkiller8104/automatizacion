# -*- coding: utf-8 -*-
"""
Indicadores Económicos IMEMSA
- Login con contraseña por defecto "imemsa79" (secrets/env override)
- Barra lateral para ingresar tokens BANXICO/INEGI (se usan en caliente)
- Consulta SIE (FIX, EUR, JPY, CETES, UDIS) e INEGI (UMA)
- Agregador de noticias financieras (RSS)
- Genera Excel con hojas: Indicadores, Noticias, (opc) Gráficos y (opc) Datos crudos
"""

import io
import os
import re
import time
import base64
from datetime import datetime, timedelta
from pathlib import Path

import requests
from requests.adapters import HTTPAdapter, Retry
import feedparser
from bs4 import BeautifulSoup
from PIL import Image

import pandas as pd
import pytz
import streamlit as st
import xlsxwriter  # para generar el Excel en memoria

# =========================
#  Config & Login
# =========================
st.set_page_config(page_title="Indicadores Económicos", page_icon="📊", layout="wide")
CDMX = pytz.timezone("America/Mexico_City")

def _get_app_password() -> str:
    # 1) st.secrets
    try:
        return st.secrets["APP_PASSWORD"]
    except Exception:
        pass
    # 2) env var
    if os.getenv("APP_PASSWORD"):
        return os.getenv("APP_PASSWORD")
    # 3) respaldo
    return "imemsa79"

def _check_password() -> bool:
    if "auth_ok" not in st.session_state:
        st.session_state.auth_ok = False
    def _try_login():
        pw = st.session_state.get("password_input", "")
        st.session_state.auth_ok = (pw == _get_app_password())
        st.session_state.password_input = ""
    if st.session_state.auth_ok:
        return True
    st.title("🔒 Acceso restringido")
    st.text_input("Contraseña", type="password", key="password_input",
                  on_change=_try_login, placeholder="Escribe tu contraseña…")
    st.stop()

_check_password()

# =========================
#  Tokens (valores por defecto; puedes cambiarlos)
# =========================
BANXICO_TOKEN = "REEMPLAZA_AQUI_SI_QUIERES"  # si no lo ingresas en el sidebar
INEGI_TOKEN   = "REEMPLAZA_AQUI_SI_QUIERES"  # si no lo ingresas en el sidebar
FRED_TOKEN    = ""  # opcional, no usado en este ejemplo

# =========================
#  Utilidades generales
# =========================
def today_cdmx():
    return datetime.now(CDMX).date()

def now_ts():
    return datetime.now(CDMX).strftime("%Y-%m-%d %H:%M:%S")

def try_float(x):
    try:
        return float(str(x).replace(",", "").strip())
    except Exception:
        return None

@st.cache_data(show_spinner=False, ttl=60*30)
def http_session(timeout=15):
    s = requests.Session()
    retries = Retry(total=3, backoff_factor=0.8,
                    status_forcelist=[429, 500, 502, 503, 504],
                    allowed_methods=frozenset(["GET"]))
    s.mount("https://", HTTPAdapter(max_retries=retries))
    s.mount("http://", HTTPAdapter(max_retries=retries))
    orig = s.request
    def _patched(method, url, **kwargs):
        kwargs.setdefault("timeout", timeout)
        return orig(method, url, **kwargs)
    s.request = _patched
    return s

def format_ddmmyyyy(yyyy_mm_dd: str):
    # Fechas SIE vienen "YYYY-MM-DD"
    try:
        d = datetime.strptime(yyyy_mm_dd, "%Y-%m-%d").date()
        return d.strftime("%d/%m/%Y")
    except Exception:
        return yyyy_mm_dd

# =========================
#  BANXICO SIE
# =========================
@st.cache_data(ttl=60*30, show_spinner=False)
def sie_opportuno(series_id, token: str):
    url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series_id}/datos/oportuno"
    r = http_session().get(url, headers={"Bmx-Token": token})
    r.raise_for_status()
    return r.json()

@st.cache_data(ttl=60*30, show_spinner=False)
def sie_range(series_id: str, start_iso: str, end_iso: str, token: str):
    url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series_id}/datos/{start_iso}/{end_iso}"
    r = http_session(20).get(url, headers={"Bmx-Token": token})
    r.raise_for_status()
    j = r.json()
    series = j.get("bmx", {}).get("series", [])
    if not series:
        return []
    return series[0].get("datos", []) or []

def sie_last_n(series_id: str, n: int, token: str):
    """Devuelve últimos n como lista de (fecha 'YYYY-MM-DD', valor float)."""
    end = today_cdmx()
    start = end - timedelta(days=2*365)
    obs = sie_range(series_id, start.isoformat(), end.isoformat(), token)
    vals = []
    for o in obs:
        f = o.get("fecha"); v = try_float(o.get("dato"))
        if f and (v is not None):
            vals.append((f, v))
    if not vals:
        return []
    vals.sort(key=lambda x: x[0])
    return vals[-n:]

# Mapa de series SIE
SIE = {
    "USD_FIX":   "SF43718",
    "EUR_MXN":   "SF46410",
    "JPY_MXN":   "SF46406",
    "UDIS":      "SP68257",
    "CETES_28":  "SF43936",
    "CETES_91":  "SF43939",
    "CETES_182": "SF43942",
    "CETES_364": "SF43945",
}

# =========================
#  INEGI (UMA)
# =========================
@st.cache_data(ttl=60*60, show_spinner=False)
def get_uma(inegi_token: str):
    """
    UMA nacional: 620706 (diaria), 620707 (mensual), 620708 (anual)
    Retorna: {'fecha','diaria','mensual','anual','_status','_source'}
    """
    base = "https://www.inegi.org.mx/app/api/indicadores/desarrolladores/jsonxml/INDICATOR"
    ids = "620706,620707,620708"
    urls = [
        f"{base}/{ids}/es/00/true/BISE/2.0/{inegi_token}?type=json",
        f"{base}/{ids}/es/00/true/BIE/2.0/{inegi_token}?type=json",  # fallback
    ]
    last_err = None
    def _num(x):
        try: return float(str(x).replace(",", ""))
        except: return None
    for u in urls:
        try:
            r = http_session(20).get(u)
            if r.status_code != 200:
                last_err = f"HTTP {r.status_code}"; continue
            data = r.json()
            series = data.get("Series") or data.get("series") or []
            if not series:
                last_err = "Sin 'Series'"; continue

            def last_obs(s):
                obs = s.get("OBSERVATIONS") or s.get("observations") or []
                return obs[-1] if obs else None

            d_obs = last_obs(series[0]); m_obs = last_obs(series[1]) if len(series)>1 else None
            a_obs = last_obs(series[2]) if len(series)>2 else None

            def get_v(o): return _num(o.get("OBS_VALUE") or o.get("value")) if o else None
            def get_f(o):
                if not o: return None
                return o.get("TIME_PERIOD") or o.get("periodo") or o.get("time_period") or o.get("fecha")

            return {
                "fecha":   get_f(d_obs) or get_f(m_obs) or get_f(a_obs),
                "diaria":  get_v(d_obs),
                "mensual": get_v(m_obs),
                "anual":   get_v(a_obs),
                "_status": "ok",
                "_source": "INEGI",
            }
        except Exception as e:
            last_err = str(e); continue

    return {"fecha": None, "diaria": None, "mensual": None, "anual": None,
            "_status": f"err: {last_err}", "_source": "fallback"}

# =========================
#  Noticias financieras (RSS)
# =========================
def build_news_bullets(max_items=12):
    feeds = [
        "https://www.reuters.com/markets/americas/mexico/feed/?rpc=401&",
        "https://www.eleconomista.com.mx/rss/economia",
        "https://www.elfinanciero.com.mx/rss/finanzas/",
        "https://www.bloomberglinea.com/mexico/rss/",
        "https://www.banxico.org.mx/rss/prensa.xml",
    ]
    keywords = [
        "México","Banxico","inflación","tasa","TIIE","CETES","dólar","tipo de cambio",
        "Pemex","FOMC","Fed","nearshoring","déficit","rating","Moody","Fitch","S&P"
    ]
    rows = []
    for url in feeds:
        try:
            fp = feedparser.parse(url)
            for e in fp.entries[:40]:
                title = (e.get("title","") or "").strip()
                summary = (e.get("summary","") or "")
                link = (e.get("link","") or "").strip()
                txt = f"{title} {summary}".lower()
                if any(k.lower() in txt for k in keywords):
                    rows.append((e.get("published",""), title, link))
        except Exception:
            pass
    try: rows.sort(reverse=True, key=lambda x: x[0])
    except Exception: pass
    bullets = [f"• {t} — {l}" for _, t, l in rows[:max_items]]
    return "\n".join(bullets) if bullets else "Sin novedades (verifica conexión y RSS)."

# =========================
#  Sidebar (logo + tokens + estado simple)
# =========================
st.sidebar.header("IMEMSA")
st.sidebar.caption(f"Última verificación: {now_ts()}")

with st.sidebar.expander("🔑 Tokens de APIs", expanded=False):
    st.caption("Si ingresas un token aquí, la app lo usará en lugar del definido en el código.")
    token_banxico_input = st.text_input("BANXICO_TOKEN", value="", type="password")
    token_inegi_input   = st.text_input("INEGI_TOKEN",   value="", type="password")
    if token_banxico_input.strip():
        BANXICO_TOKEN = token_banxico_input.strip()
    if token_inegi_input.strip():
        INEGI_TOKEN = token_inegi_input.strip()

# =========================
#  Opciones
# =========================
with st.expander("Opciones"):
    movex_window = st.number_input("Ventana MOVEX (promedio móvil sobre FIX, días)", min_value=5, max_value=60, value=20, step=1)
    add_charts   = st.toggle("Agregar hoja 'Gráficos'", value=True)
    add_raw      = st.toggle("Agregar hoja 'Datos crudos' (últimos 12)", value=True)

# =========================
#  UI principal
# =========================
st.title("📈 Indicadores (últimos 6 días) + Noticias")
st.caption("Genera un Excel con B2..G2 como fechas, filas de indicadores y (opcional) hojas de gráficos y datos crudos.")

# Vista previa rápida de noticias
with st.expander("📰 Noticias (previa)"):
    st.markdown(build_news_bullets(max_items=8).replace("•","-"))

if st.button("Generar Excel"):
    # --- Validaciones mínimas
    if not BANXICO_TOKEN.strip():
        st.error("Falta BANXICO_TOKEN (ingrésalo en la barra lateral)."); st.stop()

    # --- Descargas de datos (SIE)
    try:
        last6_fix = sie_last_n(SIE["USD_FIX"], 6, BANXICO_TOKEN)
        last6_eur = sie_last_n(SIE["EUR_MXN"], 6, BANXICO_TOKEN)
        last6_jpy = sie_last_n(SIE["JPY_MXN"], 6, BANXICO_TOKEN)
        last6_c28 = sie_last_n(SIE["CETES_28"], 6, BANXICO_TOKEN)
        last6_c91 = sie_last_n(SIE["CETES_91"], 6, BANXICO_TOKEN)
        last6_c182= sie_last_n(SIE["CETES_182"],6, BANXICO_TOKEN)
        last6_c364= sie_last_n(SIE["CETES_364"],6, BANXICO_TOKEN)
        last6_udis= sie_last_n(SIE["UDIS"],     6, BANXICO_TOKEN)
    except Exception as e:
        st.error(f"Error consultando SIE: {e}")
        st.stop()

    # --- MOVEX sobre FIX
    # Construimos serie larga para calcular PM
    end = today_cdmx()
    start = end - timedelta(days=2*365)
    obs_fix_long = sie_range(SIE["USD_FIX"], start.isoformat(), end.isoformat(), BANXICO_TOKEN)
    series_fix = [try_float(o.get("dato")) for o in obs_fix_long if try_float(o.get("dato")) is not None]
    movex = []
    for i in range(len(series_fix)):
        sub = series_fix[max(0,i-movex_window+1): i+1]
        movex.append(sum(sub)/len(sub) if sub else None)
    last6_movex = movex[-6:] if movex else [None]*6

    # --- UMA (INEGI)
    uma = get_uma(INEGI_TOKEN)

    # --- Construimos tabla alineada por fechas de FIX
    dates_fix = [d for d,_ in last6_fix]  # 'YYYY-MM-DD'
    date_labels = [format_ddmmyyyy(d) for d in dates_fix]

    def as_map(lst): return {d:v for d,v in lst}
    m_eur, m_jpy = as_map(last6_eur), as_map(last6_jpy)
    m_c28, m_c91, m_c182, m_c364 = as_map(last6_c28), as_map(last6_c91), as_map(last6_c182), as_map(last6_c364)
    m_udis = as_map(last6_udis)
    m_fix  = as_map(last6_fix)

    rowset = [
        ("USD/MXN (FIX)",       [m_fix.get(d)  for d in dates_fix]),
        (f"MOVEX {movex_window}", last6_movex if len(last6_movex)==6 else [None]*6),
        ("EUR/MXN",             [m_eur.get(d)  for d in dates_fix]),
        ("JPY/MXN",             [m_jpy.get(d)  for d in dates_fix]),
        ("CETES 28d (%)",       [m_c28.get(d)  for d in dates_fix]),
        ("CETES 91d (%)",       [m_c91.get(d)  for d in dates_fix]),
        ("CETES 182d (%)",      [m_c182.get(d) for d in dates_fix]),
        ("CETES 364d (%)",      [m_c364.get(d) for d in dates_fix]),
        ("UDIS",                [m_udis.get(d) for d in dates_fix]),
    ]

    # --- Armado del Excel en memoria (XlsxWriter)
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})

    # Formatos
    fmt_title   = wb.add_format({'bold': True, 'font_size': 14})
    fmt_hdr     = wb.add_format({'bold': True, 'align': 'center', 'bg_color': '#F2F2F2', 'border':1})
    fmt_row_lbl = wb.add_format({'bold': True, 'align': 'left'})
    fmt_num4    = wb.add_format({'num_format': '0.0000'})
    fmt_num6    = wb.add_format({'num_format': '0.000000'})
    fmt_pct2    = wb.add_format({'num_format': '0.00%'})
    fmt_date    = wb.add_format({'align':'center'})
    fmt_wrap    = wb.add_format({'text_wrap': True})

    # ---------- Hoja Indicadores ----------
    ws = wb.add_worksheet("Indicadores")
    ws.write(0, 0, "Indicadores (últimos 6 días)", fmt_title)
    ws.write(1, 0, "Fecha", fmt_hdr)
    # Fechas en B2..G2 (fila 1 indexado 0 => 1)
    for i, d in enumerate(date_labels):
        ws.write(1, 1+i, d, fmt_hdr)

    start_row = 2  # fila 3 1-based
    for r, (label, values) in enumerate(rowset):
        ws.write(start_row+r, 0, label, fmt_row_lbl)
        for c, val in enumerate(values):
            # formato según magnitud (CETES en % vs tipo de cambio)
            if "CETES" in label:
                ws.write(start_row+r, 1+c, val/100 if (val is not None) else None, fmt_pct2)
            elif label == "UDIS":
                ws.write(start_row+r, 1+c, val, fmt_num6)
            else:
                ws.write(start_row+r, 1+c, val, fmt_num4)

    # UMA (bloque aparte)
    ws.write(1, 9, "UMA", fmt_hdr)
    ws.write(2, 9, "Fecha", fmt_row_lbl);  ws.write(2,10, uma.get("fecha"))
    ws.write(3, 9, "Diaria", fmt_row_lbl); ws.write(3,10, uma.get("diaria"))
    ws.write(4, 9, "Mensual", fmt_row_lbl);ws.write(4,10, uma.get("mensual"))
    ws.write(5, 9, "Anual", fmt_row_lbl);  ws.write(5,10, uma.get("anual"))
    ws.set_column(0, 0, 20)   # Col A
    ws.set_column(1, 6, 14)   # B..G
    ws.set_column(9,10, 14)   # UMA cols

    # ---------- Hoja Noticias ----------
    ws2 = wb.add_worksheet("Noticias")
    ws2.write(0, 0, "Noticias financieras recientes", fmt_title)
    news_text = build_news_bullets(12)
    ws2.write(1, 0, news_text, fmt_wrap)
    ws2.set_column(0, 0, 120)

    # ---------- Hoja Datos crudos (opcional) ----------
    if add_raw:
        ws3 = wb.add_worksheet("Datos crudos")
        ws3.write(0,0,"Serie", fmt_hdr); ws3.write(0,1,"Fecha", fmt_hdr); ws3.write(0,2,"Valor", fmt_hdr)
        row = 1
        def dump_raw(tag, lst):
            nonlocal row
            for f,v in lst:
                ws3.write(row, 0, tag)
                ws3.write(row, 1, format_ddmmyyyy(f), fmt_date)
                ws3.write(row, 2, v, fmt_num6)
                row += 1
        dump_raw("USD/MXN (FIX)", last6_fix)
        dump_raw("EUR/MXN", last6_eur)
        dump_raw("JPY/MXN", last6_jpy)
        dump_raw("CETES 28d (%)", last6_c28)
        dump_raw("CETES 91d (%)", last6_c91)
        dump_raw("CETES 182d (%)", last6_c182)
        dump_raw("CETES 364d (%)", last6_c364)
        dump_raw("UDIS", last6_udis)
        ws3.set_column(0, 0, 18); ws3.set_column(1, 1, 12); ws3.set_column(2, 2, 14)

    # ---------- Hoja Gráficos (opcional) ----------
    if add_charts:
        ws4 = wb.add_worksheet("Gráficos")
        # Gráfico 1: USD/MXN FIX
        chart1 = wb.add_chart({'type': 'line'})
        # rango: hoja Indicadores, fila de USD/MXN = start_row + 0
        usd_row = start_row + 0
        chart1.add_series({
            'name':       "=Indicadores!$A$%d" % (usd_row+1),
            'categories': "=Indicadores!$B$2:$G$2",
            'values':     "=Indicadores!$B$%d:$G$%d" % (usd_row+1, usd_row+1),
        })
        chart1.set_title({'name': 'USD/MXN (FIX)'})
        chart1.set_x_axis({'name': 'Fecha'})
        chart1.set_y_axis({'name': 'Tipo de cambio'})
        ws4.insert_chart('B2', chart1, {'x_scale': 1.3, 'y_scale': 1.2})

        # Gráfico 2: CETES (4 líneas)
        chart2 = wb.add_chart({'type': 'line'})
        labels = [("CETES 28d (%)", 4), ("CETES 91d (%)", 5), ("CETES 182d (%)", 6), ("CETES 364d (%)", 7)]
        for lbl, offset in labels:
            r = start_row + offset
            chart2.add_series({
                'name':       "=Indicadores!$A$%d" % (r+1),
                'categories': "=Indicadores!$B$2:$G$2",
                'values':     "=Indicadores!$B$%d:$G$%d" % (r+1, r+1),
            })
        chart2.set_title({'name': 'CETES (%)'})
        chart2.set_x_axis({'name': 'Fecha'})
        chart2.set_y_axis({'num_format': '0.00%'})
        ws4.insert_chart('B18', chart2, {'x_scale': 1.3, 'y_scale': 1.2})

    # Cerrar y servir
    wb.close()
    output.seek(0)
    st.success("¡Listo! Archivo generado con datos y hojas completas.")
    st.download_button(
        "⬇️ Descargar Excel",
        data=output.getvalue(),
        file_name=f"indicadores_{today_cdmx()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
