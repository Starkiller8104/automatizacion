
import os
import io
import json
import math
import time
import tempfile
from datetime import datetime, timedelta
from pathlib import Path

import streamlit as st

# ==========================================================
# Config / Branding
# ==========================================================
st.set_page_config(page_title="Indicadores IMEMSA", layout="wide")

LOGO_PATH = str((Path(__file__).parent / "logo.png").resolve())
TEMPLATE_PATH = str((Path(__file__).parent / "Indicadores_template.xlsx").resolve())

# ==========================================================
# Password gate (simple)
# - Usa st.secrets['app_password'] o env APP_PASSWORD
# ==========================================================
APP_PASSWORD = None
try:
    APP_PASSWORD = st.secrets.get("app_password")
except Exception:
    APP_PASSWORD = None
if not APP_PASSWORD:
    APP_PASSWORD = os.environ.get("APP_PASSWORD")

def _password_ok(p: str) -> bool:
    if not APP_PASSWORD:
        return True  # si no hay password configurada, no se bloquea
    try:
        return str(p) == str(APP_PASSWORD)
    except Exception:
        return False

if "auth_ok" not in st.session_state:
    st.session_state["auth_ok"] = False

with st.container():
    cols = st.columns([1, 4])
    with cols[0]:
        try:
            st.image(LOGO_PATH, caption=None, use_container_width=True)
        except Exception:
            pass
    with cols[1]:
        st.markdown("# Indicadores (últimos 5 días hábiles) + Descarga")

    # Barra divisoria
    st.markdown("---")

# Login box
if not st.session_state["auth_ok"]:
    with st.form("login_form", clear_on_submit=False):
        st.subheader("Acceso")
        pwd = st.text_input("Password", type="password")
        submit = st.form_submit_button("Entrar")
        if submit:
            if _password_ok(pwd):
                st.session_state["auth_ok"] = True
                st.rerun()
            else:
                st.error("Password incorrecto")
    st.stop()

# ==========================================================
# Utilidades de fecha / zona horaria
# ==========================================================
def today_cdmx():
    try:
        import pytz
        tz = pytz.timezone("America/Mexico_City")
        return datetime.now(tz).replace(tzinfo=None)
    except Exception:
        return datetime.now()

def last_5_business_days(end=None):
    end = (end or today_cdmx()).date()
    days = []
    d = end
    while len(days) < 5:
        if d.weekday() < 5:
            days.append(d)
        d -= timedelta(days=1)
    return list(reversed(days))  # [C,D,E,F,G], G = hoy

# ==========================================================
# Helpers robustos
# ==========================================================
def _has(name: str) -> bool:
    return name in globals()

def _try_float(x):
    try:
        if x is None or (isinstance(x, str) and x.strip() == ""):
            return None
        return float(str(x).replace(",", ""))
    except Exception:
        return None

def _parse_any_date(s):
    try:
        from dateutil import parser as _p
        return _p.parse(s)
    except Exception:
        try:
            return datetime.fromisoformat(s)
        except Exception:
            return None

def _safe_get_uma():
    if _has("get_uma"):
        try:
            return get_uma()
        except Exception:
            try:
                return getattr(get_uma, "__wrapped__", get_uma)()
            except Exception:
                pass
    return {"diario": None, "mensual": None, "anual": None}

def _safe_rolling_movex(window=None):
    if _has("rolling_movex_for_last6"):
        try:
            return rolling_movex_for_last6(window=window) if window else rolling_movex_for_last6()
        except Exception:
            return None
    return None

def _sie_range(series_id: str, start: str, end: str):
    if _has("sie_range"):
        try:
            return sie_range(series_id, start, end)
        except Exception:
            pass
    token = None
    try:
        token = st.secrets.get("banxico_token")
    except Exception:
        token = None
    if not token:
        token = os.environ.get("BANXICO_TOKEN")

    if not token:
        return []

    import requests
    url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series_id}/datos/{start}/{end}?token={token}"
    try:
        r = requests.get(url, timeout=20)
        r.raise_for_status()
        data = r.json()
        series = data.get("bmx", {}).get("series", [])
        if not series:
            return []
        return series[0].get("datos", []) or []
    except Exception:
        return []

DEFAULT_SIE_SERIES = {
    "USD_FIX": "SF43718",
    "EUR_MXN": "SE17927",
    "JPY_MXN": "SE163190",
    "UDIS":    "SP68257",
    "CETES_28":  "SF60653",
    "CETES_91":  "SF60654",
    "CETES_182": "SF60655",
    "CETES_364": "SF60656",
    "TIIE_28":   "SF60634",
    "TIIE_91":   "SF60639",
    "TIIE_182":  "SF60640",
    "OBJETIVO": "SF61745",
}
def SIE(id_key: str) -> str:
    if _has("SIE_SERIES"):
        try:
            return SIE_SERIES[id_key]
        except Exception:
            pass
    return DEFAULT_SIE_SERIES[id_key]

# ==========================================================
# Writer (plantilla)
# ==========================================================
from app_writer_layout_v3 import write_layout_v3

# ==========================================================
# Series para la plantilla (5 días hábiles C..G)
# ==========================================================
def _series_maps_for_template(days):
    header_dates = [d.isoformat() for d in days]

    def _as_map(series_key: str, start: str, end: str):
        obs = _sie_range(SIE(series_key), start, end)
        m = {}
        for o in obs or []:
            _f = _parse_any_date(o.get("fecha"))
            _v = _try_float(o.get("dato"))
            if _f and (_v is not None):
                m[_f.date().isoformat()] = _v
        return m

    start = days[0].isoformat(); end = days[-1].isoformat()

    m_fix  = _as_map("USD_FIX", start, end)
    m_eur  = _as_map("EUR_MXN", start, end)
    m_jpy  = _as_map("JPY_MXN", start, end)
    m_udis = _as_map("UDIS",    start, end)

    def _ffill(values_map):
        vals, last = [], None
        for d in header_dates:
            if d in values_map:
                last = values_map[d]
            vals.append(last)
        return vals

    fix_vals  = _ffill(m_fix)
    eur_vals  = _ffill(m_eur)
    jpy_vals  = _ffill(m_jpy)
    udis_vals = _ffill(m_udis)

    cet_start = (days[0] - timedelta(days=450)).isoformat()
    def _asof(series_key: str):
        m = _as_map(series_key, cet_start, end)
        vals, last = [], None
        for d in header_dates:
            if d in m:
                last = m[d]
            vals.append(last)
        return vals

    c28  = _asof("CETES_28")
    c91  = _asof("CETES_91")
    c182 = _asof("CETES_182")
    c364 = _asof("CETES_364")

    t28  = _asof("TIIE_28")
    t91  = _asof("TIIE_91")
    t182 = _asof("TIIE_182")
    tobj = _asof("OBJETIVO")

    uma = _safe_get_uma()
    uma_dict = {"diario": uma.get("diario"), "mensual": uma.get("mensual"), "anual": uma.get("anual")}

    return {
        "dates": days,
        "fix": fix_vals,
        "eur": eur_vals,
        "jpy": jpy_vals,
        "udis": udis_vals,
        "c28": c28,
        "c91": c91,
        "c182": c182,
        "c364": c364,
        "t28": t28,
        "t91": t91,
        "t182": t182,
        "tobj": tobj,
        "uma": uma_dict,
    }

# ==========================================================
# Exportador (plantilla + opcionales)
# ==========================================================
def export_indicadores_template_bytes(add_fred=False, add_rss=False, add_graficos=False, add_raw=False):
    days = last_5_business_days()  # [C..G]
    S = _series_maps_for_template(days)

    mv = _safe_rolling_movex(globals().get("movex_win"))
    try:
        mpct = float(globals().get("margen_pct", 0.20))
    except Exception:
        mpct = 0.20

    payload = {
        "DOLAR": {"5": {"F": (S["fix"][-2] if len(S["fix"]) >= 2 else None),
                        "G": (S["fix"][-1] if S["fix"] else None)}},
        "YEN":   {"10": {"F": (S["jpy"][-2] if len(S["jpy"]) >= 2 else None),
                         "G": (S["jpy"][-1] if S["jpy"] else None)},
                  "11": {"G": ((S["fix"][-1] / S["jpy"][-1]) if (S["fix"] and S["jpy"] and S["fix"][-1] and S["jpy"][-1]) else None)}},
        "EURO":  {"14": {"F": (S["eur"][-2] if len(S["eur"]) >= 2 else None),
                         "G": (S["eur"][-1] if S["eur"] else None)},
                  "15": {"G": (round((S["eur"][-1] / S["fix"][-1]), 4) if (S["eur"] and S["fix"] and S["eur"][-1] and S["fix"][-1]) else None)}},
        "UDIS":  {"18": {"F": (round(S["udis"][-2], 4) if len(S["udis"]) >= 2 and S["udis"][-2] is not None else None),
                         "G": (round(S["udis"][-1], 4) if S["udis"] and S["udis"][-1] is not None else None)}},
        "TIIE":  {"21": {"G": ((S["tobj"][-1] / 100.0) if (S["tobj"] and S["tobj"][-1] is not None) else None)},
                  "22": {"G": ((S["t28"] [-1] / 100.0) if (S["t28"]  and S["t28"] [-1] is not None) else None)},
                  "23": {"G": ((S["t91"] [-1] / 100.0) if (S["t91"]  and S["t91"] [-1] is not None) else None)},
                  "24": {"G": ((S["t182"][-1] / 100.0) if (S["t182"] and S["t182"][-1] is not None) else None)}},
        "CETES": {"27": {"G": ((S["c28"] [-1] / 100.0) if (S["c28"]  and S["c28"] [-1] is not None) else None)},
                  "28": {"G": ((S["c91"] [-1] / 100.0) if (S["c91"]  and S["c91"] [-1] is not None) else None)},
                  "29": {"G": ((S["c182"][-1] / 100.0) if (S["c182"] and S["c182"][-1] is not None) else None)},
                  "30": {"G": ((S["c364"][-1] / 100.0) if (S["c364"] and S["c364"][-1] is not None) else None)}},
        "UMA":   {"33": {"G": S["uma"].get("diario")},
                  "34": {"G": S["uma"].get("mensual")},
                  "35": {"G": S["uma"].get("anual")}},
    }

    if mv and isinstance(mv, (list, tuple)):
        try:
            compra = [(x * (1 - mpct/100.0) if x is not None else None) for x in mv]
            venta  = [(x * (1 + mpct/100.0) if x is not None else None) for x in mv]
            c_y = compra[-2] if len(compra) >= 2 else None
            c_h = compra[-1] if len(compra) >= 1 else None
            v_y = venta [-2] if len(venta ) >= 2 else None
            v_h = venta [-1] if len(venta ) >= 1 else None
            if c_y is not None: payload.setdefault("DOLAR", {}).setdefault("6", {})["F"] = c_y
            if c_h is not None: payload.setdefault("DOLAR", {}).setdefault("6", {})["G"] = c_h
            if v_y is not None: payload.setdefault("DOLAR", {}).setdefault("7", {})["F"] = v_y
            if v_h is not None: payload.setdefault("DOLAR", {}).setdefault("7", {})["G"] = v_h
        except Exception:
            pass

    # Escribir archivo desde plantilla
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx"); tmp.close()
    header_dates = last_5_business_days()
    write_layout_v3(TEMPLATE_PATH, tmp.name, header_dates=header_dates, payload=payload)

    # Agregar hojas opcionales (placeholder) si el usuario las marcó
    if any([add_fred, add_rss, add_graficos, add_raw]):
        try:
            from openpyxl import load_workbook
            wb = load_workbook(tmp.name)
            def ensure_ws(name, note):
                if name in wb.sheetnames:
                    ws = wb[name]
                else:
                    ws = wb.create_sheet(title=name)
                ws["A1"] = note
                return ws
            if add_fred:
                ensure_ws("FRED", "Hoja FRED - pendiente poblar")
            if add_rss:
                ensure_ws("Noticias_RSS", "Hoja RSS - pendiente poblar")
            if add_graficos:
                ensure_ws("Gráficos", "Hoja Gráficos - pendiente poblar")
            if add_raw:
                ensure_ws("Datos crudos", "Hoja Datos crudos - pendiente poblar")
            wb.save(tmp.name)
        except Exception:
            pass

    with open(tmp.name, "rb") as f:
        content = f.read()
    try:
        os.unlink(tmp.name)
    except Exception:
        pass
    return content

# ==========================================================
# UI (diseño original con opciones)
# ==========================================================
with st.expander("Selecciona las Hojas del Excel que contendrá tu archivo"):
    st.caption("Activa/desactiva hojas opcionales del archivo Excel")
    add_fred = st.checkbox("Agregar hoja FRED", value=False)
    add_rss = st.checkbox("Agregar hoja Noticias_RSS", value=False)
    add_graficos = st.checkbox("Agregar hoja 'Gráficos'", value=False)
    add_raw = st.checkbox("Agregar hoja 'Datos crudos'", value=False)

if "xlsx_bytes" not in st.session_state:
    st.session_state["xlsx_bytes"] = None

col1, col2 = st.columns([1,1])
with col1:
    if st.button("Generar Excel", key="btn_generar"):
        with st.spinner("Generando desde la PLANTILLA nueva..."):
            st.session_state["xlsx_bytes"] = export_indicadores_template_bytes(
                add_fred=add_fred, add_rss=add_rss, add_graficos=add_graficos, add_raw=add_raw
            )
        st.success("Listo. ¡Descarga abajo!")

st.markdown("---")
st.subheader("Descarga")
st.download_button(
    "Descargar Excel",
    data=(st.session_state["xlsx_bytes"] or b""),
    file_name="Indicadores " + today_cdmx().strftime("%Y-%m-%d %H%M%S") + ".xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    disabled=(st.session_state["xlsx_bytes"] is None),
    key="btn_descargar",
)

st.caption("Plantilla: " + TEMPLATE_PATH)

