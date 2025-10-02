
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
# Config
# ==========================================================
st.set_page_config(page_title="Indicadores IMEMSA", layout="wide")

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
# Helpers robustos (no rompen si faltan dependencias)
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
    # Usa la implementación del usuario si existe
    if _has("get_uma"):
        try:
            return get_uma()
        except Exception:
            try:
                return getattr(get_uma, "__wrapped__", get_uma)()
            except Exception:
                pass
    # Fallback: valores None para no romper formato
    return {"diario": None, "mensual": None, "anual": None}

def _safe_rolling_movex(window=None):
    if _has("rolling_movex_for_last6"):
        try:
            return rolling_movex_for_last6(window=window) if window else rolling_movex_for_last6()
        except Exception:
            return None
    return None

def _sie_range(series_id: str, start: str, end: str):
    """Consulta robusta de una serie SIE: usa sie_range del usuario si existe;
    si no, intenta Banxico SIE API con token en st.secrets['banxico_token'] o env BANXICO_TOKEN.
    Devuelve lista de objetos {'fecha': 'YYYY-MM-DD', 'dato': <float/str>} o [].
    """
    # 1) Ruta nativa del usuario
    if _has("sie_range"):
        try:
            return sie_range(series_id, start, end)
        except Exception:
            pass

    # 2) Fallback vía Banxico (requiere token)
    token = None
    try:
        token = st.secrets.get("banxico_token")
    except Exception:
        token = None
    if not token:
        token = os.environ.get("BANXICO_TOKEN")

    if not token:
        # Sin token: devolvemos vacío
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
        obs = series[0].get("datos", [])
        return obs
    except Exception:
        return []

# ==========================================================
# Series mapping (usar el del usuario si existe)
# ==========================================================
DEFAULT_SIE_SERIES = {
    # Divisas (MXN por unidad de divisa)
    "USD_FIX": "SF43718",     # Tipo de cambio FIX (referencial)
    "EUR_MXN": "SE17927",     # EUR a MXN (aprox - puede variar por catálogo)
    "JPY_MXN": "SE163190",    # JPY a MXN (aprox - puede variar)
    "UDIS":    "SP68257",     # UDIS
    # CETES (tasas % anual)
    "CETES_28":  "SF60653",
    "CETES_91":  "SF60654",
    "CETES_182": "SF60655",
    "CETES_364": "SF60656",
    # TIIE (%)
    "TIIE_28":   "SF60634",
    "TIIE_91":   "SF60639",
    "TIIE_182":  "SF60640",
    # Tasa objetivo (%)
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

TEMPLATE_PATH = str((Path(__file__).parent / "Indicadores_template.xlsx").resolve())

# ==========================================================
# Cálculos de series para la plantilla (5 días hábiles C..G)
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
        vals = []
        last = None
        for d in header_dates:
            if d in values_map:
                last = values_map[d]
            vals.append(last)
        return vals

    fix_vals  = _ffill(m_fix)
    eur_vals  = _ffill(m_eur)
    jpy_vals  = _ffill(m_jpy)
    udis_vals = _ffill(m_udis)

    # CETES: usamos "asof" (último dato disponible hasta esa fecha)
    cet_start = (days[0] - timedelta(days=450)).isoformat()

    def _asof(series_key: str):
        m = _as_map(series_key, cet_start, end)
        vals = []
        last = None
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
# Exportador (plantilla nueva)
# ==========================================================
def export_indicadores_template_bytes():
    days = last_5_business_days()  # [C..G]
    S = _series_maps_for_template(days)

    # MONEX compra/venta desde función del usuario (si existe)
    mv = _safe_rolling_movex(globals().get("movex_win"))
    try:
        mpct = float(globals().get("margen_pct", 0.20))  # por ciento
    except Exception:
        mpct = 0.20

    # Construir payload para write_layout_v3
    payload = {
        "DOLAR": {
            # FIX (últimos 2 días: F=ayer, G=hoy)
            "5": {"F": (S["fix"][-2] if len(S["fix"]) >= 2 else None),
                  "G": (S["fix"][-1] if S["fix"] else None)},
        },
        "YEN": {
            "10": {"F": (S["jpy"][-2] if len(S["jpy"]) >= 2 else None),
                   "G": (S["jpy"][-1] if S["jpy"] else None)},
            # tipo cruzado USD/JPY => MXN/JPY? o YEN por USD.
            # Mantendremos relación para celda 11:G si escritor lo espera.
            "11": {"G": ((S["fix"][-1] / S["jpy"][-1]) if (S["fix"] and S["jpy"] and S["fix"][-1] and S["jpy"][-1]) else None)},
        },
        "EURO": {
            "14": {"F": (S["eur"][-2] if len(S["eur"]) >= 2 else None),
                   "G": (S["eur"][-1] if S["eur"] else None)},
            # EURUSD (cruzado)
            "15": {"G": (round((S["eur"][-1] / S["fix"][-1]), 4) if (S["eur"] and S["fix"] and S["eur"][-1] and S["fix"][-1]) else None)},
        },
        "UDIS": {
            "18": {"F": (round(S["udis"][-2], 4) if len(S["udis"]) >= 2 and S["udis"][-2] is not None else None),
                   "G": (round(S["udis"][-1], 4) if S["udis"] and S["udis"][-1] is not None else None)},
        },
        "TIIE": {
            # dividir entre 100 para porcentaje correcto
            "21": {"G": ((S["tobj"][-1] / 100.0) if (S["tobj"] and S["tobj"][-1] is not None) else None)},
            "22": {"G": ((S["t28"] [-1] / 100.0) if (S["t28"]  and S["t28"] [-1] is not None) else None)},
            "23": {"G": ((S["t91"] [-1] / 100.0) if (S["t91"]  and S["t91"] [-1] is not None) else None)},
            "24": {"G": ((S["t182"][-1] / 100.0) if (S["t182"] and S["t182"][-1] is not None) else None)},
        },
        "CETES": {
            "27": {"G": ((S["c28"] [-1] / 100.0) if (S["c28"]  and S["c28"] [-1] is not None) else None)},
            "28": {"G": ((S["c91"] [-1] / 100.0) if (S["c91"]  and S["c91"] [-1] is not None) else None)},
            "29": {"G": ((S["c182"][-1] / 100.0) if (S["c182"] and S["c182"][-1] is not None) else None)},
            "30": {"G": ((S["c364"][-1] / 100.0) if (S["c364"] and S["c364"][-1] is not None) else None)},
        },
        "UMA": {
            "33": {"G": S["uma"].get("diario")},
            "34": {"G": S["uma"].get("mensual")},
            "35": {"G": S["uma"].get("anual")},
        },
    }

    # MONEX compra/venta con margen (si tenemos serie)
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
    with open(tmp.name, "rb") as f:
        content = f.read()
    try:
        os.unlink(tmp.name)
    except Exception:
        pass
    return content

# ==========================================================
# UI
# ==========================================================
st.title("Indicadores (últimos 5 días hábiles) + Descarga")

with st.expander("Fechas (C..G) — G = hoy", expanded=False):
    st.write([d.strftime("%Y-%m-%d") for d in last_5_business_days()])

if "xlsx_bytes" not in st.session_state:
    st.session_state["xlsx_bytes"] = None

col1, col2 = st.columns([1,1])
with col1:
    if st.button("Generar Excel", key="btn_generar"):
        with st.spinner("Generando desde la PLANTILLA nueva..."):
            st.session_state["xlsx_bytes"] = export_indicadores_template_bytes()
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

