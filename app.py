
import os
import inspect
import tempfile
from datetime import datetime, timedelta, date
from pathlib import Path

import streamlit as st

st.set_page_config(page_title="Indicadores IMEMSA", layout="wide")
LOGO_PATH = str((Path(__file__).parent / "logo.png").resolve())
TEMPLATE_PATH = str((Path(__file__).parent / "Indicadores_template_2col.xlsx").resolve())

# ---------------- Password ----------------
APP_PASSWORD = None
try:
    APP_PASSWORD = st.secrets.get("app_password")
except Exception:
    APP_PASSWORD = None
if not APP_PASSWORD:
    APP_PASSWORD = os.environ.get("APP_PASSWORD")

def _password_ok(p: str) -> bool:
    if not APP_PASSWORD:
        return True
    return str(p) == str(APP_PASSWORD)

if "auth_ok" not in st.session_state:
    st.session_state["auth_ok"] = False

cols = st.columns([1,4])
with cols[0]:
    try:
        st.image(LOGO_PATH, use_container_width=True)
    except Exception:
        pass
with cols[1]:
    st.markdown("# Indicadores (día actual y día anterior)")
st.markdown("---")

if not st.session_state["auth_ok"]:
    with st.form("login_form"):
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

# ---------------- Fechas ----------------
def today_cdmx():
    try:
        import pytz
        tz = pytz.timezone("America/Mexico_City")
        return datetime.now(tz).replace(tzinfo=None)
    except Exception:
        return datetime.now()

def business_days_back(n=10, end_date=None):
    end = (end_date or today_cdmx().date())
    days = []
    d = end
    while len(days) < n:
        if d.weekday() < 5:
            days.append(d)
        d -= timedelta(days=1)
    return days

# ---------------- Tokens/Secrets ----------------
def _get_secret(name: str):
    v = None
    try:
        v = st.secrets.get(name)
    except Exception:
        v = None
    if not v:
        v = os.environ.get(name) or os.environ.get(name.upper())
    return v

BANXICO_TOKEN = globals().get("BANXICO_TOKEN") or _get_secret("banxico_token") or _get_secret("BANXICO_TOKEN")
INEGI_TOKEN   = globals().get("INEGI_TOKEN")   or _get_secret("inegi_token")   or _get_secret("INEGI_TOKEN")
FRED_API_KEY  = globals().get("FRED_API_KEY")  or _get_secret("fred_api_key")  or _get_secret("FRED_API_KEY")

# ---------------- Series SIE ----------------
SIE_SERIES = {
    "USD_FIX":   "SF43718",
    "EUR_MXN":   "SF46410",
    "JPY_MXN":   "SF46406",
    "UDIS":      "SP68257",
    "TIIE_28":   "SF60648",
    "TIIE_91":   "SF60649",
    "TIIE_182":  "SF60650",
    "CETES_28":  "SF43936",
    "CETES_91":  "SF43939",
    "CETES_182": "SF43942",
    "CETES_364": "SF43945",
    "OBJETIVO":  "SF61745",
}
def SIE(key: str) -> str:
    return SIE_SERIES[key]

# ---------------- Helpers fetch ----------------
def _has(name: str) -> bool:
    return name in globals()

def _parse_any_date(s):
    try:
        from dateutil import parser as _p
        return _p.parse(s)
    except Exception:
        try:
            return datetime.fromisoformat(s)
        except Exception:
            return None

def _try_float(x):
    try:
        if x is None or (isinstance(x, str) and x.strip() == ""):
            return None
        return float(str(x).replace(",", ""))
    except Exception:
        return None

def _sie_range(series_id: str, start: str, end: str):
    if _has("sie_range"):
        try:
            return sie_range(series_id, start, end)
        except Exception:
            pass
    token = BANXICO_TOKEN
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

def _safe_get_uma():
    # try user's get_uma
    if _has("get_uma"):
        try:
            sig = inspect.signature(get_uma)
            if len(sig.parameters) >= 1:
                return get_uma(INEGI_TOKEN)
            else:
                return get_uma()
        except Exception:
            try:
                base = getattr(get_uma, "__wrapped__", get_uma)
                sig = inspect.signature(base)
                if len(sig.parameters) >= 1:
                    return base(INEGI_TOKEN)
                else:
                    return base()
            except Exception:
                pass
    # fallback: try INEGI indicator id from env if provided
    UMA_INDICATOR_ID = os.environ.get("INEGI_UMA_ID", "")
    if INEGI_TOKEN and UMA_INDICATOR_ID:
        import requests
        url = f"https://www.inegi.org.mx/app/api/indicadores/desarrolladores/jsonxml/INDICATOR/{UMA_INDICATOR_ID}/00000/es/{INEGI_TOKEN}?type=json"
        try:
            r = requests.get(url, timeout=20); r.raise_for_status()
            j = r.json()
            series = j.get("Series", [])
            if series:
                val = _try_float(series[0]["OBSERVATIONS"][-1]["OBS_VALUE"])
                return {"diario": val, "mensual": val, "anual": val}
        except Exception:
            pass
    # last fallback from secrets if provided
    try:
        diario = float(_get_secret("uma_diario") or "nan")
        mensual = float(_get_secret("uma_mensual") or "nan")
        anual = float(_get_secret("uma_anual") or "nan")
        return {"diario": diario, "mensual": mensual, "anual": anual}
    except Exception:
        return {"diario": None, "mensual": None, "anual": None}

def _safe_rolling_movex(window=None):
    if _has("rolling_movex_for_last6"):
        try:
            return rolling_movex_for_last6(window=window) if window else rolling_movex_for_last6()
        except Exception:
            return None
    return None

def _fred_inflation_yoy_for(months):
    """Return dict {YYYY-MM: yoy_decimal} using FRED CPIAUCSL (YoY). Requires FRED_API_KEY."""
    if not FRED_API_KEY:
        return {}
    import requests
    # get monthly CPI values last 36 months
    url = f"https://api.stlouisfed.org/fred/series/observations?series_id=CPIAUCSL&api_key={FRED_API_KEY}&file_type=json&frequency=m&observation_start=2015-01-01"
    try:
        r = requests.get(url, timeout=20); r.raise_for_status()
        data = r.json()
        obs = data.get("observations", [])
        # map date->value
        vals = {}
        for o in obs:
            try:
                v = float(o["value"])
            except Exception:
                continue
            vals[o["date"][:7]] = v
        yoy = {}
        for ym in months:
            # year-month like '2025-09'
            y, m = ym.split("-")
            y_prev = f"{int(y)-1}-{m}"
            if ym in vals and y_prev in vals:
                yoy[ym] = (vals[ym] / vals[y_prev]) - 1.0
        return yoy
    except Exception:
        return {}

# ---------------- Fechas clave por FIX ----------------
def _latest_and_previous_value_dates():
    end = today_cdmx().date()
    lookback = business_days_back(25, end)
    start = lookback[-1].isoformat()
    obs = _sie_range(SIE("USD_FIX"), start, end.isoformat())
    have = []
    for o in obs:
        d = _parse_any_date(o.get("fecha"))
        v = _try_float(o.get("dato"))
        if d and (v is not None):
            dd = d.date()
            if dd <= end:
                have.append(dd)
    have = sorted(set(have))
    if not have:
        latest = end
        prev = next(d for d in business_days_back(10, end) if d < end)
        return (prev, latest)
    latest = have[-1]
    prevs = [d for d in have if d < latest]
    prev = (prevs[-1] if prevs else next(d for d in business_days_back(10, latest) if d < latest))
    return (prev, latest)

# ---------------- Series as-of ----------------
def _series_values_for_dates(d_prev: date, d_latest: date):
    start = (d_prev - timedelta(days=450)).isoformat()
    end = d_latest.isoformat()

    def _as_map(series_key):
        obs = _sie_range(SIE(series_key), start, end)
        m = {}
        for o in obs or []:
            d = _parse_any_date(o.get("fecha")); v = _try_float(o.get("dato"))
            if d and (v is not None):
                m[d.date().isoformat()] = v
        return m

    def _asof(m, d):
        keys = sorted(k for k in m.keys() if k <= d.isoformat())
        return (m[keys[-1]] if keys else None)

    m_fix  = _as_map("USD_FIX")
    m_jpy  = _as_map("JPY_MXN")
    m_eur  = _as_map("EUR_MXN")
    m_udis = _as_map("UDIS")
    m_c28  = _as_map("CETES_28")
    m_c91  = _as_map("CETES_91")
    m_c182 = _as_map("CETES_182")
    m_c364 = _as_map("CETES_364")
    m_t28  = _as_map("TIIE_28")
    m_t91  = _as_map("TIIE_91")
    m_t182 = _as_map("TIIE_182")
    m_tobj = _as_map("OBJETIVO")

    uma = _safe_get_uma()

    def _two(m, scale=1.0, rnd=None):
        v_prev   = _asof(m, d_prev)
        v_latest = _asof(m, d_latest)
        if v_prev   is not None:   v_prev   = (v_prev   / scale)
        if v_latest is not None:   v_latest = (v_latest / scale)
        if rnd is not None:
            v_prev   = (round(v_prev, rnd)   if v_prev   is not None else None)
            v_latest = (round(v_latest, rnd) if v_latest is not None else None)
        return v_prev, v_latest

    fix_prev, fix_latest     = _two(m_fix)
    jpy_prev, jpy_latest     = _two(m_jpy)
    eur_prev, eur_latest     = _two(m_eur)
    udis_prev, udis_latest   = _two(m_udis, rnd=4)
    c28_prev, c28_latest     = _two(m_c28,  scale=100.0)
    c91_prev, c91_latest     = _two(m_c91,  scale=100.0)
    c182_prev, c182_latest   = _two(m_c182, scale=100.0)
    c364_prev, c364_latest   = _two(m_c364, scale=100.0)
    t28_prev, t28_latest     = _two(m_t28,  scale=100.0)
    t91_prev, t91_latest     = _two(m_t91,  scale=100.0)
    t182_prev, t182_latest   = _two(m_t182, scale=100.0)
    tobj_prev, tobj_latest   = _two(m_tobj, scale=100.0)

    # Cross USD/JPY (JPY per USD) = FIX / (MXN per JPY)
    usdjpy_prev   = (fix_prev / jpy_prev)   if (fix_prev and jpy_prev)   else None
    usdjpy_latest = (fix_latest/ jpy_latest)if (fix_latest and jpy_latest) else None

    # MONEX (si existe) con margen
    mv = _safe_rolling_movex(globals().get("movex_win"))
    try:
        mpct = float(globals().get("margen_pct", 0.20))
    except Exception:
        mpct = 0.20
    compra_prev = compra_latest = venta_prev = venta_latest = None
    if mv and isinstance(mv, (list, tuple)):
        try:
            compra = [(x * (1 - mpct/100.0) if x is not None else None) for x in mv]
            venta  = [(x * (1 + mpct/100.0) if x is not None else None) for x in mv]
            if len(compra) >= 2:
                compra_prev, compra_latest = compra[-2], compra[-1]
            if len(venta) >= 2:
                venta_prev,  venta_latest  = venta[-2],  venta[-1]
        except Exception:
            pass

    return {
        "fix": (fix_prev, fix_latest),
        "jpy": (jpy_prev, jpy_latest),
        "eur": (eur_prev, eur_latest),
        "udis": (udis_prev, udis_latest),
        "c28": (c28_prev, c28_latest),
        "c91": (c91_prev, c91_latest),
        "c182": (c182_prev, c182_latest),
        "c364": (c364_prev, c364_latest),
        "t28": (t28_prev, t28_latest),
        "t91": (t91_prev, t91_latest),
        "t182": (t182_prev, t182_latest),
        "tobj": (tobj_prev, tobj_latest),
        "uma": uma,
        "usdjpy": (usdjpy_prev, usdjpy_latest),
        "monex_compra": (compra_prev, compra_latest),
        "monex_venta":  (venta_prev,  venta_latest),
    }

# ---------------- Writer ----------------
def write_two_col_template(template_path: str, out_path: str, d_prev: date, d_latest: date, values: dict):
    from openpyxl import load_workbook
    wb = load_workbook(template_path)
    ws = wb.active

    # Encabezados con formato dd/mm/aaaa
    ws["C2"].value = d_prev
    ws["D2"].value = today_cdmx().date()
    ws["C2"].number_format = "dd/mm/yyyy"
    ws["D2"].number_format = "dd/mm/yyyy"

    rows = {
        "fix":   5,
        "monex_compra": 6,
        "monex_venta":  7,
        "jpy":   10,
        # fila 11: USD/JPY (JPY por USD)
        "eur":   14,
        # fila 15: EURUSD = EUR/FIX (4 dec)
        "udis":  18,
        "tobj":  21,
        "t28":   22,
        "t91":   23,
        "t182":  24,
        "c28":   27,
        "c91":   28,
        "c182":  29,
        "c364":  30,
    }

    def write_pair(key, round_to=None):
        r = rows[key]
        v_prev, v_latest = values.get(key, (None, None))
        if round_to is not None:
            v_prev   = (round(v_prev, round_to)   if v_prev   is not None else None)
            v_latest = (round(v_latest, round_to) if v_latest is not None else None)
        ws[f"C{r}"] = v_prev
        ws[f"D{r}"] = v_latest

    write_pair("fix")
    write_pair("monex_compra")
    write_pair("monex_venta")
    write_pair("jpy")

    # USD/JPY en fila 11
    v_prev, v_latest = values.get("usdjpy", (None, None))
    ws["C11"] = v_prev
    ws["D11"] = v_latest

    write_pair("eur")

    # EURUSD (fila 15) 4 dec
    try:
        eur_prev, eur_latest = values.get("eur", (None, None))
        fix_prev, fix_latest = values.get("fix", (None, None))
        eurusd_prev   = (eur_prev  / fix_prev)   if (eur_prev  and fix_prev)  else None
        eurusd_latest = (eur_latest/ fix_latest) if (eur_latest and fix_latest) else None
    except Exception:
        eurusd_prev = eurusd_latest = None
    ws["C15"] = (round(eurusd_prev, 4) if eurusd_prev is not None else None)
    ws["D15"] = (round(eurusd_latest, 4) if eurusd_latest is not None else None)

    write_pair("udis", round_to=4)
    write_pair("tobj")
    write_pair("t28")
    write_pair("t91")
    write_pair("t182")
    write_pair("c28")
    write_pair("c91")
    write_pair("c182")
    write_pair("c364")

    # UMA SOLO EN B33-B35
    uma = values.get("uma", {})
    ws["B33"] = uma.get("diario")
    ws["B34"] = uma.get("mensual")
    ws["B35"] = uma.get("anual")

    # Inflación EUA (B41=Octubre, B42=Septiembre) — % YoY si hay FRED_API_KEY
    try:
        y = today_cdmx().year
        ym_oct = f"{y}-10"
        ym_sep = f"{y}-09"
        fred = _fred_inflation_yoy_for([ym_oct, ym_sep])
        # escribir como decimal (p.ej 0.034 = 3.4%); la plantilla puede tener formato %
        ws["B41"] = fred.get(ym_oct)
        ws["B42"] = fred.get(ym_sep)
    except Exception:
        pass

    wb.save(out_path)

# ---------------- Exportador / UI ----------------
def export_indicadores_2col_bytes():
    d_prev, d_latest = _latest_and_previous_value_dates()
    vals = _series_values_for_dates(d_prev, d_latest)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx"); tmp.close()
    write_two_col_template(TEMPLATE_PATH, tmp.name, d_prev, d_latest, vals)
    with open(tmp.name, "rb") as f:
        content = f.read()
    try:
        os.unlink(tmp.name)
    except Exception:
        pass
    return content, d_prev, d_latest

with st.expander("Diagnóstico / Fechas y tokens", expanded=False):
    d_prev, d_latest = _latest_and_previous_value_dates()
    st.write({
        "dia_anterior (C2)": d_prev.strftime("%d/%m/%Y"),
        "dia_actual (D2)": today_cdmx().strftime("%d/%m/%Y"),
        "ultimo_con_valor (por FIX)": d_latest.strftime("%d/%m/%Y"),
        "banxico_token": bool(BANXICO_TOKEN),
        "inegi_token": bool(INEGI_TOKEN),
        "fred_api_key": bool(FRED_API_KEY),
    })

if "xlsx_bytes" not in st.session_state:
    st.session_state["xlsx_bytes"] = None

col1, col2 = st.columns([1,1])
with col1:
    if st.button("Generar Excel (2 columnas)"):
        with st.spinner("Generando desde plantilla 2 columnas..."):
            bytes_, d_prev, d_latest = export_indicadores_2col_bytes()
            st.session_state["xlsx_bytes"] = bytes_
        st.success("Listo. Descarga abajo.")

st.markdown("---")
st.subheader("Descarga")
st.download_button(
    "Descargar Excel",
    data=(st.session_state["xlsx_bytes"] or b""),
    file_name="Indicadores " + today_cdmx().strftime("%Y-%m-%d %H%M%S") + ".xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    disabled=(st.session_state["xlsx_bytes"] is None),
)
