
def _fetch_fix_direct(date_obj):
    """
    Obtiene el FIX (SF43718) exactamente para `date_obj` directo de Banxico SIE.
    Requiere BANXICO_TOKEN en variables de entorno.
    """
    import os, requests

# --- UMA shim (added by patch): never raises NameError if not provided elsewhere
def _uma_values():
    try:
        import streamlit as st
        v = st.session_state.get("uma_values")
        if isinstance(v, dict):
            return v
    except Exception:
        pass
    return {"diaria": None, "mensual": None, "anual": None}
    sid = "SF43718"
    token = os.environ.get("BANXICO_TOKEN") or os.environ.get("BANXICO_API_TOKEN")
    if not token:
        return None
    d = date_obj.strftime("%Y-%m-%d")
    url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{sid}/datos/{d}/{d}"
    headers = {"Bmx-Token": token}
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        j = resp.json()
        obs = j["bmx"]["series"][0]["datos"]
        if not obs:
            return None
        val = obs[-1]["dato"]
        val = float(str(val).replace(",", ""))
        return round(val, 4)
    except Exception:
        return None


import os
import re
import inspect
import tempfile
from datetime import datetime, timedelta, date
from pathlib import Path
from xml.etree import ElementTree as ET

import streamlit as st
# --- UI helpers import with fallback ---
try:
    from ui_helpers import inject_base_css, header, section_card, metric_row
except ModuleNotFoundError:
    def inject_base_css():
        st.markdown(
            """<style>
            .im-card{padding:1rem 1.25rem;border-radius:.9rem;background:rgba(255,255,255,.03);
                     border:1px solid rgba(255,255,255,.08);box-shadow:0 2px 8px rgba(0,0,0,.15);
                     margin-bottom:.75rem}
            .im-title{font-weight:700;font-size:1.1rem;margin-bottom:.25rem}
            .im-subtle{opacity:.8;font-size:.9rem}
            .stButton > button{border-radius:.75rem;padding:.6rem 1rem;font-weight:600}
            div[data-testid="metric-container"]{padding:.6rem .8rem;border-radius:.75rem;
                background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.08)}
            div[data-testid="stMetricDelta"] svg{display:none}
            </style>""", unsafe_allow_html=True
        )
    def header(logo_path:str, title:str, subtitle:str="", updated:str|None=None):
        cols = st.columns([1,5,3])
        with cols[0]:
            try: st.image(logo_path, use_container_width=True)
            except Exception: st.write("")
        with cols[1]:
            st.markdown(f"### {title}")
            if subtitle: st.markdown(f"<span class='im-subtle'>{subtitle}</span>", unsafe_allow_html=True)
        with cols[2]:
            if updated:
                st.markdown(
                    f"<div class='im-card'><div class='im-title'>Estado</div>"
                    f"<div class='im-subtle'>Actualizado: {updated}</div></div>",
                    unsafe_allow_html=True
                )
    def section_card(title:str, body_builder):
        st.markdown(f"<div class='im-card'><div class='im-title'>{title}</div>", unsafe_allow_html=True)
        body_builder()
        st.markdown("</div>", unsafe_allow_html=True)
    def metric_row(items):
        cols = st.columns(len(items))
        for i,(label,value,delta) in enumerate(items):
            with cols[i]:
                st.metric(label, value, delta if delta else None)


# ---------- Fine-grained progress helper ----------
class _Progress:
    def __init__(self, placeholder):
        self.placeholder = placeholder
        self.value = 0
        try:
            self._bar = placeholder.progress(0, text="")
            self._has_text = True
        except TypeError:
            self._bar = placeholder.progress(0)
            self._has_text = False

    def set(self, v, text=None):
        v = max(0, min(100, int(v)))
        self.value = v
        if self._has_text and text is not None:
            self._bar.progress(v, text=text)
        else:
            self._bar.progress(v)

    def inc(self, delta, text=None):
        self.set(self.value + delta, text=text)

# ======================
# Config / Branding
# ======================
st.set_page_config(page_title="Indicadores IMEMSA", layout="wide")

# --- Zona horaria CDMX y helpers de fecha/tiempo (cargados temprano) ---
from datetime import datetime, timedelta
try:
    from zoneinfo import ZoneInfo
    _MX_TZ = ZoneInfo("America/Mexico_City")
except Exception:
    _MX_TZ = None

def _cdmx_now():
    try:
        if _MX_TZ:
            return datetime.now(_MX_TZ)
    except Exception:
        pass
    return datetime.now()

# Polyfill: define today_cdmx si aún no existe cuando el módulo se evalúa
try:
    today_cdmx  # type: ignore
except NameError:
    def today_cdmx():
        return _cdmx_now()


# --- Hora/fecha CDMX segura para uso inmediato en nombre de archivo ---
from datetime import datetime
try:
    from zoneinfo import ZoneInfo
    _MX_TZ = ZoneInfo("America/Mexico_City")
except Exception:
    _MX_TZ = None

def _now_str_cdmx(fmt="%Y-%m-%d %H%M%S"):
    try:
        if _MX_TZ:
            return datetime.now(_MX_TZ).strftime(fmt)
    except Exception:
        pass
    return datetime.now().strftime(fmt)


def _get_secret_env(key, default=None):
    try:
        v = st.secrets.get(key)
    except Exception:
        v = None
    if v is None:
        v = os.environ.get(key) or os.environ.get(key.upper())
    return v if v is not None else default

# Optional clean UI (toggle with HIDE_DEFAULT_UI; default=on)
if str((_get_secret_env("HIDE_DEFAULT_UI", "1"))).strip().lower() in ("1","true","yes","on"):
    st.markdown(
        "<style>#MainMenu{visibility:hidden;} footer{visibility:hidden;} header{visibility:hidden;}</style>",
        unsafe_allow_html=True
    )

LOGO_PATH = str((Path(__file__).parent / "logo.png").resolve())
TEMPLATE_PATH = str((Path(__file__).parent / "Indicadores_template_2col.xlsx").resolve())

# ======================
# Password (opcional)
# ======================
APP_PASSWORD = _get_secret_env("app_password")
if not APP_PASSWORD:
    APP_PASSWORD = _get_secret_env("APP_PASSWORD")

def _password_ok(p: str) -> bool:
    if not APP_PASSWORD:
        return True
    return str(p) == str(APP_PASSWORD)

if "auth_ok" not in st.session_state:
    st.session_state["auth_ok"] = False


inject_base_css()
from datetime import datetime
try:
    import pytz
    updated_str = datetime.now(pytz.timezone("America/Mexico_City")).strftime("%Y-%m-%d %H:%M (CDMX)")
except Exception:
    updated_str = datetime.now().strftime("%Y-%m-%d %H:%M")
header(
    logo_path=LOGO_PATH,
    title="Indicadores de Tipo de Cambio",
    subtitle="IMEMSA · Reporte ejecutivo",
    updated=updated_str
)


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

# ======================
# Utilidades de fecha
# --- Encabezados efectivos (respeta corte 12:00 CDMX) ---
def _prev_business_day(base: date) -> date:
    d = base - timedelta(days=1)
    while d.weekday() >= 5:  # 5=Sat,6=Sun
        d -= timedelta(days=1)
    return d



def header_dates_effective():
    """
    Devuelve (d_prev, d_latest) usando *hoy* ajustado a día hábil:
    - d_latest: hoy si es hábil; si no, el hábil anterior
    - d_prev:   el hábil inmediatamente anterior a d_latest
    """
    now = _cdmx_now()
    d_latest = now.date()
    # Si es fin de semana, retrocede al hábil previo
    while d_latest.weekday() >= 5:
        d_latest -= timedelta(days=1)
    d_prev = _prev_business_day(d_latest)
    return (d_prev, d_latest)

def _prev_business_day(d):
    # Lunes → retrocede a viernes; fines de semana → retrocede a viernes
    if d.weekday() == 0:  # lunes
        return d - timedelta(days=3)
    while d.weekday() >= 5:  # sábado/domingo
        d -= timedelta(days=1)
    # Si ya es martes-viernes, regresa un día; si es lunes ya se manejó arriba
    if d.weekday() in (1,2,3,4):
        return d - timedelta(days=1)
    return d

def business_days_back(n, start_date):
    days = []
    d = start_date
    while len(days) < n:
        d = _prev_business_day(d)
        days.append(d)
    return days  # descendente




# ======================
# Tokens / Flags
# ======================
# Normalización de token Banxico: admite secrets['banxico_token'], secrets['BANXICO_TOKEN']
# y variables de entorno BANXICO_TOKEN o banxico_token.
def _get_env_any(keys, default=None):
    import os
    for k in keys:
        v = os.environ.get(k)
        if v:
            return v
    return default

BANXICO_TOKEN = (
    _get_secret_env("banxico_token") or
    _get_secret_env("BANXICO_TOKEN") or
    _get_env_any(["BANXICO_TOKEN","banxico_token"], "")
)
INEGI_TOKEN   = _get_secret_env("inegi_token",   "")
FRED_API_KEY  = _get_secret_env("fred_api_key",  "")


MONEX_FALLBACK = (_get_secret_env("MONEX_FALLBACK", "scrape") or "scrape").strip().lower()
def _fetch_monex_scrape():
    """
    Intenta extraer compra/venta USD y EUR desde https://www.monex.com.mx/.
    Devuelve {'usd': {'compra': float, 'venta': float}, 'eur': {...}} o None.
    """
    try:
        import re, requests
        try:
            from bs4 import BeautifulSoup
        except Exception:
            BeautifulSoup = None
        resp = requests.get("https://www.monex.com.mx/", headers={"User-Agent":"Mozilla/5.0"}, timeout=15)
        resp.raise_for_status()
        html = resp.text
        # Obtener texto
        if BeautifulSoup:
            soup = BeautifulSoup(html, "html.parser")
            txt = soup.get_text(" ", strip=True)
        else:
            txt = re.sub(r"<[^>]+>", " ", html)
            txt = re.sub(r"\s+", " ", txt)
        def parse_many(txt):
            out = {}
            pats = [
                r"(USD)\s*([0-9]+(?:[\.,][0-9]+)?)\s*/\s*([0-9]+(?:[\.,][0-9]+)?)",
                r"(EUR)\s*([0-9]+(?:[\.,][0-9]+)?)\s*/\s*([0-9]+(?:[\.,][0-9]+)?)",
                r"(USD|D[oó]lar)[^0-9]{0,40}Compra[^0-9]{0,12}([0-9]+(?:[\.,][0-9]+)?)\D+Venta[^0-9]{0,12}([0-9]+(?:[\.,][0-9]+)?)",
                r"(EUR|Euro)[^0-9]{0,40}Compra[^0-9]{0,12}([0-9]+(?:[\.,][0-9]+)?)\D+Venta[^0-9]{0,12}([0-9]+(?:[\.,][0-9]+)?)",
            ]
            for pat in pats:
                for m in re.finditer(pat, txt, flags=re.I):
                    cur = m.group(1).upper()
                    c = float(m.group(2).replace(",", "."))
                    v = float(m.group(3).replace(",", "."))
                    if "USD" in cur:
                        out["usd"] = {"compra": c, "venta": v}
                    elif "EUR" in cur:
                        out["eur"] = {"compra": c, "venta": v}
            return out
        out = parse_many(txt)
        return out or None
    except Exception:
        return None


def _get_margin_pct():
    try:
        v = _get_secret_env("MARGEN_PCT")
        if v is None:
            return 0.20
        return float(v)
    except Exception:
        return 0.20

# ======================
# Series SIE (base) + candidatos para TIIE 28/91/182
# ======================

# === Old stable helpers (restored) ===

def parse_any_date(s: str):
    """Devuelve datetime naive (sin tz)."""
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(s), fmt)
        except:
            pass
    return None

def try_float(x):
    try:
        return float(str(x).replace(",", "").strip())
    except:
        return None

def rolling_movex_for_last6(window:int=20):
    end = today_cdmx()
    start = end - timedelta(days=2*365)
    obs = sie_range(SIE_SERIES["USD_FIX"], start.isoformat(), end.isoformat())
    vals = []
    for o in obs:
        f = o.get("fecha"); v = try_float(o.get("dato"))
        if f and (v is not None):
            vals.append((f, v))
    if not vals:
        return []
    vals.sort(key=lambda x: parse_any_date(x[0]) or datetime.utcnow())
    series = [v for _, v in vals]
    out = []
    for i in range(len(series)):
        sub = series[max(0, i-window+1): i+1]
        out.append(sum(sub)/len(sub) if sub else None)
    return out

SIE_SERIES = {
    "USD_FIX":   "SF43718",
    "EUR_MXN":   "SF46410",
    "JPY_MXN":   "SF46406",
    "UDIS":      "SP68257",
    "TIIE_28": "SF60648",
    "TIIE_91": "SF60649",
    "TIIE_182": "SF60650",

    "CETES_28":  "SF43936",
    "CETES_91":  "SF43939",
    "CETES_182": "SF43942",
    "CETES_364": "SF43945",
}


SIE_SERIES = {
    "USD_FIX":   "SF43718",
    "EUR_MXN":   "SF46410",
    "JPY_MXN":   "SF46406",
    "UDIS":      "SP68257",
    "CETES_28":  "SF43936",
    "CETES_91":  "SF43939",
    "CETES_182": "SF43942",
    "CETES_364": "SF43945",
    "OBJETIVO":  "SF61745",
    "TIIE_28":   "SF60648",
    "TIIE_91":   "SF60649",
    "TIIE_182":  "SF60650",
}

def _parse_candidates_env(name: str, default_list):
    raw = _get_secret_env(name)
    if raw:
        return [s.strip() for s in str(raw).split(",") if s.strip()]
    return default_list

SIE_SERIES_CANDIDATES = {
    "TIIE_28":  _parse_candidates_env("SERIES_OVERRIDE__TIIE_28",  [SIE_SERIES["TIIE_28"]]),
    "TIIE_91":  _parse_candidates_env("SERIES_OVERRIDE__TIIE_91",  [SIE_SERIES["TIIE_91"]]),
    "TIIE_182": _parse_candidates_env("SERIES_OVERRIDE__TIIE_182", [SIE_SERIES["TIIE_182"]]),
}

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
        if x is None or (isinstance(x, str) and str(x).strip() == ""):
            return None
        return float(str(x).replace(",", ""))
    except Exception:
        return None

# ======================
# Banxico SIE helpers
# ======================
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
    url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series_id}/datos/{start}/{end}"
    try:
        r = requests.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0", "Bmx-Token": token})
        r.raise_for_status()
        data = r.json()
        series = data.get("bmx", {}).get("series", [])
        if not series:
            return []
        return series[0].get("datos", []) or []
    except Exception:
        return []


# --- Compatibility shim (old function name used by rolling_movex_for_last6) ---
def sie_range(series_id: str, start: str, end: str):
    """Backwards-compatible wrapper -> _sie_range"""
    return _sie_range(series_id, start, end)


def _sie_range_first_that_has_data(series_ids, start: str, end: str):
    for sid in series_ids:
        datos = _sie_range(sid, start, end)
        if datos:
            return sid, datos
    return None, []

def _to_map_from_obs(obs_list):
    m = {}
    for o in obs_list or []:
        d = _parse_any_date(o.get("fecha")); v = _try_float(o.get("dato"))
        if d and (v is not None):
            m[d.date().isoformat()] = v
    return m

# ======================
# UMA helpers
# ======================
def _safe_get_uma():
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
    try:
        import requests
        resp = requests.get("https://www.inegi.org.mx/temas/uma/", timeout=20, headers={"User-Agent": "Mozilla/5.0"})
        if resp.ok:
            txt = resp.text
            y = today_cdmx().year
            m = re.search(rf">{y}<.*?\$\s*([0-9]+\.?[0-9]*)\s*,\s*\$\s*([0-9,\.]+)\s*,\s*\$\s*([0-9,\.]+)", txt, flags=re.S)
            if m:
                diario  = _try_float(m.group(1))
                mensual = _try_float(m.group(2))
                anual   = _try_float(m.group(3))
                return {"diario": diario, "mensual": mensual, "anual": anual}
    except Exception:
        pass
    UMA_MAP = {
        2025: {"diario": 113.14, "mensual": 3439.46, "anual": 41273.52},
        2024: {"diario": 108.57, "mensual": 3300.53, "anual": 39606.36},
        2023: {"diario": 103.74, "mensual": 3153.70, "anual": 37844.40},
    }
    y = today_cdmx().year
    if y in UMA_MAP:
        return UMA_MAP[y]
    def _sf(name):
        v = _get_secret_env(name)
        if v is None or str(v).strip() == "":
            return None
        try:
            return float(str(v).replace(",", ""))
        except Exception:
            return None
    return {"diario": _sf("uma_diario"), "mensual": _sf("uma_mensual"), "anual": _sf("uma_anual")}

# ======================
# FRED helpers (inflación EUA)
# ======================
def _fred_inflation_yoy_map():
    if not FRED_API_KEY:
        return {}
    import requests
    url = f"https://api.stlouisfed.org/fred/series/observations?series_id=CPIAUCSL&api_key={FRED_API_KEY}&file_type=json&frequency=m&observation_start=2010-01-01"
    try:
        r = requests.get(url, timeout=20); r.raise_for_status()
        data = r.json()
        obs = data.get("observations", [])
        vals = {}
        for o in obs:
            try:
                v = float(o["value"])
            except Exception:
                continue
            vals[o["date"][:7]] = v
        yoy = {}
        for k, v in vals.items():
            y, m = k.split("-")
            prev = f"{int(y)-1}-{m}"
            if prev in vals:
                yoy[k] = (v / vals[prev]) - 1.0
        return yoy
    except Exception:
        return {}

def _fred_last_sept_oct_yoy():
    yoy = _fred_inflation_yoy_map()
    if not yoy:
        return (None, None)
    avail = sorted(yoy.keys())
    cutoff = today_cdmx().strftime("%Y-%m")
    sep_candidates = [k for k in avail if k.endswith("-09") and k <= cutoff]
    oct_candidates = [k for k in avail if k.endswith("-10") and k <= cutoff]
    sep = yoy.get(sep_candidates[-1]) if sep_candidates else None
    octo = yoy.get(oct_candidates[-1]) if oct_candidates else None
    return (sep, octo)

# ======================
# Fechas clave por FIX
# ======================
def _latest_and_previous_value_dates():
    # Fechas de encabezado efectivas (independientes de disponibilidad de FIX)
    return header_dates_effective()

    latest = have[-1]
    prevs = [d for d in have if d < latest]
    prev = (prevs[-1] if prevs else next(d for d in business_days_back(10, latest) if d < latest))
    return (prev, latest)

# ======================
# Series as-of (TIIE 182 fija)
# ======================

def _series_values_for_dates(d_prev: date, d_latest: date, prog: _Progress | None = None):
    start = (d_prev - timedelta(days=450)).isoformat()
    end = d_latest.isoformat()

    used_series = {}

    # Progreso visual
    step = 0.0
    if prog is not None:
        fetch_span = 60.0
        ops_total = 10.0  # FIX, JPY, EUR, UDIS, CETES*4, OBJETIVO
        step = fetch_span / ops_total

    def _as_map_fixed(series_id):
        obs = _sie_range(series_id, start, end)
        return _to_map_from_obs(obs)

    if prog: prog.inc(step, "Banxico SIE: USD FIX")
    m_fix  = _as_map_fixed(SIE_SERIES["USD_FIX"])
    if prog: prog.inc(step, "Banxico SIE: JPY/MXN")
    m_jpy  = _as_map_fixed(SIE_SERIES["JPY_MXN"])
    if prog: prog.inc(step, "Banxico SIE: EUR/MXN")
    m_eur  = _as_map_fixed(SIE_SERIES["EUR_MXN"])
    if prog: prog.inc(step, "Banxico SIE: UDIS")
    m_udis = _as_map_fixed(SIE_SERIES["UDIS"])

    if prog: prog.inc(step, "Banxico SIE: CETES 28/91/182/364")
    m_c28 = _as_map_fixed(SIE_SERIES["CETES_28"])
    m_c91 = _as_map_fixed(SIE_SERIES["CETES_91"])
    m_c182= _as_map_fixed(SIE_SERIES["CETES_182"])
    m_c364= _as_map_fixed(SIE_SERIES["CETES_364"])

    if prog: prog.inc(step, "Banxico SIE: Objetivo de tasa")
    m_tobj = _as_map_fixed(SIE_SERIES["OBJETIVO"]) if "OBJETIVO" in SIE_SERIES else {}

    def _asof(m, d):
        keys = sorted(k for k in m.keys() if k <= d.isoformat())
        return (m[keys[-1]] if keys else None)
    def _two(m, scale=1.0, rnd=None):
        v_prev   = _asof(m, d_prev)
        v_latest = _asof(m, d_latest)
        if v_prev   is not None:   v_prev   = v_prev/scale
        if v_latest is not None:   v_latest = v_latest/scale
        if rnd is not None:
            v_prev   = round(v_prev, rnd)   if v_prev   is not None else None
            v_latest = round(v_latest, rnd) if v_latest is not None else None
        return v_prev, v_latest

    fix_prev   = _fetch_fix_direct(d_prev) or _asof(m_fix, d_prev)
    fix_latest = _fetch_fix_direct(d_latest) or _asof(m_fix, d_latest)
    jpy_prev, jpy_latest     = _two(m_jpy)
    eur_prev, eur_latest     = _two(m_eur)
    udis_prev, udis_latest   = _two(m_udis, rnd=4)
    c28_prev, c28_latest     = _two(m_c28,  scale=100.0)
    c91_prev, c91_latest     = _two(m_c91,  scale=100.0)
    c182_prev, c182_latest   = _two(m_c182, scale=100.0)
    c364_prev, c364_latest   = _two(m_c364, scale=100.0)
    tobj_prev, tobj_latest = _two(m_tobj, scale=100.0) if m_tobj else (None, None)
    t28_prev, t28_latest = (None, None)
    t91_prev, t91_latest = (None, None)  # si no se usa en Excel actual

    # USD/JPY cruz
    usdjpy_prev   = (fix_prev / jpy_prev)     if (fix_prev is not None and jpy_prev is not None)      else None
    usdjpy_latest = (fix_latest / jpy_latest) if (fix_latest is not None and jpy_latest is not None)  else None

    # --- MONEX / Compra-Venta USD ---
    compra_prev = compra_latest = venta_prev = venta_latest = None
    mode = (_get_secret_env("MONEX_MODE", "old") or "old").strip().lower()

    if mode == "old":
        mv = rolling_movex_for_last6(window=globals().get("movex_win", 20))
        if mv and isinstance(mv, (list, tuple)):
            try:
                mpct = float(globals().get("margen_pct", _get_margin_pct()))
                compra = [(x * (1 - mpct/100.0) if x is not None else None) for x in mv]
                venta  = [(x * (1 + mpct/100.0) if x is not None else None) for x in mv]
                if len(compra) >= 2: compra_prev, compra_latest = compra[-2], compra[-1]
                if len(venta)  >= 2: venta_prev,  venta_latest  = venta[-2],  venta[-1]
            except Exception:
                pass
    elif mode == "scrape":
        monex = _fetch_monex_scrape() or {}
        usd = monex.get("usd") or {}
        compra_prev   = usd.get("compra"); compra_latest = usd.get("compra")
        venta_prev    = usd.get("venta");  venta_latest  = usd.get("venta")

    # Fallback final a FIX ± margen si falta algo
    mpct = _get_margin_pct()
    if fix_prev is not None and (compra_prev is None):   compra_prev   = fix_prev   * (1 - mpct/100.0)
    if fix_latest is not None and (compra_latest is None): compra_latest = fix_latest * (1 - mpct/100.0)
    if fix_prev is not None and (venta_prev is None):    venta_prev    = fix_prev    * (1 + mpct/100.0)
    if fix_latest is not None and (venta_latest is None):  venta_latest  = fix_latest  * (1 + mpct/100.0)

    st.session_state["used_series_ids"] = {
        "TIIE_28":  st.session_state.get("used_series_ids", {}).get("TIIE_28"),
        "TIIE_91":  st.session_state.get("used_series_ids", {}).get("TIIE_91"),
        "TIIE_182": used_series.get("TIIE_182") if "TIIE_182" in used_series else None,
    }

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
        "tobj": (tobj_prev, tobj_latest),
        "uma": (_uma_values() if "st" in globals() or "streamlit" in sys.modules else {"diaria": None, "mensual": None, "anual": None}),
        "usdjpy": (usdjpy_prev, usdjpy_latest),
        "monex_compra": (compra_prev, compra_latest),
        "monex_venta":  (venta_prev,  venta_latest),
    }

# ======================
# RSS helpers
# ======================
RSS_FEEDS = [
    ("Google News MX – Economía",
     "https://news.google.com/rss/headlines/section/topic/BUSINESS?hl=es-419&gl=MX&ceid=MX:es-419"),
    ("Google News MX – BMV",
     "https://news.google.com/rss/search?q=Bolsa%20Mexicana%20de%20Valores&hl=es-419&gl=MX&ceid=MX:es-419"),
    ("Google News MX – Banxico",
     "https://news.google.com/rss/search?q=Banxico&hl=es-419&gl=MX&ceid=MX:es-419"),
    ("Google News MX – Inflación INEGI",
     "https://news.google.com/rss/search?q=inflaci%C3%B3n%20M%C3%A9xico%20INEGI&hl=es-419&gl=MX&ceid=MX:es-419"),
]

def fetch_rss_items(url: str, max_items: int = 12):
    import requests
    try:
        r = requests.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
        r.raise_for_status()
        root = ET.fromstring(r.content)
        items = []
        for item in root.findall(".//item"):
            title = (item.findtext("title") or "").strip()
            link = (item.findtext("link") or "").strip()
            pubDate = (item.findtext("pubDate") or "").strip()
            source = ""
            s = item.find("source")
            if s is not None and (s.text or "").strip():
                source = s.text.strip()
            items.append({"title": title, "link": link, "pubDate": pubDate, "source": source})
            if len(items) >= max_items:
                break
        return items
    except Exception:
        return []

# ======================
# Writer 2 columnas + hoja RSS con estilo Arial 12 y sin grid
# ======================
def write_two_col_template(template_path: str, out_path: str, d_prev: date, d_latest: date, values: dict, prog: _Progress | None = None):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    wb = load_workbook(template_path)
    if prog:
        prog.set(65, "Escribiendo hoja principal…")
    ws = wb.active

    # Fechas
    ws["C2"].value = d_prev
    ws["D2"].value = d_latest
    ws["C2"].number_format = "dd/mm/yyyy"
    ws["D2"].number_format = "dd/mm/yyyy"

    rows = {
        "fix":   5,
        "monex_compra": 6,
        "monex_venta":  7,
        "jpy":   10,
        "eur":   14,
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
    ws["C11"], ws["D11"] = values.get("usdjpy", (None, None))
    write_pair("eur")
    # EURUSD 4 dec
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

    # UMA en B33-B35
    uma = values.get("uma", {})
    ws["B33"] = uma.get("diario")
    ws["B34"] = uma.get("mensual")
    ws["B35"] = uma.get("anual")

    # Inflación EUA
    def _as_pct_decimal(x):
        if x is None:
            return None
        try:
            x = float(x)
            if abs(x) > 1:
                x = x / 100.0
            return x
        except Exception:
            return None

    try:
        sep_yoy, oct_yoy = _fred_last_sept_oct_yoy()
        us_sep = _get_secret_env("us_inflation_sep")
        us_oct = _get_secret_env("us_inflation_oct")
        if sep_yoy is None and us_sep is not None:
            sep_yoy = _as_pct_decimal(us_sep)
        if oct_yoy is None and us_oct is not None:
            oct_yoy = _as_pct_decimal(us_oct)
        ws["B42"] = sep_yoy
        ws["B41"] = oct_yoy
        ws["B41"].number_format = "0.00%"
        ws["B42"].number_format = "0.00%"
    except Exception:
        pass

    if prog:
        prog.set(82, "Aplicando limpieza/formatos…")
    # Eliminar hojas no deseadas si existen
    for sheet_name in ["Lógica de datos", "Metadatos"]:
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]

    if prog:
        prog.set(86, "Generando hoja RSS…")
    # Hoja RSS con estilo Arial 12 y sin grid
    news_ws = wb.create_sheet("Noticias Financieras RSS")
    header = ["Fuente", "Título", "Fecha (CDMX)", "Link"]
    news_ws.append(header)
    header_font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    for col_idx in range(1, len(header)+1):
        c = news_ws.cell(row=1, column=col_idx)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(vertical="center")

    def _parse_pubdate(s):
        try:
            from email.utils import parsedate_to_datetime
            dt = parsedate_to_datetime(s)
            try:
                import pytz
                tz = pytz.timezone("America/Mexico_City")
                if dt.tzinfo:
                    dt = dt.astimezone(tz).replace(tzinfo=None)
            except Exception:
                pass
            return dt
        except Exception:
            return None

    for idx_feed, (fuente, url) in enumerate(RSS_FEEDS, start=1):
        if prog:
            prog.inc(12.0 / max(1, len(RSS_FEEDS)), f"RSS: {fuente}")
        items = fetch_rss_items(url, max_items=12)
        for it in items:
            title = it.get("title") or ""
            link  = it.get("link") or ""
            pub   = it.get("pubDate") or ""
            dt    = _parse_pubdate(pub)
            if dt is None:
                row = [f"{fuente} · {it.get('source')}" if it.get("source") else fuente,
                       title, pub, link]
            else:
                row = [f"{fuente} · {it.get('source')}" if it.get("source") else fuente,
                       title, dt, link]
            news_ws.append(row)

    # Arial 12 toda la hoja, sin gridlines
    news_ws.sheet_view.showGridLines = False
    max_r = news_ws.max_row
    max_c = news_ws.max_column
    for r in range(1, max_r+1):
        for c in range(1, max_c+1):
            cell = news_ws.cell(row=r, column=c)
            if r == 1:
                cell.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
                cell.alignment = Alignment(vertical="center")
            else:
                if cell.hyperlink:
                    cell.font = Font(name="Arial", size=12, color="0563C1")
                else:
                    cell.font = Font(name="Arial", size=12)
                if c == 2:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                else:
                    cell.alignment = Alignment(vertical="top")

    # Anchos
    widths = {1: 28, 2: 100, 3: 22, 4: 60}
    for idx, w in widths.items():
        news_ws.column_dimensions[chr(64+idx)].width = w

    # Fechas con formato
    for r in range(2, max_r+1):
        c3 = news_ws.cell(row=r, column=3)
        if isinstance(c3.value, datetime):
            c3.number_format = "dd/mm/yyyy HH:MM"
            c3.alignment = Alignment(vertical="top")
        else:
            c3.alignment = Alignment(vertical="top")
        c4 = news_ws.cell(row=r, column=4)
        if isinstance(c4.value, str) and c4.value.startswith("http"):
            c4.hyperlink = c4.value
            c4.style = "Hyperlink"

    news_ws.freeze_panes = "A2"

    
    # === Formatos de número y Rangos con Nombre ===
    try:
        # Formatos: FX 4 decimales, UDIS 6 decimales, Tasas como % con 2 decimales.
        fx_rows = [rows["fix"], rows["eur"], rows["jpy"]]
        udis_row = rows["udis"]
        tasa_rows = [rows["tobj"], rows["t28"], rows["t91"], rows["t182"], rows["c28"], rows["c91"], rows["c182"], rows["c364"]]

        for r in fx_rows:
            for col in ("C","D"):
                ws[f"{col}{r}"].number_format = "0.0000"

        for col in ("C","D"):
            ws[f"{col}{udis_row}"].number_format = "0.000000"

        for r in tasa_rows:
            for col in ("C","D"):
                ws[f"{col}{r}"].number_format = "0.00%"

        # Fechas
        ws["C2"].number_format = "dd/mm/yyyy"
        ws["D2"].number_format = "dd/mm/yyyy"

        # Rangos con nombre
        from openpyxl.workbook.defined_name import DefinedName
        def add_name(name, ref):
            try:
                existing = [dn for dn in wb.defined_names.definedName if dn.name == name]
                for dn in existing:
                    wb.defined_names.definedName.remove(dn)
            except Exception:
                pass
            wb.defined_names.append(DefinedName(name=name, attr_text=f"{ws.title}!{ref}"))

        add_name("RANGO_FECHAS", f"$C$2:$D$2")
        add_name("RANGO_USDMXN", f"$C${rows['fix']}:$D${rows['fix']}")
        add_name("RANGO_EURMXN", f"$C${rows['eur']}:$D${rows['eur']}")
        add_name("RANGO_JPYMXN", f"$C${rows['jpy']}:$D${rows['jpy']}")
        add_name("RANGO_UDIS",   f"$C${rows['udis']}:$D${rows['udis']}")
        add_name("RANGO_TOBJ",   f"$C${rows['tobj']}:$D${rows['tobj']}")
        add_name("RANGO_TIIE28", f"$C${rows['t28']}:$D${rows['t28']}")
        add_name("RANGO_TIIE91", f"$C${rows['t91']}:$D${rows['t91']}")
        add_name("RANGO_TIIE182",f"$C${rows['t182']}:$D${rows['t182']}")
        add_name("RANGO_C28",    f"$C${rows['c28']}:$D${rows['c28']}")
        add_name("RANGO_C91",    f"$C${rows['c91']}:$D${rows['c91']}")
        add_name("RANGO_C182",   f"$C${rows['c182']}:$D${rows['c182']}")
        add_name("RANGO_C364",   f"$C${rows['c364']}:$D${rows['c364']}")
    except Exception as _e:
        pass

    wb.save(out_path)
    if prog:
        prog.set(100, "Archivo listo.")

# ======================
# Exportador (no usado por el botón, lo dejo por compatibilidad)
# ======================
def export_indicadores_2col_bytes():
    d_prev, d_latest = _latest_and_previous_value_dates()
    vals = _series_values_for_dates(d_prev, d_latest)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.close()
    write_two_col_template(TEMPLATE_PATH, tmp.name, d_prev, d_latest, vals)
    with open(tmp.name, "rb") as f:
        content = f.read()
    try:
        os.unlink(tmp.name)
    except Exception:
        pass
    return content, d_prev, d_latest

# ======================
# Diagnóstico (oculto por defecto)
# ======================
SHOW_DIAGNOSTICS = str((_get_secret_env("SHOW_DIAGNOSTICS", "0"))).strip().lower() in ("1","true","yes","on")
if SHOW_DIAGNOSTICS:
    with st.expander("Diagnóstico / Fechas y tokens", expanded=False):
        d_prev, d_latest = _latest_and_previous_value_dates()
        used = st.session_state.get("used_series_ids", {})
        st.write({
            "dia_anterior (C2)": d_prev.strftime("%d/%m/%Y"),
            "dia_actual (D2)": today_cdmx().strftime("%d/%m/%Y"),
            "banxico_token_detectado": bool(BANXICO_TOKEN),
            "inegi_token_detectado": bool(INEGI_TOKEN),
            "fred_api_key_detectada": bool(FRED_API_KEY),
            "monex_fallback": MONEX_FALLBACK,
            "margen_pct": _get_margin_pct(),
            "TIIE_28_series_usada": used.get("TIIE_28"),
            "TIIE_91_series_usada": used.get("TIIE_91"),
            "TIIE_182_series_usada": used.get("TIIE_182"),
            "nota_TIIE182": "Se usa TIIE182 fija (secrets/env TIIE182_FIXED, default 7.9871%).",
        })

# ======================
# UI
# ======================
if "xlsx_bytes" not in st.session_state:
    st.session_state["xlsx_bytes"] = None

# Contenedor para la barra de progreso (aparece debajo del botón)
_progress_placeholder = st.empty()

if section_card('Generación de Excel', lambda: st.button("Generar Excel")):
    prog = _Progress(_progress_placeholder)
    prog.set(5, "Preparando fechas…")
    d_prev, d_latest = _latest_and_previous_value_dates()
    prog.set(10, "Preparando consultas…")
    vals = _series_values_for_dates(d_prev, d_latest, prog)
    prog.set(64, "Construyendo archivo…")
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.close()
    write_two_col_template(TEMPLATE_PATH, tmp.name, d_prev, d_latest, vals, prog)
    with open(tmp.name, "rb") as f:
        st.session_state["xlsx_bytes"] = f.read()
    try:
        os.unlink(tmp.name)
    except Exception:
        pass
    prog.set(100, "Listo!!!")
    st.success("Listo!!!")

st.download_button(
    "Descargar Excel",
    data=(st.session_state["xlsx_bytes"] or b""),
    file_name="Indicadores " + _now_str_cdmx("%Y-%m-%d %H%M%S") + ".xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    disabled=(st.session_state["xlsx_bytes"] is None),
)

# --- Safe SIE range (always available) ---
def _safe_sie_range(series_id: str, start: str, end: str):
    import requests
    import os
    token = os.environ.get("BANXICO_TOKEN","")
    url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series_id}/datos/{start}/{end}"
    headers = {"User-Agent":"Mozilla/5.0"}
    if token:
        headers["Bmx-Token"] = token
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    data = r.json()
    series = data.get("bmx",{}).get("series",[])
    return series[0].get("datos",[]) if series else []
