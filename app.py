
import os
import re
import inspect
import tempfile
from datetime import datetime, timedelta, date
from pathlib import Path
from xml.etree import ElementTree as ET

import streamlit as st

# ======================
# Config / Branding
# ======================
st.set_page_config(page_title="Indicadores IMEMSA", layout="wide")
LOGO_PATH = str((Path(__file__).parent / "logo.png").resolve())
TEMPLATE_PATH = str((Path(__file__).parent / "Indicadores_template_2col.xlsx").resolve())

# ======================
# Password (opcional)
# ======================
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

cols = st.columns([1, 4])
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

# ======================
# Utilidades de fecha
# ======================
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
    return days  # descendente

# ======================
# Secrets / Tokens / Flags
# ======================
def _get_secret(name: str, default=None):
    v = None
    try:
        v = st.secrets.get(name)
    except Exception:
        v = None
    if not v:
        v = os.environ.get(name) or os.environ.get(name.upper())
    return v if v is not None else default

# Defaults provided by Ariel (puedes sobreescribir en secrets/env)
BANXICO_TOKEN = _get_secret("banxico_token", "677aaedf11d11712aa2ccf73da4d77b6b785474eaeb2e092f6bad31b29de6609")
INEGI_TOKEN   = _get_secret("inegi_token",   "0146a9ed-b70f-4ea2-8781-744b900c19d1")
FRED_API_KEY  = _get_secret("fred_api_key",  "b4f11681f441da78103a3706d0dab1cf")

# MONEX fallback: por defecto activado para no dejar celdas vacías
MONEX_FALLBACK = (_get_secret("MONEX_FALLBACK", "fix") or "fix").strip().lower()
def _get_margin_pct():
    # default 0.20 (%). If you pass "0.3" => 0.3%; "1.5" => 1.5%
    try:
        v = _get_secret("MARGEN_PCT")
        if v is None: 
            return 0.20
        return float(v)
    except Exception:
        return 0.20

# ======================
# Series SIE (base) + candidatos para TIIE 28/91/182
# ======================
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
    # TIIE defaults (clásicas)
    "TIIE_28":   "SF60648",
    "TIIE_91":   "SF60649",
    "TIIE_182":  "SF60650",
}

def _parse_candidates_env(name: str, default_list):
    raw = _get_secret(name)
    if raw:
        return [s.strip() for s in str(raw).split(",") if s.strip()]
    return default_list

# Puedes agregar nuevos IDs en secrets (coma-separados) si Banxico cambia los códigos.
SIE_SERIES_CANDIDATES = {
    "TIIE_28":  _parse_candidates_env("SERIES_OVERRIDE__TIIE_28",  [SIE_SERIES["TIIE_28"]]),
    "TIIE_91":  _parse_candidates_env("SERIES_OVERRIDE__TIIE_91",  [SIE_SERIES["TIIE_91"]]),
    "TIIE_182": _parse_candidates_env("SERIES_OVERRIDE__TIIE_182", [SIE_SERIES["TIIE_182"]]),
}

def SIE(key: str) -> str:
    return SIE_SERIES[key]

# ======================
# Fetchers robustos
# ======================
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
        r = requests.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
        r.raise_for_status()
        data = r.json()
        series = data.get("bmx", {}).get("series", [])
        if not series:
            return []
        return series[0].get("datos", []) or []
    except Exception:
        return []

def _sie_range_first_that_has_data(series_ids, start: str, end: str):
    """Devuelve (series_id_usada, lista_de_datos) para la primera serie con datos en el rango; si ninguna trae, (None, [])."""
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

# ------- Fallback TIIE-182 vía HTML Banxico (tabla "TIIE 26 semanas - valores del banco") + variantes
def _tiie182_map_from_banxico_html():
    import requests
    headers = {"User-Agent": "Mozilla/5.0"}
    urls = [
        "https://www.banxico.org.mx/mercados/tiie-26-semanas-valores-banco.html",
        "https://www.banxico.org.mx/mercados/tiie-26-semanas-posturas-presentadas.html",
    ]
    for url in urls:
        try:
            r = requests.get(url, timeout=20, headers=headers)
            if not r.ok:
                continue
            html = r.text
            # Permitir coma o punto decimal; capturar 2 a 6 decimales; permitir espacios intermedios.
            rows = re.findall(r"(\\d{2}/\\d{2}/\\d{4}).{0,200}?([0-9]{1,2}[\\.,][0-9]{2,6})", html, flags=re.S)
            m = {}
            for (ddmmyyyy, rate) in rows:
                try:
                    rate = rate.replace(",", ".")
                    d = datetime.strptime(ddmmyyyy, "%d/%m/%Y").date().isoformat()
                    m[d] = float(rate)
                except Exception:
                    continue
            if m:
                return m
        except Exception:
            continue
    return {}

# ======================
# UMA helpers
# ======================
def _safe_get_uma():
    # 1) try user's get_uma
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
    # 2) try INEGI page scrape (simple regex)
    try:
        import requests
        resp = requests.get("https://www.inegi.org.mx/temas/uma/", timeout=20, headers={"User-Agent": "Mozilla/5.0"})
        if resp.ok:
            txt = resp.text
            y = today_cdmx().year
            m = re.search(rf">{y}<.*?\\$\\s*([0-9]+\\.?[0-9]*)\\s*,\\s*\\$\\s*([0-9,\\.]+)\\s*,\\s*\\$\\s*([0-9,\\.]+)", txt, flags=re.S)
            if m:
                diario  = _try_float(m.group(1))
                mensual = _try_float(m.group(2))
                anual   = _try_float(m.group(3))
                return {"diario": diario, "mensual": mensual, "anual": anual}
    except Exception:
        pass
    # 3) fixed map for recent years
    UMA_MAP = {
        2025: {"diario": 113.14, "mensual": 3439.46, "anual": 41273.52},
        2024: {"diario": 108.57, "mensual": 3300.53, "anual": 39606.36},
        2023: {"diario": 103.74, "mensual": 3153.70, "anual": 37844.40},
    }
    y = today_cdmx().year
    if y in UMA_MAP:
        return UMA_MAP[y]
    # 4) secrets manual
    def _sf(name):
        v = _get_secret(name)
        if v is None or str(v).strip() == "":
            return None
        try:
            return float(str(v).replace(",", ""))
        except Exception:
            return None
    diario  = _sf("uma_diario")
    mensual = _sf("uma_mensual")
    anual   = _sf("uma_anual")
    return {"diario": diario, "mensual": mensual, "anual": anual}

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
            vals[o["date"][:7]] = v  # YYYY-MM
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
    end = today_cdmx().date()
    lookback = business_days_back(25, end)
    start = lookback[-1].isoformat()
    obs = _sie_range(SIE_SERIES["USD_FIX"], start, end.isoformat())
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

# ======================
# Series as-of (con selección automática de TIIE y fallbacks mejorados)
# ======================
def _series_values_for_dates(d_prev: date, d_latest: date):
    # rango amplio para garantizar as-of
    start = (d_prev - timedelta(days=450)).isoformat()
    end = d_latest.isoformat()

    used_series = {}  # para diagnóstico: qué ID se usó

    # Series simples (sin candidatos)
    def _as_map_fixed(series_id):
        obs = _sie_range(series_id, start, end)
        return _to_map_from_obs(obs)

    m_fix  = _as_map_fixed(SIE_SERIES["USD_FIX"])
    m_jpy  = _as_map_fixed(SIE_SERIES["JPY_MXN"])
    m_eur  = _as_map_fixed(SIE_SERIES["EUR_MXN"])
    m_udis = _as_map_fixed(SIE_SERIES["UDIS"])

    m_c28  = _as_map_fixed(SIE_SERIES["CETES_28"])
    m_c91  = _as_map_fixed(SIE_SERIES["CETES_91"])
    m_c182 = _as_map_fixed(SIE_SERIES["CETES_182"])
    m_c364 = _as_map_fixed(SIE_SERIES["CETES_364"])

    m_tobj = _as_map_fixed(SIE_SERIES["OBJETIVO"])

    # TIIE con candidatos (probar en orden hasta encontrar datos)
    def _as_map_candidates(key_logic: str):
        sids = SIE_SERIES_CANDIDATES[key_logic]
        sid, obs = _sie_range_first_that_has_data(sids, start, end)
        if sid:
            used_series[key_logic] = sid
        else:
            used_series[key_logic] = None
        return _to_map_from_obs(obs)

    m_t28  = _as_map_candidates("TIIE_28")
    m_t91  = _as_map_candidates("TIIE_91")
    m_t182 = _as_map_candidates("TIIE_182")

    # Fallback para TIIE-182 si la serie SIE no trajo nada
    if not m_t182:
        html_map = _tiie182_map_from_banxico_html()
        if html_map:
            used_series["TIIE_182"] = "banxico_html_fallback"
            m_t182 = html_map

    # Fallback manual (secrets) si todo lo demás falla
    def _sf(name):
        v = _get_secret(name)
        if v is None or str(v).strip() == "":
            return None
        try:
            return float(str(v).replace(",", ""))
        except Exception:
            return None

    uma = _safe_get_uma()

    def _asof(m, d):
        keys = sorted(k for k in m.keys() if k <= d.isoformat())
        return (m[keys[-1]] if keys else None)

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

    # Manual secrets override as LAST resort for TIIE-182
    if t182_prev is None:
        t182_prev = _sf("TIIE182_prev")
    if t182_latest is None:
        t182_latest = _sf("TIIE182_latest")

    # USD/JPY (JPY por USD) = FIX / (MXN por JPY).
    usdjpy_prev   = (fix_prev / jpy_prev)     if (fix_prev is not None and jpy_prev is not None)      else None
    usdjpy_latest = (fix_latest / jpy_latest) if (fix_latest is not None and jpy_latest is not None)  else None

    # MONEX (si existe) con margen o fallback por FIX
    mv = None
    if "rolling_movex_for_last6" in globals():
        try:
            mv = rolling_movex_for_last6(window=globals().get("movex_win"))
        except Exception:
            mv = None

    compra_prev = compra_latest = venta_prev = venta_latest = None
    if mv and isinstance(mv, (list, tuple)):
        try:
            mpct = float(globals().get("margen_pct", _get_margin_pct()))
            compra = [(x * (1 - mpct/100.0) if x is not None else None) for x in mv]
            venta  = [(x * (1 + mpct/100.0) if x is not None else None) for x in mv]
            if len(compra) >= 2:
                compra_prev, compra_latest = compra[-2], compra[-1]
            if len(venta) >= 2:
                venta_prev,  venta_latest  = venta[-2],  venta[-1]
        except Exception:
            pass
    if MONEX_FALLBACK == "fix":
        mpct = _get_margin_pct()
        if fix_prev is not None and compra_prev is None:
            compra_prev = fix_prev * (1 - mpct/100.0)
        if fix_latest is not None and compra_latest is None:
            compra_latest = fix_latest * (1 - mpct/100.0)
        if fix_prev is not None and venta_prev is None:
            venta_prev = fix_prev * (1 + mpct/100.0)
        if fix_latest is not None and venta_latest is None:
            venta_latest = fix_latest * (1 + mpct/100.0)

    # Guardar IDs usados para diagnóstico
    used_series = st.session_state.get("used_series_ids", {})
    used_series["TIIE_28"]  = used_series.get("TIIE_28")
    used_series["TIIE_91"]  = used_series.get("TIIE_91")
    used_series["TIIE_182"] = used_series.get("TIIE_182")
    st.session_state["used_series_ids"] = used_series

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
        # namespace-safe: items are usually channel/item
        items = []
        for item in root.findall(".//item"):
            title = (item.findtext("title") or "").strip()
            link = (item.findtext("link") or "").strip()
            pubDate = (item.findtext("pubDate") or "").strip()
            source = ""
            # Google News incluye <source>
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
# Writer 2 columnas: respeta plantilla + formatos + limpia hojas + RSS (presentable)
# ======================
def write_two_col_template(template_path: str, out_path: str, d_prev: date, d_latest: date, values: dict):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.worksheet.dimensions import ColumnDimension
    wb = load_workbook(template_path)
    ws = wb.active  # hoja principal (donde están los indicadores)

    # Encabezados C2 (anterior) y D2 (hoy) con formato dd/mm/aaaa
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

    # Inflación EUA (B41=Octubre, B42=Septiembre) — % YoY
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
        if sep_yoy is None:
            sep_yoy = _as_pct_decimal(_get_secret("us_inflation_sep"))
        if oct_yoy is None:
            oct_yoy = _as_pct_decimal(_get_secret("us_inflation_oct"))
        ws["B42"] = sep_yoy
        ws["B41"] = oct_yoy
        ws["B41"].number_format = "0.00%"
        ws["B42"].number_format = "0.00%"
    except Exception:
        pass

    # --- Limpiar hojas que ya no quieres ---
    for sheet_name in ["Lógica de datos", "Metadatos"]:
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]

    # --- Hoja de Noticias Financieras RSS (con formato) ---
    news_ws = wb.create_sheet("Noticias Financieras RSS")

    # Header
    header = ["Fuente", "Título", "Fecha (CDMX)", "Link"]
    news_ws.append(header)
    from openpyxl.styles import Font, Alignment, PatternFill
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    for col_idx in range(1, len(header)+1):
        c = news_ws.cell(row=1, column=col_idx)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(vertical="center")

    # Content
    def _parse_pubdate(s):
        # Try common RFC822 -> datetime localized
        try:
            from email.utils import parsedate_to_datetime
            dt = parsedate_to_datetime(s)
            # Convert to CDMX if tz-aware
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

    for fuente, url in RSS_FEEDS:
        items = fetch_rss_items(url, max_items=12)
        for it in items:
            title = it.get("title") or ""
            link  = it.get("link") or ""
            pub   = it.get("pubDate") or ""
            dt    = _parse_pubdate(pub)
            if dt is None:
                # keep as text
                row = [f"{fuente} · {it.get('source')}" if it.get("source") else fuente,
                       title, pub, link]
            else:
                row = [f"{fuente} · {it.get('source')}" if it.get("source") else fuente,
                       title, dt, link]
            news_ws.append(row)

    # Column widths (heurístico)
    widths = {
        1: 28,   # Fuente
        2: 100,  # Título
        3: 22,   # Fecha
        4: 60,   # Link
    }
    for idx, w in widths.items():
        news_ws.column_dimensions[chr(64+idx)].width = w

    # Wrap en título, vertical top, hyperlinks, formato fecha
    last_row = news_ws.max_row
    for r in range(2, last_row+1):
        # título
        c2 = news_ws.cell(row=r, column=2)
        c2.alignment = Alignment(wrap_text=True, vertical="top")
        # fecha
        c3 = news_ws.cell(row=r, column=3)
        if isinstance(c3.value, datetime):
            c3.number_format = "dd/mm/yyyy HH:MM"
            c3.alignment = Alignment(vertical="top")
        else:
            c3.alignment = Alignment(vertical="top")
        # link
        c4 = news_ws.cell(row=r, column=4)
        if isinstance(c4.value, str) and c4.value.startswith("http"):
            c4.hyperlink = c4.value
            c4.style = "Hyperlink"
        # fuente
        c1 = news_ws.cell(row=r, column=1)
        c1.alignment = Alignment(vertical="top")

    # Freeze header
    news_ws.freeze_panes = "A2"

    wb.save(out_path)

# ======================
# Exportador
# ======================
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

# ======================
# UI
# ======================
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
        "nota_TIIE182": "Si sigue vacío, agrega SERIES_OVERRIDE__TIIE_182 o usa secrets TIIE182_prev / TIIE182_latest.",
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
