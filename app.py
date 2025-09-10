
import io
import re
import time
import html
import base64
import pytz
import re
import requests
import feedparser
import xlsxwriter
import streamlit as st
from datetime import datetime, timedelta, date
from email.utils import parsedate_to_datetime
from pathlib import Path
from PIL import Image
from urllib.parse import urlparse
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from requests.adapters import HTTPAdapter, Retry

def _fred_req_v1():
    s = requests.Session()
    try:
        from requests.adapters import HTTPAdapter, Retry
        rty = Retry(total=3, backoff_factor=0.5, status_forcelist=[429, 500, 502, 503, 504])
        s.mount("https://", HTTPAdapter(max_retries=rty))
    except Exception:
        pass
    return s

def _fred_fetch_v1(series_id: str, start: str, end: str, api_key: str):
    url = "https://api.stlouisfed.org/fred/series/observations"
    params = {"series_id": series_id, "api_key": api_key, "file_type": "json",
              "observation_start": start, "observation_end": end}
    r = _fred_req_v1().get(url, params=params, timeout=30)
    r.raise_for_status()
    out = []
    try:
        observations = r.json().get("observations", [])
    except Exception:
        observations = []
    for o in observations:
        v = o.get("value")
        if v not in (None, ".", ""):
            try:
                out.append((o["date"], float(v)))
            except Exception:
                pass
    return out

def _fred_write_v1(wb, series_dict, sheet_name="FRED_v2"):
    ws = wb.add_worksheet(sheet_name)
    fmt_bold = wb.add_format({'font_name': 'Arial', "bold": True, "align": "center"})
    fmt_date = wb.add_format({'font_name': 'Arial', "num_format": "yyyy-mm-dd"})
    fmt_num  = wb.add_format({'font_name': 'Arial', "num_format": "#,##0.0000"})
    headers = ["Fecha"] + list(series_dict.keys())
    ws.write_row(0, 0, headers, fmt_bold)
    fechas = sorted({d for vals in series_dict.values() for d, _ in vals})
    lookup = {name: {d: v for d, v in vals} for name, vals in series_dict.items()}
    for i, d in enumerate(fechas, start=1):
        try:
            ws.write_datetime(i, 0, datetime.fromisoformat(d), fmt_date)
        except Exception:
            ws.write_string(i, 0, d)
        for j, name in enumerate(series_dict.keys(), start=1):
            v = lookup[name].get(d)
            if v is not None:
                ws.write_number(i, j, v, fmt_num)
    last_row = 1 + len(fechas)
    cats = f"='{sheet_name}'!$A$2:$A${last_row}"
    for j, name in enumerate(series_dict.keys(), start=1):
        ch = wb.add_chart({"type": "line"})
        col_letter = chr(64 + j + 0)  # B, C, ...
        ch.add_series({
            "name":       f"='{sheet_name}'!${col_letter}$1",
            "categories": cats,
            "values":     f"='{sheet_name}'!${col_letter}$2:${col_letter}${last_row}",
        })
        ch.set_title({"name": name})
        ch.set_y_axis({"num_format": "#,##0.0000"})
        ws.insert_chart(3 + (j-1)*16, 4, ch, {"x_scale": 1.2, "y_scale": 1.1})
    ws.set_column(0, 0, 12)
    ws.set_column(1, len(series_dict), 16)
    return ws

def _mx_news_get_v1(max_items=12):
    feeds = [
        ("Google News MX – Economía",
         "https://news.google.com/rss/headlines/section/topic/BUSINESS?hl=es-419&gl=MX&ceid=MX:es-419"),
        ("Google News MX – BMV",
         "https://news.google.com/rss/search?q=Bolsa%20Mexicana%20de%20Valores&hl=es-419&gl=MX&ceid=MX:es-419"),
        ("Google News MX – Banxico",
         "https://news.google.com/rss/search?q=Banxico&hl=es-419&gl=MX&ceid=MX:es-419"),
        ("Google News MX – Inflación INEGI",
         "https://news.google.com/rss/search?q=inflaci%C3%B3n%20M%C3%A9xico%20INEGI&hl=es-419&gl=MX&ceid=MX:es-419"),
    ]
    items = []
    try:
        import feedparser as _fp
    except Exception:
        return items
    for source, url in feeds:
        try:
            fp = _fp.parse(url)
            for e in fp.get("entries", []):
                title = (e.get("title") or "").strip()
                link  = (e.get("link") or "").strip()
                pub   = e.get("published") or e.get("updated") or ""
                try:
                    dt = parsedate_to_datetime(pub) if pub else None
                except Exception:
                    dt = None
                items.append({"title": title, "link": link, "published_dt": dt, "source": source})
        except Exception:
            continue
    items.sort(key=lambda x: x["published_dt"] or datetime(1970,1,1), reverse=True)
    return items[:max_items]

def _mx_news_write_v1(wb, news_list, sheet_name="Noticias_RSS"):
    if not news_list:
        return None
    ws = wb.add_worksheet(sheet_name)
    fmt_bold = wb.add_format({'font_name': 'Arial', "bold": True})
    fmt_link = wb.add_format({'font_name': 'Arial', "font_color": "blue", "underline": 1})
    fmt_date = wb.add_format({'font_name': 'Arial', "num_format": "yyyy-mm-dd hh:mm"})
    ws.write_row(0, 0, ["Título", "Link", "Fecha", "Fuente"], fmt_bold)
    for i, n in enumerate(news_list, start=1):
        ws.write_string(i, 0, n.get("title",""))
        link = n.get("link","")
        if link:
            ws.write_url(i, 1, link, fmt_link, string="Abrir")
        dt = n.get("published_dt")
        if dt:
            try:
                ws.write_datetime(i, 2, dt, fmt_date)
            except Exception:
                ws.write_string(i, 2, str(dt))
        ws.write_string(i, 3, n.get("source",""))
    ws.set_column(0, 0, 80); ws.set_column(1, 1, 12)
    ws.set_column(2, 2, 20); ws.set_column(3, 3, 18)
    return ws

import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

DEBUG = False  

def _noop(*args, **kwargs):
    """Función vacía para suprimir salidas visibles."""
    return None

if not DEBUG:
    
    import builtins as _b
    _b.print = _noop

   
    try:
        st.write   = _noop   
        st.json    = _noop   
        st.success = _noop   
        
    except Exception:
        pass


st.set_page_config(page_title="IMEMSA - Indicadores", layout="wide")

st.markdown("""
<style>
/* ---------- Layout general ---------- */
.block-container { padding-top: 1.5rem; }

/* Encabezado */
.imemsa-header {
  display: flex; gap: 1.25rem; align-items: center; 
  margin-bottom: 0.75rem;
}

/* Logo */
.imemsa-logo img {
  max-height: 5px;        
  width: auto;
  border-radius: 10px;
}

/* Títulos */
.imemsa-title {
  line-height: 1.1;
}
.imemsa-title h1 {
  margin: 0 0 0.25rem 0; 
  font-size: clamp(1.6rem, 2.4vw, 2.2rem);
  font-weight: 500;
}
.imemsa-title h3 {
  margin: 0; 
  font-weight: 500; 
  opacity: 0.95;
}

/* Línea divisoria con colores del logo */
.imemsa-divider {
  height: 6px;
  width: 100%;
  border-radius: 999px;
  margin: 0.75rem 0 1rem 0;
  background: linear-gradient(90deg, #0A4FA3 0%, #0A4FA3 40%, #E32028 40%, #E32028 60%, #0A4FA3 60%, #0A4FA3 100%);
}

/* Espaciado inferior tras el header */
.imemsa-spacer { height: 12px; }
</style>
""", unsafe_allow_html=True)

st.markdown(
    """
    <div class="imemsa-header">
      <div class="imemsa-logo">
        <img src="logo.png" alt="IMEMSA logo">
      </div>
      <div class="imemsa-title">
        <h1>Indicadores de Tipo de Cambio</h1>
      </div>
    </div>
    <div class="imemsa-divider"></div>
    <div class="imemsa-spacer"></div>
    """,
    unsafe_allow_html=True
)

import os, pytz as _pytz_for_login  

def _get_app_password() -> str:
    try:
        return st.secrets["APP_PASSWORD"]
    except Exception:
        pass
    if os.getenv("APP_PASSWORD"):
        return os.getenv("APP_PASSWORD")
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
    st.text_input("Contraseña", type="password", key="password_input", on_change=_try_login, placeholder="Escribe tu contraseña…")
    st.stop()

CDMX = pytz.timezone("America/Mexico_City")

def today_cdmx():
    return datetime.now(CDMX).date()

def now_ts():
    return datetime.now(CDMX).strftime("%Y-%m-%d %H:%M:%S")

def try_float(x):
    try:
        return float(str(x).replace(",", "").strip())
    except:
        return None

def logo_image_or_emoji():
    p = Path("logo.png")
    return "🛟" if not p.exists() else "logo.png"

def logo_base64(max_height_px: int = 40):
    """Devuelve base64 de logo.png si existe; si no, None."""
    try:
        p = Path("logo.png")
        if not p.exists():
            return None
        im = Image.open(p)
        w, h = im.size
        if h > max_height_px:
            im = im.resize((int(w * max_height_px / h), max_height_px))
        bio = io.BytesIO()
        im.save(bio, format="PNG")
        return base64.b64encode(bio.getvalue()).decode("ascii")
    except Exception:
        return None


BANXICO_TOKEN = "677aaedf11d11712aa2ccf73da4d77b6b785474eaeb2e092f6bad31b29de6609"
INEGI_TOKEN   = "0146a9ed-b70f-4ea2-8781-744b900c19d1"
FRED_TOKEN    = "b4f11681f441da78103a3706d0dab1cf"  

def fred_fetch_series(series_id: str, start: str | None = None, end: str | None = None, units: str = "lin"):
    """
    Consulta FRED 
    """
    try:
        token = FRED_TOKEN.strip()
    except Exception:
        token = ""
    if not token:
        return []
    params = {"series_id": series_id, "api_key": token, "file_type": "json", "units": units}
    if start: params["observation_start"] = start
    if end:   params["observation_end"]   = end
    try:
        r = requests.get("https://api.stlouisfed.org/fred/series/observations", params=params, timeout=20)
        if r.status_code != 200:
            return []
        data = r.json().get("observations", [])
        out = []
        for row in data:
            d = row.get("date")
            v = row.get("value")
            try:
                v = float(v)
            except Exception:
                v = None
            out.append({"date": d, "value": v})
        return out
    except Exception:
        return []


TZ_MX = pytz.timezone("America/Mexico_City")


st.set_page_config(
    page_title="Indicadores Tipos de Cambio",
    page_icon=logo_image_or_emoji(),
    layout="centered"
)
_check_password() 


st.markdown("""
<style>
#MainMenu {visibility: hidden;}      /* oculta hamburguesa */
footer {visibility: hidden;}         /* oculta footer */

.app-header {
  position: sticky; top: 0; z-index: 999;
  background: white; border-bottom: 1px solid #eee;
  display: flex; align-items: center; gap: 16px;
  padding: 8px 6px;
}
.app-header img.logo { height: 40px; }
.app-header .titles h1 {
  font-size: 20px; margin: 0;
}
.app-header .titles p {
  margin: 0; color: #666;
}
</style>
""", unsafe_allow_html=True)

_logo_b64 = logo_base64()
if _logo_b64:
    st.markdown(
        f"""
        <div class="app-header">
          <img class="logo" src="data:image/png;base64,{_logo_b64}" alt="logo"/>
          <div class="titles">
            <h1>Indicadores (últimos 5 días) + Noticias</h1>
            <p> </p>
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )
else:
   
    st.title("📈 Indicadores (últimos 5 días) + Noticias")
    st.caption("Excel con tipos de cambio, noticias y gráficos.")

if _logo_b64:
 
    st.sidebar.image(f"data:image/png;base64,{_logo_b64}", use_container_width=True)

def http_session(timeout=15):
    s = requests.Session()
    retries = Retry(total=3, backoff_factor=0.8,
                    status_forcelist=[429, 500, 502, 503, 504],
                    allowed_methods=frozenset(["GET"]))
    s.mount("https://", HTTPAdapter(max_retries=retries))
    s.mount("http://", HTTPAdapter(max_retries=retries))
    s.request = (lambda orig: (lambda *a, **k: orig(*a, timeout=k.pop("timeout", timeout), **k)))(s.request)
    return s

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

def _check_tokens():
    missing = []
    if not BANXICO_TOKEN.strip(): missing.append("BANXICO_TOKEN")
    if not INEGI_TOKEN.strip():   missing.append("INEGI_TOKEN")
    if missing:
        st.error("Faltan tokens: " + ", ".join(missing))
        st.stop()

@st.cache_data(ttl=60*30)
def sie_opportuno(series_id):
    url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series_id}/datos/oportuno"
    headers = {"Bmx-Token": BANXICO_TOKEN}
    r = http_session().get(url, headers=headers, timeout=15)
    r.raise_for_status()
    return r.json()

def sie_latest(series_id):
    try:
        data = sie_opportuno(series_id)
        serie = data["bmx"]["series"][0]["datos"]
        if not serie: return None, None
        last = serie[-1]
        return last["fecha"], try_float(last["dato"])
    except:
        return None, None

@st.cache_data(ttl=60*30)
def sie_range(series_id: str, start_iso: str, end_iso: str):
    url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series_id}/datos/{start_iso}/{end_iso}"
    headers = {"Bmx-Token": BANXICO_TOKEN}
    r = http_session(20).get(url, headers=headers, timeout=20)
    r.raise_for_status()
    j = r.json()
    series = j.get("bmx", {}).get("series", [])
    if not series:
        return []
    return series[0].get("datos", []) or []

def sie_last_n(series_id: str, n: int = 6):
    end = today_cdmx()
    start = end - timedelta(days=2*365)
    obs = sie_range(series_id, start.isoformat(), end.isoformat())
    vals = []
    for o in obs:
        f = o.get("fecha"); v = try_float(o.get("dato"))
        if f and (v is not None):
            vals.append((f, v))
    if not vals:
        return []
    vals.sort(key=lambda x: parse_any_date(x[0]) or datetime.utcnow())
    return vals[-n:]

def _ffill_by_dates(map_vals: dict, dates: list):
    # keys are ISO-like strings. choose last available <= date
    from datetime import datetime
    def to_dt(s):
        try:
            if "/" in s:
                return datetime.strptime(s, "%d/%m/%Y").date()
            return datetime.fromisoformat(s).date()
        except Exception:
            return None
    pairs = sorted([(to_dt(k), v) for k,v in map_vals.items() if to_dt(k)], key=lambda x: x[0])
    out = []
    last = None
    for ds in dates:
        d = to_dt(ds)
        if d is None:
            out.append(None); continue
        while pairs and pairs[0][0] <= d:
            last = pairs.pop(0)[1]
        out.append(last)
    return out

def _ffill_with_flags(map_vals: dict, dates: list):
    # Similar a _ffill_by_dates pero devuelve (valores, flags_ffill)
    from datetime import datetime
    def to_dt(s):
        try:
            if isinstance(s, str) and "/" in s:
                return datetime.strptime(s, "%d/%m/%Y").date()
            return datetime.fromisoformat(str(s)).date()
        except Exception:
            return None
    normalized = {}
    for k, v in map_vals.items():
        kd = to_dt(k)
        if kd:
            normalized[kd.isoformat()] = v
    out_vals, out_flags = [], []
    last = None
    for ds in dates:
        key = ds if isinstance(ds, str) else (ds.isoformat() if ds else None)
        if key in normalized and normalized[key] is not None:
            last = normalized[key]
            out_vals.append(last)
            out_flags.append(False)
        else:
            out_vals.append(last)
            out_flags.append(last is not None)
    return out_vals, out_flags

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

@st.cache_data(ttl=60*60)
def get_uma(inegi_token: str):
    """
    UMA nacional: 620706 (diaria), 620707 (mensual), 620708 (anual)
    Retorna: {'fecha','diaria','mensual','anual','_status','_source'}
    """
    base = "https://www.inegi.org.mx/app/api/indicadores/desarrolladores/jsonxml/INDICATOR"
    ids = "620706,620707,620708"
    urls = [
        f"{base}/{ids}/es/00/true/BISE/2.0/{inegi_token}?type=json",
        f"{base}/{ids}/es/00/true/BIE/2.0/{inegi_token}?type=json",  
    ]

    def _num(x):
        try:
            return float(str(x).replace(",", ""))
        except:
            return None

    last_err = None
    for u in urls:
        try:
            r = http_session(20).get(u, timeout=20)
            if r.status_code != 200:
                last_err = f"HTTP {r.status_code}"
                continue
            data = r.json()
            series = data.get("Series") or data.get("series") or []
            if not series:
                last_err = "Sin 'Series'"; continue

            def last_obs(s):
                obs = s.get("OBSERVATIONS") or s.get("observations") or []
                return obs[-1] if obs else None

            d_obs = last_obs(series[0]); m_obs = last_obs(series[1]) if len(series)>1 else None
            a_obs = last_obs(series[2]) if len(series)>2 else None

            def get_v(o):
                if not o: return None
                return _num(o.get("OBS_VALUE") or o.get("value"))
            def get_f(o):
                if not o: return None
                return o.get("TIME_PERIOD") or o.get("periodo") or o.get("time_period") or o.get("fecha")

            return {
                "fecha": get_f(d_obs) or get_f(m_obs) or get_f(a_obs),
                "diaria":  get_v(d_obs),
                "mensual": get_v(m_obs),
                "anual":   get_v(a_obs),
                "_status": "ok",
                "_source": "INEGI",
            }
        except Exception as e:
            last_err = str(e)
            continue

    return {"fecha": None, "diaria": None, "mensual": None, "anual": None,
            "_status": f"err: {last_err}", "_source": "fallback"}

def _probe(fn, ok_condition):
    t0 = time.time()
    try:
        res = fn()
        ms = int((time.time() - t0)*1000)
        ok = ok_condition(res)
        return ("ok" if ok else "warn"), ("OK" if ok else "Parcial"), ms
    except Exception as e:
        return ("err", "Error", 0)

def _render_sidebar_status():
    st.sidebar.header("🔎 Estado de fuentes")
    st.sidebar.caption(f"Última verificación: {now_ts()}")

    b_status, b_msg, b_ms = _probe(lambda: sie_latest(SIE_SERIES["USD_FIX"]),
                                   lambda res: "ok" if isinstance(res, tuple) and res[0] and (res[1] is not None) else "err")
    i_status, i_msg, i_ms = _probe(lambda: get_uma(INEGI_TOKEN),
                                   lambda res: "ok" if isinstance(res, dict) and (res.get("diaria") is not None) else ("warn" if isinstance(res, dict) else "err"))
    f_status, f_msg, f_ms = ("warn", "Sin token (fallback)", 0) if not FRED_TOKEN.strip() else ("ok", "OK", 0)

    def badge(status, label, msg, ms):
        dot = "🟢" if status=="ok" else ("🟡" if status=="warn" else "🔴")
        st.sidebar.write(f"{dot} **{label}** — {msg} · {ms} ms")

    badge(b_status, "Banxico (SIE)", b_msg, b_ms)
    badge(i_status, "INEGI (UMA)",  i_msg, i_ms)
    badge(f_status, "FRED (USA)",   f_msg, f_ms)

    st.sidebar.divider()

with st.sidebar.expander("🔑 Tokens de APIs", expanded=False):
    st.caption("Si ingresas un token aquí, la app lo usará en lugar del definido en el código.")
    token_banxico_input = st.text_input("BANXICO_TOKEN", value="", type="password")
    token_inegi_input   = st.text_input("INEGI_TOKEN",   value="", type="password")
    
    if token_banxico_input.strip():
        BANXICO_TOKEN = token_banxico_input.strip()
    if token_inegi_input.strip():
        INEGI_TOKEN = token_inegi_input.strip()

    with st.sidebar.expander("Herramientas"):
        c1, c2 = st.columns(2)
        if c1.button("Limpiar cachés Banxico"):
            sie_opportuno.clear(); sie_range.clear()
        if c2.button("Limpiar caché UMA"):
            get_uma.clear()
    with st.sidebar.expander("Diagnóstico UMA"):
        if st.button("Probar INEGI ahora"):
            res = get_uma(INEGI_TOKEN)






with st.expander("📄 Selecciona las Hojas del Excel que contendra tu archivo", expanded=True):
    st.caption("Activa/desactiva hojas opcionales del archivo Excel")
    want_fred   = st.checkbox("Agregar hoja FRED", value=st.session_state.get("want_fred", False))
    want_news   = st.checkbox("Agregar hoja Noticias_RSS", value=st.session_state.get("want_news", False))
    want_charts = st.checkbox("Agregar hoja 'Gráficos' ", value=st.session_state.get("want_charts", False))
    want_raw    = st.checkbox("Agregar hoja 'Datos crudos' ", value=st.session_state.get("want_raw", False))
    st.session_state["want_fred"] = want_fred
    st.session_state["want_news"] = want_news
    st.session_state["want_charts"] = want_charts
    st.session_state["want_raw"] = want_raw

# Sincroniza flags de Excel con las opciones del UI
do_fred   = st.session_state.get("want_fred", False)
do_news   = st.session_state.get("want_news", False)
do_charts = st.session_state.get("want_charts", False)
do_raw    = st.session_state.get("want_raw", False)

# Parámetros fijos (Opciones retiradas del UI)
movex_win = 5
margen_pct = 0.20  # 0.20% por lado
import os
UMA_DIARIA = 0.0
try:
    UMA_DIARIA = float(st.secrets.get("UMA_DIARIA", os.getenv("UMA_DIARIA","0") or "0"))
except Exception:
    UMA_DIARIA = 0.0
uma_manual = UMA_DIARIA


_check_tokens()
_render_sidebar_status()

if st.button("Generar Excel"):
    def pad6(lst): return ([None]*(6-len(lst)))+lst if len(lst) < 6 else lst[-6:]
    none6 = [None]*6

    fix6 = pad6([v for _, v in sie_last_n(SIE_SERIES["USD_FIX"], n=6)])
    
    eur6 = pad6([v for _, v in sie_last_n(SIE_SERIES["EUR_MXN"], n=6)])
    jpy6 = pad6([v for _, v in sie_last_n(SIE_SERIES["JPY_MXN"], n=6)])

    
    movex_series = rolling_movex_for_last6(window=movex_win)
    movex6 = pad6(movex_series)

    
    cetes28_6 = pad6([v for _, v in sie_last_n(SIE_SERIES["CETES_28"], n=6)])
    cetes91_6 = pad6([v for _, v in sie_last_n(SIE_SERIES["CETES_91"], n=6)])
    cetes182_6 = pad6([v for _, v in sie_last_n(SIE_SERIES["CETES_182"], n=6)])
    cetes364_6 = pad6([v for _, v in sie_last_n(SIE_SERIES["CETES_364"], n=6)])

    
    uma = get_uma(INEGI_TOKEN)
    if uma.get("diaria") is None and uma_manual > 0:
        uma["diaria"]  = uma_manual
        uma["mensual"] = uma_manual * 30.4
        uma["anual"]   = uma["mensual"] * 12


    fred_rows = None
    try:
        if add_fred and fred_id.strip() and isinstance(fred_start, (datetime, date)) and isinstance(fred_end, (datetime, date)):
            fred_rows = fred_fetch_series(
                series_id=fred_id.strip(),
                start=fred_start.isoformat(),
                end=fred_end.isoformat(),
                units=fred_units
            )
    except NameError:
        fred_rows = None  
    bio = io.BytesIO()
    wb = xlsxwriter.Workbook(bio, {'in_memory': True})


    fmt_bold  = wb.add_format({'font_name': 'Arial', 'bold': True})
    fmt_hdr   = wb.add_format({'font_name': 'Arial', 'bold': True, 'bg_color': '#F2F2F2', 'align':'center'})
    fmt_num4  = wb.add_format({'font_name': 'Arial', 'num_format': '0.0000'})
    fmt_num6  = wb.add_format({'font_name': 'Arial', 'num_format': '0.000000'})
    fmt_wrap  = wb.add_format({'font_name': 'Arial', 'text_wrap': True})
    fmt_date_dm = wb.add_format({'font_name': 'Arial', 'num_format': 'dd "de" mmm'})

    fmt_all = wb.add_format({'font_name': 'Arial', 'font_name': 'Arial'})
    # Formatos adicionales
    fmt_num4_ffill = wb.add_format({'font_name': 'Arial', 'num_format': '0.0000', 'italic': True, 'font_color': '#666666'})
    fmt_num6_ffill = wb.add_format({'font_name': 'Arial', 'num_format': '0.000000', 'italic': True, 'font_color': '#666666'})
    fmt_pct2      = wb.add_format({'font_name': 'Arial', 'num_format': '0.00%'})
    fmt_pct2_ffill= wb.add_format({'font_name': 'Arial', 'num_format': '0.00%', 'italic': True, 'font_color': '#666666'})

    end = today_cdmx()
    # Últimos 6 días hábiles (lun-vie), incluyendo hoy si aplica
    header_dates_date = []
    d = end
    while len(header_dates_date) < 6:
        if d.weekday() < 5:  # 0=lunes, 6=domingo
            header_dates_date.append(d)
        d -= timedelta(days=1)
    header_dates_date = list(reversed(header_dates_date))
    # Lista paralela en ISO para consultas a diccionarios de series
    header_dates = [x.isoformat() for x in header_dates_date]

    def _as_map(pairs): return {d:v for d,v in pairs}
    # Construir mapas FX con rango que cubra el span del encabezado (con buffer) y luego ffill
    fx_start = (header_dates_date[0] - timedelta(days=30)).isoformat()
    fx_end   = header_dates_date[-1].isoformat()
    def _as_map_from_range(series_key):
        obs = sie_range(SIE_SERIES[series_key], fx_start, fx_end)
        m = {}
        for o in obs:
            _f = parse_any_date(o.get('fecha'))
            _v = try_float(o.get('dato'))
            if _f and (_v is not None):
                m[_f.date().isoformat()] = _v
        return m
    m_fix  = _as_map_from_range('USD_FIX')
    m_eur  = _as_map_from_range('EUR_MXN')
    m_jpy  = _as_map_from_range('JPY_MXN')
    m_udis = _as_map_from_range('UDIS')
    # Alinear a header_dates con forward-fill para evitar huecos en días sin publicación
    fix_vals, fix_fflags = _ffill_with_flags(m_fix, header_dates)
    eur_vals, eur_fflags = _ffill_with_flags(m_eur, header_dates)
    jpy_vals, jpy_fflags = _ffill_with_flags(m_jpy, header_dates)
    m_c28  = _as_map_from_range('CETES_28')
    m_c91  = _as_map_from_range('CETES_91')
    m_c182 = _as_map_from_range('CETES_182')
    m_c364 = _as_map_from_range('CETES_364')
    cetes28, cetes28_f = _ffill_with_flags(m_c28, header_dates)
    cetes91, cetes91_f = _ffill_with_flags(m_c91, header_dates)
    cetes182, cetes182_f = _ffill_with_flags(m_c182, header_dates)
    cetes364, cetes364_f = _ffill_with_flags(m_c364, header_dates)


    try:
        movex6  
    except NameError:
        movex6 = rolling_movex_for_last6(window=movex_win)
    compra = [(x*(1 - margen_pct/100) if x is not None else None) for x in movex6]
    venta  = [(x*(1 + margen_pct/100) if x is not None else None) for x in movex6]
    usd_jpy = [((u/j) if (u is not None and j not in (None, 0)) else None) for u,j in zip(fix_vals, jpy_vals)]
    eur_usd = [((e/u) if (e is not None and u not in (None, 0)) else None) for e,u in zip(eur_vals, fix_vals)]

    try:
        uma  
    except NameError:
        uma = get_uma(INEGI_TOKEN)

    def _last_or_none(series_pairs): 
        return series_pairs[-1][1] if series_pairs else None

        
    # TIIE (Banxico SIE) - usar histórico y alinear por fechas del encabezado
    m_t28  = _as_map_from_range('TIIE_28')
    m_t91  = _as_map_from_range('TIIE_91')
    m_t182 = _as_map_from_range('TIIE_182')
    # Objetivo (tasa de política monetaria)
    _obs_obj = sie_range('SF61745', fx_start, fx_end)
    m_obj = {}
    for o in _obs_obj:
        _f = parse_any_date(o.get('fecha'))
        _v = try_float(o.get('dato'))
        if _f and (_v is not None):
            m_obj[_f.date().isoformat()] = _v

    tiie28, tiie28_f = _ffill_with_flags(m_t28, header_dates)
    tiie91, tiie91_f = _ffill_with_flags(m_t91, header_dates)
    tiie182, tiie182_f = _ffill_with_flags(m_t182, header_dates)
    tiie_obj, tiie_obj_f = _ffill_with_flags(m_obj, header_dates)


    # --- Fallback robusto para TIIE 182 días ---
    try:
        # Si todo quedó en None (no hubo match de fechas o no hay histórico),
        # 
        if all(v is None for v in tiie182):
            v182_op = None
            try:
                _, v182_op = sie_latest(SIE_SERIES["TIIE_182"], BANXICO_TOKEN)
            except Exception:
                v182_op = None
            if v182_op is not None:
                # Replicamos el oportuno a las 6 columnas
                tiie182 = [round(float(v182_op), 4)] * len(header_dates)
            else:
                # Como segunda opción, tomamos el último 'last_n' y replicamos
                try:
                    _pairs182 = sie_last_n(SIE_SERIES["TIIE_182"], 6, BANXICO_TOKEN)
                    _last = _pairs182[-1][1] if _pairs182 else None
                    if _last is not None:
                        tiie182 = [round(float(_last), 4)] * len(header_dates)
                except Exception:
                    pass
    except Exception:
        pass
    # --- /fallback ---
    ws = wb.add_worksheet("Indicadores")
    ws.set_column(0, 6, 16)
    # ULTIMA MODIFICACION
    try:
        ws.insert_image(
        'A1',
        'logo.png',
        {'x_scale': 0.35, 'y_scale': 0.25}
        )
    except Exception:
        pass
    try:
        ws.set_column(0, 50, None, fmt_all)
        ws.hide_gridlines(2)
    except Exception:
        pass

    # Claridad inmediata: anchos y congelar encabezado (hasta B3)
    ws.set_column(0, 0, 22)   # columna A (rótulos)
    ws.set_column(1, 7, 13)   # columnas B..H (fechas y sparklines)
    ws.freeze_panes(2, 1)

    # Leyenda para arrastres (ffill)
    ws.write(0, 7, '* Valor copiado cuando no hay publicación del día', wb.add_format({'font_name': 'Arial', 'italic': True, 'font_color': '#666'}))


    ws.write(1, 0, "Fecha:", fmt_bold)
    from datetime import datetime as _dt
    for i, d in enumerate(header_dates_date):
        ws.write_datetime(1, 1+i, _dt(d.year, d.month, d.day), fmt_date_dm)

    ws.write(3, 0, "TIPOS DE CAMBIO:", fmt_bold)
    ws.write(5, 0, "DÓLAR AMERICANO.", fmt_bold)
    ws.write(6, 0, "Dólar/Pesos:")
    for i, v in enumerate(fix_vals):
        ws.write(6, 1+i, v, fmt_num4_ffill if (fix_fflags[i]) else fmt_num4)
    ws.write(7, 0, "MONEX:")

    ws.write(8, 0, "Compra:")
    for i, v in enumerate(compra):
        ws.write(8, 1+i, v, fmt_num6)
    ws.write(9, 0, "Venta:")
    for i, v in enumerate(venta):
        ws.write(9, 1+i, v, fmt_num6)
    # Asegura explícitamente las celdas G9/G10 (columna 6, fila 8 y 9)
    try:
        if compra: ws.write(8, 6, compra[-1], fmt_num6)
        if venta:  ws.write(9, 6, venta[-1],  fmt_num6)
    except Exception:
        pass


    ws.write(11, 0, "YEN JAPONÉS.", fmt_bold)
    ws.write(12, 0, "Yen Japonés/Peso:")
    for i, v in enumerate(jpy_vals):
        ws.write(12, 1+i, v, fmt_num6_ffill if (jpy_fflags[i]) else fmt_num6)
    ws.write(13, 0, "Dólar/Yen Japonés:")
    for i, v in enumerate(usd_jpy):
        ws.write(13, 1+i, v, fmt_num6)

    ws.write(15, 0, "EURO.", fmt_bold)
    ws.write(16, 0, "Euro/Peso:")
    for i, v in enumerate(eur_vals):
        ws.write(16, 1+i, v, fmt_num6_ffill if (eur_fflags[i]) else fmt_num6)
    ws.write(17, 0, "Euro/Dólar:")
    for i, v in enumerate(eur_usd):
        ws.write(17, 1+i, v, fmt_num6)

    ws.write(19, 0, "UDIS:", fmt_bold)
    ws.write(21, 0, "UDIS: ")
    # Trae rango suficiente para cubrir el span de header_dates (días hábiles)
    udi_start = (header_dates_date[0] - timedelta(days=30)).isoformat()
    udi_end   = header_dates_date[-1].isoformat()
    udi_obs   = sie_range(SIE_SERIES["UDIS"], udi_start, udi_end)
    m_udis_r  = {}
    for o in udi_obs:
        _f = parse_any_date(o.get("fecha"))
        _v = try_float(o.get("dato"))
        if _f and (_v is not None):
            m_udis_r[_f.date().isoformat()] = _v
    udis_vals, udis_fflags = _ffill_with_flags(m_udis_r, header_dates)
    for i, v in enumerate(udis_vals):
        ws.write(21, 1+i, v, fmt_num6_ffill if (udis_fflags[i]) else fmt_num6)

    ws.write(23, 0, "TASAS TIIE:", fmt_bold)
    ws.write(25, 0, "TIIE objetivo:")
    ws.write(26, 0, "TIIE 28 Días:")
    ws.write(27, 0, "TIIE 91 Días:")
    ws.write(28, 0, "TIIE 182 Días:")
    for i in range(6):
        vobj = (tiie_obj[i]/100.0) if (tiie_obj[i] is not None) else None
        ws.write(25, 1+i, vobj, fmt_pct2_ffill if (tiie_obj_f[i]) else fmt_pct2)
        v28 = (tiie28[i]/100.0) if (tiie28[i] is not None) else None
        ws.write(26, 1+i, v28, fmt_pct2_ffill if (tiie28_f[i]) else fmt_pct2)
        v91 = (tiie91[i]/100.0) if (tiie91[i] is not None) else None
        ws.write(27, 1+i, v91, fmt_pct2_ffill if (tiie91_f[i]) else fmt_pct2)
        v182 = (tiie182[i]/100.0) if (tiie182[i] is not None) else None
        ws.write(28, 1+i, v182, fmt_pct2_ffill if (tiie182_f[i]) else fmt_pct2)
    ws.write(30, 0, "CETES:", fmt_bold)
    ws.write(32, 0, "CETES 28 Días:")
    ws.write(33, 0, "CETES 91 Días:")
    ws.write(34, 0, "CETES 182 Días:")
    ws.write(35, 0, "CETES 364 Días:")
    for i in range(6):
        v0 = (cetes28[i]/100.0) if (cetes28[i] is not None) else None
        ws.write(32, 1+i, v0, fmt_pct2_ffill if (cetes28_f[i]) else fmt_pct2)
        v1 = (cetes91[i]/100.0) if (cetes91[i] is not None) else None
        ws.write(33, 1+i, v1, fmt_pct2_ffill if (cetes91_f[i]) else fmt_pct2)
        v2 = (cetes182[i]/100.0) if (cetes182[i] is not None) else None
        ws.write(34, 1+i, v2, fmt_pct2_ffill if (cetes182_f[i]) else fmt_pct2)
        v3 = (cetes364[i]/100.0) if (cetes364[i] is not None) else None
        ws.write(35, 1+i, v3, fmt_pct2_ffill if (cetes364_f[i]) else fmt_pct2)
    ws.write(37, 0, "UMA:", fmt_bold)
    ws.write(39, 0, "Diario:");  ws.write(39, 1, uma.get("diaria"))
    ws.write(40, 0, "Mensual:"); ws.write(40, 1, uma.get("mensual"))
    ws.write(41, 0, "Anual:");   ws.write(41, 1, uma.get("anual"))

do_raw = globals().get('do_raw', True)
if do_raw and ('wb' in globals()):
    ws3 = wb.add_worksheet("Datos crudos")
    try:
        ws3.set_column(0, 50, None, fmt_all)
        ws3.hide_gridlines(2)
    except Exception:
        pass
    ws3.write(0,0,"Serie", fmt_hdr); ws3.write(0,1,"Fecha", fmt_hdr); ws3.write(0,2,"Valor", fmt_hdr)
    def _dump(ws_sheet, start_row, tag, pairs):
        r = start_row
        for d, v in pairs:
            ws_sheet.write(r, 0, tag)
            ws_sheet.write(r, 1, d)
            ws_sheet.write(r, 2, v)
            r += 1

        return r

    r = 1
    r = _dump(ws3, r, "USD/MXN (FIX)", sie_last_n(SIE_SERIES["USD_FIX"], 6))
    r = _dump(ws3, r, "EUR/MXN",       sie_last_n(SIE_SERIES["EUR_MXN"], 6))
    r = _dump(ws3, r, "JPY/MXN",       sie_last_n(SIE_SERIES["JPY_MXN"], 6))
    r = _dump(ws3, r, "UDIS",          sie_last_n(SIE_SERIES["UDIS"],    6))
    r = _dump(ws3, r, "CETES 28d (%)", sie_last_n(SIE_SERIES["CETES_28"],6))
    r = _dump(ws3, r, "CETES 91d (%)", sie_last_n(SIE_SERIES["CETES_91"],6))
    r = _dump(ws3, r, "CETES 182d (%)",sie_last_n(SIE_SERIES["CETES_182"],6))
    r = _dump(ws3, r, "CETES 364d (%)",sie_last_n(SIE_SERIES["CETES_364"],6))
    ws3.set_column(0, 0, 18); ws3.set_column(1, 1, 12); ws3.set_column(2, 2, 16)

    
do_charts = globals().get('do_charts', True)
if do_charts and ('wb' in globals()):
    ws4 = wb.add_worksheet("Gráficos")
    try:
        ws4.set_column(0, 50, None, fmt_all)
        ws4.hide_gridlines(2)
    except Exception:
        pass
    chart1 = wb.add_chart({'type': 'line'})
    chart1.add_series({
    'name':       "USD/MXN (FIX)",
    'categories': "=Indicadores!$B$2:$G$2",
    'values':     "=Indicadores!$B$7:$G$7",
    })
    chart1.set_title({'name': 'USD/MXN (FIX)'})
    ws4.insert_chart('B2', chart1, {'x_scale': 1.3, 'y_scale': 1.2})

    chart2 = wb.add_chart({'type': 'line'})
    for row in (33,34,35,36):
        chart2.add_series({
            'name':       f"=Indicadores!$A${row}",
            'categories': "=Indicadores!$B$2:$G$2",
            'values':     f"=Indicadores!$B${row}:$G${row}",
        })
    chart2.set_title({'name': 'CETES (%)'})
    ws4.insert_chart('B18', chart2, {'x_scale': 1.3, 'y_scale': 1.2})


    try:
        if fred_rows and st.session_state.get('want_fred', False):
            wsname  = f"FRED_{fred_id[:25]}"
            wsfred  = wb.add_worksheet(wsname)

            try:
                wsfred.set_column(0, 50, None, fmt_all)
                wsfred.hide_gridlines(2)
            except Exception:
                pass
            fmt_bold = wb.add_format({'font_name': 'Arial', "bold": True})
            fmt_num  = wb.add_format({'font_name': 'Arial', "num_format": "#,##0.0000"})
            fmt_date = wb.add_format({'font_name': 'Arial', "num_format": "yyyy-mm-dd"})

            
            wsfred.write(0, 0, f"FRED – {fred_id}", fmt_bold)
            wsfred.write(1, 0, f"Generado: {today_cdmx('%Y-%m-%d %H:%M')} (CDMX)")
            wsfred.write_row(3, 0, ["date", fred_id], fmt_bold)

            
            r_start = 4
            r = r_start
            valid_count = 0

            for row in fred_rows:
                d = row.get("date")
                v = row.get("value")

                
                try:
                    dt = pd.to_datetime(d).to_pydatetime()
                    wsfred.write_datetime(r, 0, dt, fmt_date)
                except Exception:
                    wsfred.write(r, 0, str(d))

                
                try:
                    if v is not None:
                        v_float = float(v)
                        if not pd.isna(v_float):
                            wsfred.write_number(r, 1, v_float, fmt_num)
                            valid_count += 1
                        else:
                            wsfred.write_blank(r, 1, None)
                    else:
                        wsfred.write_blank(r, 1, None)
                except Exception:
                    wsfred.write_blank(r, 1, None)

                r += 1

            wsfred.set_column(0, 0, 12)
            wsfred.set_column(1, 1, 16)

            if valid_count >= 2:
                first_excel_row = r_start + 1
                last_excel_row  = r
                ch = wb.add_chart({"type": "line"})
                ch.add_series({
                    "name": fred_id,
                    "categories": f"={wsname}!$A${first_excel_row}:$A${last_excel_row}",
                    "values":     f"={wsname}!$B${first_excel_row}:$B${last_excel_row}",
                })
                ch.set_title({"name": f"{fred_id} (FRED)"})
                ch.set_y_axis({"num_format": "#,##0.0000"})
                wsfred.insert_chart("D4", ch, {"x_scale": 1.2, "y_scale": 1.2})
    except Exception as _e:
        pass

try:
    fred_key = ""
    try:
        fred_key = st.secrets.get("FRED_API_KEY", "").strip()
    except Exception:
        pass
    if fred_key and st.session_state.get('want_fred', False):
        end_dt = datetime.now()
        start_dt = end_dt - timedelta(days=180)
        fred_series = {
            "US 10Y (DGS10)": "DGS10",
            "Fed Funds (DFF)": "DFF",
            "MXN/USD (DEXMXUS)": "DEXMXUS",
        }
        fred_data = {}
        for label, sid in fred_series.items():
            try:
                fred_data[label] = _fred_fetch_v1(
                    sid,
                    start_dt.strftime("%Y-%m-%d"),
                    end_dt.strftime("%Y-%m-%d"),
                    fred_key
                )
            except Exception:
                fred_data[label] = []
        if any(len(v) > 0 for v in fred_data.values()):
            _fred_write_v1(wb, fred_data, sheet_name="FRED_v2")

    _news = []
    try:
        _news = _mx_news_get_v1(max_items=12)
    except Exception:
        _news = []
    if _news and st.session_state.get('want_news', False):
        _mx_news_write_v1(wb, _news, sheet_name="Noticias_RSS")
except Exception:
    pass

try:
    # === Hoja 'Lógica de datos' (siempre en ruta normal) ===
    try:
        # Reutiliza formatos si existen; si no, crea mínimos
        try:
            fmt_all = fmt_all
        except NameError:
            fmt_all = wb.add_format({'font_name': 'Arial'})
        try:
            fmt_hdr = fmt_hdr
        except NameError:
            fmt_hdr = wb.add_format({'font_name': 'Arial', 'bold': True, 'bg_color': '#F2F2F2'})
        try:
            fmt_bold = fmt_bold
        except NameError:
            fmt_bold = wb.add_format({'font_name': 'Arial', 'bold': True})
        fmt_wrap = wb.add_format({'font_name': 'Arial', 'text_wrap': True})
    
        wsh_ld = wb.add_worksheet("Lógica de datos")
        wsh_ld.set_column(0, 0, 28, fmt_all)
        wsh_ld.set_column(1, 1, 95, fmt_all)
        try:
            wsh_ld.hide_gridlines(2)
        except Exception:
            pass
        wsh_ld.write(0, 0, "Sección", fmt_hdr)
        wsh_ld.write(0, 1, "Contenido", fmt_hdr)
    
        row = 1
        # Contenido detallado
        wsh_ld.write(row, 0, "Propósito", fmt_bold); wsh_ld.write(row, 1, "Concentrar indicadores (FX, UDIS, TIIE, CETES) para los últimos 6 días hábiles.", fmt_wrap); row += 1
        wsh_ld.write(row, 0, "Flujo de generación", fmt_bold); wsh_ld.write(row, 1, "1) Encabezado con días hábiles.\n2) Consulta Banxico SIE por rango.\n3) Normalización numérica.\n4) Forward‑fill por fecha.", fmt_wrap); row += 1
        wsh_ld.write(row, 0, "Fuentes / Series SIE", fmt_bold); wsh_ld.write(row, 1, "USD/MXN FIX (SF43718), EUR/MXN (SF46410), JPY/MXN (SF46406), UDIS (SP68257), CETES 28/91/182/364 (SF60634/5/6/7), TIIE 28/91/182 (SF60653/4/5), Tasa objetivo (SF61745).", fmt_wrap); row += 1
        wsh_ld.write(row, 0, "Tratamiento de datos", fmt_bold); wsh_ld.write(row, 1, "• Forward‑fill por fecha cuando falte publicación.\n• Tasas en %: si valor > 1.0, dividir entre 100 (11.25% = 0.1125).\n• UDIS / FX con 6/4 decimales según corresponda.", fmt_wrap); row += 1
        wsh_ld.write(row, 0, "Rangos con nombre", fmt_bold); wsh_ld.write(row, 1, "RANGO_FECHAS, RANGO_USDMXN, RANGO_EURMXN, RANGO_JPYMXN, RANGO_UDIS, RANGO_TOBJ, RANGO_TIIE28/91/182, RANGO_C28/91/182/364.", fmt_wrap); row += 1
        wsh_ld.write(row, 0, "Trazabilidad y metadatos", fmt_bold); wsh_ld.write(row, 1, "Ver hoja 'Metadatos': fecha/hora CDMX, zona, reglas y claves SIE. Cotejar encabezado con calendario hábil y disponibilidad de Banxico.", fmt_wrap); row += 1
        wsh_ld.write(row, 0, "Limitaciones", fmt_bold); wsh_ld.write(row, 1, "Feriados/rezagos de publicación; valores nulos permanecen vacíos.", fmt_wrap); row += 1
        wsh_ld.write(row, 0, "Versión", fmt_bold); wsh_ld.write(row, 1, "Indicadores de Tipo de Cambio Ver.3.0", fmt_wrap); row += 1
    
    except Exception:
        # No bloquear la generación del archivo si falla esta hoja
        pass
    # === Hoja "Metadatos" (crear siempre, al final) ===
    try:
        # Reutiliza formatos si existen; si no, crea mínimos
        try:
            fmt_all = fmt_all
        except NameError:
            fmt_all = wb.add_format({"font_name": "Arial"})
        try:
            fmt_bold = fmt_bold
        except NameError:
            fmt_bold = wb.add_format({"font_name": "Arial", "bold": True})
    
        # Crear hoja; si ya existe, usar nombre alterno
        try:
            wsm = wb.add_worksheet("Metadatos")
        except Exception:
            wsm = wb.add_worksheet("Meta datos")
    
        # Presentación básica
        try:
            wsm.set_column(0, 0, 28, fmt_all)
            wsm.set_column(1, 1, 48, fmt_all)
            wsm.hide_gridlines(2)
        except Exception:
            pass
    
        # Contenido
        from datetime import datetime
        ts = None
        try:
            ts = today_cdmx("%Y-%m-%d %H:%M (CDMX)")
        except Exception:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M")
        rows = [
            ("Generado", ts),
            ("Zona horaria", "America/Mexico_City"),
        ]
        # Series SIE si existe el dict
        try:
            rows.extend([
                ("SIE USD/MXN", SIE_SERIES.get("USD_FIX","")),
                ("SIE EUR/MXN", SIE_SERIES.get("EUR_MXN","")),
                ("SIE JPY/MXN", SIE_SERIES.get("JPY_MXN","")),
                ("SIE UDIS", SIE_SERIES.get("UDIS","")),
                ("SIE CETES 28", SIE_SERIES.get("CETES_28","")),
                ("SIE CETES 91", SIE_SERIES.get("CETES_91","")),
                ("SIE CETES 182", SIE_SERIES.get("CETES_182","")),
                ("SIE CETES 364", SIE_SERIES.get("CETES_364","")),
                ("SIE TIIE 28", SIE_SERIES.get("TIIE_28","")),
                ("SIE TIIE 91", SIE_SERIES.get("TIIE_91","")),
                ("SIE TIIE 182", SIE_SERIES.get("TIIE_182","")),
                ("SIE Tasa objetivo", SIE_SERIES.get("OBJETIVO","")),
            ])
        except Exception:
            pass
        for i, (k, v) in enumerate(rows):
            try:
                wsm.write(i, 0, k, fmt_bold)
            except Exception:
                wsm.write(i, 0, k)
            wsm.write(i, 1, v)
    except Exception:
        # No bloquear la generación del archivo si falla esta hoja
        pass
    wb.close()
    try:
        st.session_state['xlsx_bytes'] = bio.getvalue()
    except Exception:
        pass
except NameError:
    pass
except Exception:
    pass
    st.download_button(
    "Descargar Excel",
        data=bio.getvalue(),
        file_name=f"indicadores_{today_cdmx()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    



try:
    xbytes = st.session_state.get('xlsx_bytes')
    if xbytes:
        st.download_button(
            'Descargar Excel',
            data=xbytes,
            file_name=f"indicadores_{today_cdmx()}.xlsx",
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True
        )
except Exception:
    pass

    # Hoja Manual / Ayuda para usuarios
    try:
        wsh = wb.add_worksheet("Manual")
        wsh.set_column(0, 0, 28, fmt_all)
        wsh.set_column(1, 1, 90, fmt_all)
        wsh.hide_gridlines(2)
        wsh.write(0,0,"Sección", fmt_hdr); wsh.write(0,1,"Contenido", fmt_hdr)

        manual_rows = [
            ("Propósito", "Este archivo presenta indicadores de tipo de cambio, UDIS, TIIE y CETES para los últimos 6 días hábiles. Incluye tendencias (sparklines), metadatos y rangos con nombre para su integración en reportes."),
            ("Fechas", "Se usan días hábiles (lun-vie). Formato de fecha en cabecera: dd \"de\" mmm (ej.: 09 de sep)."),
            ("Fuentes", "Banxico SIE para FIX, EUR/MXN, JPY/MXN, UDIS, CETES (28/91/182/364) y TIIE (28/91/182) + SF61745 (tasa objetivo)."),
            ("Cálculos derivados", "USD/JPY = USD/MXN ÷ JPY/MXN; Euro/Dólar = EUR/MXN ÷ USD/MXN. UDIS/TIIE/CETES se muestran con relleno (ffill) cuando no hay publicación del día."),
            ("Relleno (ffill)", "Cuando el día hábil no tiene aún publicación, el valor se arrastra desde el último disponible. En la hoja Indicadores, los valores arrastrados se distinguen en itálicas color gris y con la leyenda *."),
            ("Sparklines", "Columna H muestra la tendencia de B..G para cada indicador principal."),
            ("Rangos con nombre", "RANGO_FECHAS, RANGO_USDMXN, RANGO_JPYMXN, RANGO_EURMXN, RANGO_UDIS, RANGO_TOBJ, RANGO_TIIE28, RANGO_TIIE91, RANGO_TIIE182, RANGO_C28, RANGO_C91, RANGO_C182, RANGO_C364."),
            ("Branding", "Se inserta logo.png (si existe) en la hoja Indicadores."),
            ("Trazabilidad", "Ver hoja Metadatos: zona horaria, reglas de negocio y claves SIE/FRED utilizadas."),
        ]
        for i,(k,v) in enumerate(manual_rows, start=1):
            wsh.write(i,0,k, fmt_bold); wsh.write(i,1,v, fmt_wrap)
    except Exception:
        pass
