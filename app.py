

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
    fmt_bold = wb.add_format({"bold": True, "align": "center"})
    fmt_date = wb.add_format({"num_format": "yyyy-mm-dd"})
    fmt_num  = wb.add_format({"num_format": "#,##0.0000"})
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
    fmt_bold = wb.add_format({"bold": True})
    fmt_link = wb.add_format({"font_color": "blue", "underline": 1})
    fmt_date = wb.add_format({"num_format": "yyyy-mm-dd hh:mm"})
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
    want_fred   = st.checkbox("Agregar hoja FRED", value=st.session_state.get("want_fred", True))
    want_news   = st.checkbox("Agregar hoja Noticias_RSS", value=st.session_state.get("want_news", True))
    want_charts = st.checkbox("Agregar hoja 'Gráficos' (últimos 12)", value=st.session_state.get("want_charts", True))
    want_raw    = st.checkbox("Agregar hoja 'Datos crudos' (últimos 12)", value=st.session_state.get("want_raw", True))
    st.session_state["want_fred"] = want_fred
    st.session_state["want_news"] = want_news
    st.session_state["want_charts"] = want_charts
    st.session_state["want_raw"] = want_raw

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


    fmt_bold  = wb.add_format({'bold': True})
    fmt_hdr   = wb.add_format({'bold': True, 'bg_color': '#F2F2F2', 'align':'center'})
    fmt_num4  = wb.add_format({'num_format': '0.0000'})
    fmt_num6  = wb.add_format({'num_format': '0.000000'})
    fmt_wrap  = wb.add_format({'text_wrap': True})

    
    _fix_pairs = sie_last_n(SIE_SERIES["USD_FIX"], n=6)
    header_dates = [d for d,_ in _fix_pairs]
    if len(header_dates) < 6:
        header_dates = ([""]*(6-len(header_dates))) + header_dates

    def _as_map(pairs): return {d:v for d,v in pairs}
    m_fix  = _as_map(sie_last_n(SIE_SERIES["USD_FIX"], 6))
    m_eur  = _as_map(sie_last_n(SIE_SERIES["EUR_MXN"], 6))
    m_jpy  = _as_map(sie_last_n(SIE_SERIES["JPY_MXN"], 6))
    m_udis = _as_map(sie_last_n(SIE_SERIES["UDIS"],    6))
    m_c28  = _as_map(sie_last_n(SIE_SERIES["CETES_28"],6))
    m_c91  = _as_map(sie_last_n(SIE_SERIES["CETES_91"],6))
    m_c182 = _as_map(sie_last_n(SIE_SERIES["CETES_182"],6))
    m_c364 = _as_map(sie_last_n(SIE_SERIES["CETES_364"],6))
    cetes28 = _ffill_by_dates(m_c28, header_dates)
    cetes91 = _ffill_by_dates(m_c91, header_dates)
    cetes182 = _ffill_by_dates(m_c182, header_dates)
    cetes364 = _ffill_by_dates(m_c364, header_dates)


    try:
        movex6  
    except NameError:
        movex6 = rolling_movex_for_last6(window=movex_win)
    compra = [(x*(1 - margen_pct/100) if x is not None else None) for x in movex6]
    venta  = [(x*(1 + margen_pct/100) if x is not None else None) for x in movex6]

    usd_jpy = []
    eur_usd = []
    for d in header_dates:
        u = m_fix.get(d); j = m_jpy.get(d); e = m_eur.get(d)
        usd_jpy.append((u/j) if (u and j) else None)
        eur_usd.append((e/u) if (e and u) else None)

    try:
        uma  
    except NameError:
        uma = get_uma(INEGI_TOKEN)

    def _last_or_none(series_pairs): 
        return series_pairs[-1][1] if series_pairs else None

        
    # TIIE (Banxico SIE) - usar histórico y alinear por fechas del encabezado
    m_t28  = _as_map(sie_last_n(SIE_SERIES["TIIE_28"],  20))
    m_t91  = _as_map(sie_last_n(SIE_SERIES["TIIE_91"],  20))
    m_t182 = _as_map(sie_last_n(SIE_SERIES["TIIE_182"], 20))
    # Objetivo (tasa de política monetaria)
    m_obj  = _as_map(sie_last_n("SF61745", 20))

    tiie28 = _ffill_by_dates(m_t28,  header_dates)
    tiie91 = _ffill_by_dates(m_t91,  header_dates)
    tiie182= _ffill_by_dates(m_t182, header_dates)
    tiie_obj = _ffill_by_dates(m_obj, header_dates)

    ws = wb.add_worksheet("Indicadores")
    ws.set_column(0, 6, 16)

    ws.write(1, 0, "Fecha:", fmt_bold)
    for i, d in enumerate(header_dates):
        ws.write(1, 1+i, d)

    ws.write(3, 0, "TIPOS DE CAMBIO:", fmt_bold)
    ws.write(5, 0, "DÓLAR AMERICANO.", fmt_bold)
    ws.write(6, 0, "Dólar/Pesos:")
    for i, d in enumerate(header_dates):
        ws.write(6, 1+i, m_fix.get(d), fmt_num4)
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
    for i, d in enumerate(header_dates):
        ws.write(12, 1+i, m_jpy.get(d), fmt_num6)
    ws.write(13, 0, "Dólar/Yen Japonés:")
    for i, v in enumerate(usd_jpy):
        ws.write(13, 1+i, v, fmt_num6)

    ws.write(15, 0, "EURO.", fmt_bold)
    ws.write(16, 0, "Euro/Peso:")
    for i, d in enumerate(header_dates):
        ws.write(16, 1+i, m_eur.get(d), fmt_num6)
    ws.write(17, 0, "Euro/Dólar:")
    for i, v in enumerate(eur_usd):
        ws.write(17, 1+i, v, fmt_num6)

    ws.write(19, 0, "UDIS:", fmt_bold)
    ws.write(21, 0, "UDIS: ")
    for i, d in enumerate(header_dates):
        ws.write(21, 1+i, m_udis.get(d), fmt_num6)

    ws.write(23, 0, "TASAS TIIE:", fmt_bold)
    ws.write(25, 0, "TIIE objetivo:")
    ws.write(26, 0, "TIIE 28 Días:")
    ws.write(27, 0, "TIIE 91 Días:")
    ws.write(28, 0, "TIIE 182 Días:")
    for i in range(6):
        ws.write(25, 1+i, tiie_obj[i])
        ws.write(26, 1+i, tiie28[i])
        ws.write(27, 1+i, tiie91[i])
        ws.write(28, 1+i, tiie182[i])
    ws.write(30, 0, "CETES:", fmt_bold)
    ws.write(32, 0, "CETES 28 Días:")
    ws.write(33, 0, "CETES 91 Días:")
    ws.write(34, 0, "CETES 182 Días:")
    ws.write(35, 0, "CETES 364 Días:")
    for i in range(6):
        ws.write(32, 1+i, cetes28[i])
        ws.write(33, 1+i, cetes91[i])
        ws.write(34, 1+i, cetes182[i])
        ws.write(35, 1+i, cetes364[i])
    ws.write(37, 0, "UMA:", fmt_bold)
    ws.write(39, 0, "Diario:");  ws.write(39, 1, uma.get("diaria"))
    ws.write(40, 0, "Mensual:"); ws.write(40, 1, uma.get("mensual"))
    ws.write(41, 0, "Anual:");   ws.write(41, 1, uma.get("anual"))

do_raw = globals().get('do_raw', True)
if do_raw and ('wb' in globals()):
    ws3 = wb.add_worksheet("Datos crudos")
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
        if fred_rows:
            wsname  = f"FRED_{fred_id[:25]}"
            wsfred  = wb.add_worksheet(wsname)

            fmt_bold = wb.add_format({"bold": True})
            fmt_num  = wb.add_format({"num_format": "#,##0.0000"})
            fmt_date = wb.add_format({"num_format": "yyyy-mm-dd"})

            
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
    if fred_key and st.session_state.get('want_fred', True):
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
    if _news and st.session_state.get('want_news', True):
        _mx_news_write_v1(wb, _news, sheet_name="Noticias_RSS")
except Exception:
    pass

try:
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
