# -*- coding: utf-8 -*-



import io
import re
import time
import html
import base64
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime
from pathlib import Path

import pytz
import requests
from requests.adapters import HTTPAdapter, Retry
import streamlit as st
import xlsxwriter  # Motor para Excel y gráficos

# Logo (favicon / encabezado / sidebar)
try:
    from PIL import Image
except Exception:
    Image = None

def load_logo():
    """Busca logo.png en ubicación estándar."""
    candidates = [
        Path(__file__).parent / "logo.png",
        Path("logo.png"),
        Path(__file__).parent / "assets" / "logo.png",
    ]
    for p in candidates:
        if p.exists():
            return p
    return None

def logo_image_or_emoji():
    """Devuelve objeto PIL.Image para favicon si es posible; si no, un emoji."""
    p = load_logo()
    if Image and p:
        try:
            return Image.open(p)
        except Exception:
            return "📈"
    return "📈"

def logo_base64(max_height_px=40):
    """Devuelve el logo en base64 para incrustarlo en HTML sticky."""
    p = load_logo()
    if not p:
        return None
    try:
        if Image:
            im = Image.open(p).convert("RGBA")
            
            if im.height > max_height_px:
                w = int(im.width * (max_height_px / im.height))
                im = im.resize((w, max_height_px))
            bio = io.BytesIO()
            im.save(bio, format="PNG")
            data = bio.getvalue()
        else:
            data = Path(p).read_bytes()
        return base64.b64encode(data).decode("ascii")
    except Exception:
        return None

BANXICO_TOKEN = "677aaedf11d11712aa2ccf73da4d77b6b785474eaeb2e092f6bad31b29de6609"
INEGI_TOKEN   = "0146a9ed-b70f-4ea2-8781-744b900c19d1"
FRED_TOKEN    = ""  

TZ_MX = pytz.timezone("America/Mexico_City")


st.set_page_config(
    page_title="Indicadores Económicos",
    page_icon=logo_image_or_emoji(),
    layout="centered"
)


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
  font-size: 1.35rem; line-height: 1.25; margin: 0;
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
            <h1>Indicadores (últimos 6 días) + Noticias</h1>
            <p>Indicadores Económicos.</p>
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )
else:
    
    st.title("📈 Indicadores (últimos 6 días) + Noticias")
    st.caption("Indicadores Económicos.")


if _logo_b64:
    st.sidebar.image(f"data:image/png;base64,{_logo_b64}", use_column_width=True)


SIE_SERIES = {
    "USD_FIX":   "SF43718",
    "EUR_FIX":   "SF46410",
    "JPY_FIX":   "SF46406",

    
    "TIIE_OBJ":  "",
    "TIIE_28":   "SF60653",
    "TIIE_91":   "",
    "TIIE_182":  "",

    
    "CETES_28":  "SF43936",
    "CETES_91":  "SF43939",
    "CETES_182": "SF43942",
    "CETES_364": "SF43945",

    "UDIS":      "SP68257",
}


def http_session(timeout=15):
    s = requests.Session()
    retries = Retry(total=3, backoff_factor=0.8,
                    status_forcelist=[429, 500, 502, 503, 504],
                    allowed_methods=frozenset(["GET", "POST"]))
    s.mount("https://", HTTPAdapter(max_retries=retries))
    s.request_timeout = timeout
    return s

def now_ts():
    return datetime.now(TZ_MX).strftime("%Y-%m-%d %H:%M:%S")

def today_cdmx():
    return datetime.now(TZ_MX).date()

def try_float(x):
    try:
        return float(str(x).replace(",", ""))
    except:
        return None

def parse_any_date(s: str):
    """Devuelve datetime naive (sin tz)."""
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(s), fmt)
        except:
            pass
    try:
        dt = parsedate_to_datetime(str(s))
        if dt.tzinfo is not None:
            dt = dt.astimezone(TZ_MX).replace(tzinfo=None)
        return dt
    except:
        return None

def fmt_date_str(dt: datetime | None):
    return dt.strftime("%Y-%m-%d") if isinstance(dt, datetime) else ""

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

def rolling_movex_for_last6(window:int=20):
    end = today_cdmx()
    start = end - timedelta(days=365)
    obs = sie_range(SIE_SERIES["USD_FIX"], start.isoformat(), end.isoformat())
    series = [try_float(o.get("dato")) for o in obs if try_float(o.get("dato")) is not None]
    if not series:
        return [None]*6
    out = []
    for k in range(6, 0, -1):
        idx = len(series) - k
        sub = series[max(0, idx-window+1): idx+1]
        out.append(sum(sub)/len(sub) if sub else None)
    return out


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

    s = http_session(timeout=20)
    last_err = None
    for url in urls:
        try:
            r = s.get(url, timeout=s.request_timeout)
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

            def get_v(o):
                if not o: return None
                return _num(o.get("OBS_VALUE") or o.get("OBS_VALUE_STR") or o.get("value"))

            fecha   = (d_obs or m_obs or a_obs or {}).get("TIME_PERIOD") or str(today_cdmx())
            diaria  = get_v(d_obs)
            mensual = get_v(m_obs)
            anual   = get_v(a_obs)

            if diaria is not None:
                if mensual is None: mensual = diaria * 30.4
                if anual   is None: anual   = mensual * 12

            return {"fecha": fecha, "diaria": diaria, "mensual": mensual, "anual": anual,
                    "_status": "OK", "_source": url}
        except Exception as e:
            last_err = f"{type(e).__name__}"

    return {"fecha": str(today_cdmx()), "diaria": None, "mensual": None, "anual": None,
            "_status": f"Fallo ({last_err})", "_source": urls[-1]}


@st.cache_data(ttl=60*30)
def fred_observations(series_id: str, start_date: str = None, end_date: str = None):
    base = "https://api.stlouisfed.org/fred/series/observations"
    params = {"series_id": series_id, "file_type": "json"}
    if FRED_TOKEN.strip(): params["api_key"] = FRED_TOKEN.strip()
    if start_date: params["observation_start"] = start_date
    if end_date:   params["observation_end"]   = end_date
    r = http_session(20).get(base, params=params, timeout=20)
    r.raise_for_status()
    return r.json().get("observations", [])

def fred_last_n(series_id: str, n: int = 12):
    try:
        if not FRED_TOKEN.strip():
            return []
        end = datetime.utcnow().date()
        start = (end - timedelta(days=5*365)).isoformat()
        obs = fred_observations(series_id, start_date=start, end_date=end.isoformat())
        out = [(o["date"], try_float(o["value"])) for o in obs if o.get("value") not in (".", None)]
        return out[-n:] if out else []
    except:
        return []

def fred_cpi_yoy_series(n: int = 12):
    try:
        if not FRED_TOKEN.strip():
            return []
        end = datetime.utcnow().date()
        start = (end - timedelta(days=6*365)).isoformat()
        obs = fred_observations("CPIAUCSL", start_date=start, end_date=end.isoformat())
        obs = [(o["date"], try_float(o["value"])) for o in obs if o.get("value") not in (".", None)]
        if len(obs) < 13: return []
        yoy = []
        for i in range(12, len(obs)):
            f_now, v_now = obs[i]
            f_prev, v_prev = obs[i-12]
            if v_now is None or not v_prev: continue
            yoy.append((f_now, (v_now / v_prev - 1) * 100.0))
        return yoy[-n:] if yoy else []
    except:
        return []


RSS_FEEDS = [
    "https://feeds.reuters.com/reuters/businessNews",
    "https://feeds.reuters.com/reuters/marketsNews",
    "https://finance.yahoo.com/news/rssindex",
]

def _strip_html(s: str) -> str:
    if not s: return ""
    s = html.unescape(s)
    s = re.sub(r"<[^>]+>", "", s)
    return s.replace("\xa0", " ").strip()

@st.cache_data(ttl=60*15)
def fetch_financial_news(limit_per_feed=8, total_limit=20):
    items = []
    s = http_session(15)
    for url in RSS_FEEDS:
        try:
            r = s.get(url, timeout=s.request_timeout); r.raise_for_status()
            from xml.etree import ElementTree as ET
            root = ET.fromstring(r.content)
            for item in root.findall(".//item")[:limit_per_feed]:
                title = _strip_html(item.findtext("title") or "")
                link  = (item.findtext("link") or "").strip()
                desc  = _strip_html(item.findtext("description") or "")
                pub   = item.findtext("pubDate") or ""
                dt    = parse_any_date(pub) or datetime.utcnow()
                source = re.sub(r"^https?://(www\.)?([^/]+)/?.*$", r"\2", link) if link else "rss"
                items.append({
                    "dt_str": dt.strftime("%Y-%m-%d %H:%M"),
                    "title": title, "link": link, "summary": desc, "source": source
                })
        except Exception:
            continue
    items.sort(key=lambda x: x["dt_str"], reverse=True)
    return items[:total_limit]


def _probe(fn, ok_pred):
    t0 = time.perf_counter()
    try:
        res = fn()
        status = ok_pred(res)
        msg = "OK" if status=="ok" else ("Parcial" if status=="warn" else "Sin datos")
    except Exception as e:
        status, msg = "err", f"Excepción: {type(e).__name__}"
    ms = int((time.perf_counter()-t0)*1000)
    return status, msg, ms

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
    with st.sidebar.expander("Herramientas"):
        c1, c2 = st.columns(2)
        if c1.button("Limpiar cachés Banxico"):
            sie_opportuno.clear(); sie_range.clear()
            st.success("Cachés SIE limpiadas.")
        if c2.button("Limpiar caché UMA"):
            get_uma.clear()
            st.success("Caché UMA limpiada.")
    with st.sidebar.expander("Diagnóstico UMA"):
        if st.button("Probar INEGI ahora"):
            res = get_uma(INEGI_TOKEN)
            st.write("Estado:", res.get("_status"), "— Fuente:", res.get("_source"))
            st.write("Diaria:", res.get("diaria"), "Mensual:", res.get("mensual"), "Anual:", res.get("anual"))


with st.expander("Opciones"):
    movex_win = st.number_input("Ventana MONEX (días hábiles)", min_value=5, max_value=60, value=20, step=1)
    margen_pct = st.number_input("Margen Compra/Venta sobre FIX (% por lado)", min_value=0.0, max_value=5.0, value=0.5, step=0.1)
    uma_manual = st.number_input("UMA diaria (manual, si INEGI falla)", min_value=0.0, value=0.0, step=0.01)
    do_charts = st.toggle("Agregar hoja 'Gráficos' (últimos 12)", value=True)
    do_raw    = st.toggle("Agregar hoja 'Datos crudos' (últimos 12)", value=True)

_check_tokens()
_render_sidebar_status()


if st.button("Generar Excel"):
    def pad6(lst): return ([None]*(6-len(lst)))+lst if len(lst)<6 else lst

    
    usd6_pairs = sie_last_n(SIE_SERIES["USD_FIX"], n=6)
    fechas6_dt  = pad6([parse_any_date(f) for f,_ in usd6_pairs])
    fechas6_str = [fmt_date_str(d) for d in fechas6_dt]
    usd6  = pad6([v for _, v in usd6_pairs])

    eur6  = pad6([v for _, v in sie_last_n(SIE_SERIES["EUR_FIX"], n=6)])
    jpy6  = pad6([v for _, v in sie_last_n(SIE_SERIES["JPY_FIX"], n=6)])
    udis6 = pad6([v for _, v in sie_last_n(SIE_SERIES["UDIS"],    n=6)])

    movex6  = rolling_movex_for_last6(window=int(movex_win))
    compra6 = [x*(1 - margen_pct/100.0) if x is not None else None for x in usd6]
    venta6  = [x*(1 + margen_pct/100.0) if x is not None else None for x in usd6]
    eurusd6 = [(e/u if (e is not None and u) else None) for e,u in zip(eur6, usd6)]
    usdjpy6 = [(u/j if (u is not None and j) else None) for u,j in zip(usd6, jpy6)]

    tiie28_6   = pad6([v for _, v in sie_last_n(SIE_SERIES["TIIE_28"], n=6)])
    cetes28_6  = pad6([v for _, v in sie_last_n(SIE_SERIES["CETES_28"], n=6)])
    cetes91_6  = pad6([v for _, v in sie_last_n(SIE_SERIES["CETES_91"], n=6)])
    cetes182_6 = pad6([v for _, v in sie_last_n(SIE_SERIES["CETES_182"], n=6)])
    cetes364_6 = pad6([v for _, v in sie_last_n(SIE_SERIES["CETES_364"], n=6)])

    none6 = [None]*6
    tiie_obj6 = none6 if not SIE_SERIES["TIIE_OBJ"] else pad6([v for _, v in sie_last_n(SIE_SERIES["TIIE_OBJ"], n=6)])
    tiie91_6  = none6 if not SIE_SERIES["TIIE_91"]  else pad6([v for _, v in sie_last_n(SIE_SERIES["TIIE_91"],  n=6)])
    tiie182_6 = none6 if not SIE_SERIES["TIIE_182"] else pad6([v for _, v in sie_last_n(SIE_SERIES["TIIE_182"], n=6)])

    
    uma = get_uma(INEGI_TOKEN)
    if uma.get("diaria") is None and uma_manual > 0:
        uma["diaria"]  = uma_manual
        uma["mensual"] = uma_manual * 30.4
        uma["anual"]   = uma["mensual"] * 12
        uma["_status"] = "MANUAL"

    
    news = fetch_financial_news(limit_per_feed=8, total_limit=20)

    
    bio = io.BytesIO()
    wb = xlsxwriter.Workbook(bio, {'in_memory': True})

    
    fmt_bold = wb.add_format({'bold': True})
    fmt_wrap = wb.add_format({'text_wrap': True, 'valign': 'top'})
    fmt_date = wb.add_format({'num_format': 'yyyy-mm-dd'})
    fmt_head = wb.add_format({'bold': True, 'bg_color': '#F2F2F2'})

    
    ws = wb.add_worksheet("Indicadores")
    ws.write("A2", "Fecha:", fmt_bold)
    for idx, s in enumerate(fechas6_str):
        ws.write(1, 1+idx, s, fmt_date)  

    
    labels = [
        (4, "TIPOS DE CAMBIO:", True),
        (6, "DÓLAR AMERICANO.", True),
        (7, "Dólar/Pesos:", False),
        (8, "MONEX:", False),
        (9, "Compra:", False),
        (10,"Venta:", False),

        (12,"YEN JAPONÉS.", True),
        (13,"Yen Japonés/Peso:", False),
        (14,"Dólar/Yen Japonés:", False),

        (16,"EURO.", True),
        (17,"Euro/Peso:", False),
        (18,"Euro/Dólar:", False),

        (20,"UDIS:", True),
        (22,"UDIS: ", False),

        (24,"TASAS TIIE:", True),
        (26,"TIIE objetivo:", False),
        (27,"TIIE 28 Días:", False),
        (28,"TIIE 91 Días:", False),
        (29,"TIIE 182 Días:", False),

        (31,"CETES:", True),
        (33,"CETES 28 Días:", False),
        (34,"CETES 91 Días:", False),
        (35,"Cetes 182 Días:", False),
        (36,"Cetes 364 Días:", False),

        (38,"UMA:", True),
        (40,"Diario:", False),
        (41,"Mensual:", False),
        (42,"Anual:", False),
    ]
    for row, text, bold in labels:
        ws.write(row-1, 0, text, fmt_bold if bold else None)

    def write_row_values(row_idx, values6):
        for j, v in enumerate(values6):
            if v is None:
                ws.write_blank(row_idx-1, 1+j, None)
            else:
                ws.write_number(row_idx-1, 1+j, v)

    uma_diaria6  = [uma["diaria"]]*6
    uma_mensual6 = [uma["mensual"]]*6
    uma_anual6   = [uma["anual"]]*6

    write_row_values(7,  usd6)
    write_row_values(8,  monex6)
    write_row_values(9,  compra6)
    write_row_values(10, venta6)

    write_row_values(13, jpy6)
    write_row_values(14, usdjpy6)

    write_row_values(17, eur6)
    write_row_values(18, eurusd6)

    write_row_values(22, udis6)

    write_row_values(26, tiie_obj6)
    write_row_values(27, tiie28_6)
    write_row_values(28, tiie91_6)
    write_row_values(29, tiie182_6)

    write_row_values(33, cetes28_6)
    write_row_values(34, cetes91_6)
    write_row_values(35, cetes182_6)
    write_row_values(36, cetes364_6)

    write_row_values(40, uma_diaria6)
    write_row_values(41, uma_mensual6)
    write_row_values(42, uma_anual6)

    ws.set_column(0, 0, 26)   
    ws.set_column(1, 6, 14)   

    
    ws2 = wb.add_worksheet("Noticias")
    ws2.write(0, 0, "Resumen de noticias financieras", fmt_bold)
    ws2.write(1, 0, f"Generado: {now_ts()} (CDMX)")
    headers = ["Fecha", "Fuente", "Título", "Resumen", "Link"]
    for col, h in enumerate(headers):
        ws2.write(3, col, h, fmt_head)
    rnews = 4
    if not news:
        ws2.write(rnews, 0, "Sin datos"); ws2.write(rnews, 2, "No se pudieron descargar noticias.")
    else:
        for it in news:
            ws2.write(rnews, 0, it["dt_str"])
            ws2.write(rnews, 1, it["source"])
            ws2.write(rnews, 2, it["title"])
            ws2.write(rnews, 3, (it["summary"][:400] + ("..." if len(it["summary"])>400 else "")), fmt_wrap)
            ws2.write(rnews, 4, it["link"])
            rnews += 1
    ws2.set_column(0, 0, 18); ws2.set_column(1, 1, 14); ws2.set_column(2, 2, 60); ws2.set_column(3, 3, 90); ws2.set_column(4, 4, 40)

    
    def add_table_and_chart(workbook, sheet, sheet_name, start_row, start_col, title, series_pairs, chart_anchor):
        """Escribe tabla 'Fecha, Valor' y grafica si hay >=2 puntos."""
        pairs = [(parse_any_date(f), v) for (f, v) in series_pairs if v is not None]
        pairs = [(p[0], p[1]) for p in pairs if p[0] is not None]
        if len(pairs) < 2:
            return start_row  
        
        sheet.write(start_row,   start_col, title, fmt_bold)
        sheet.write(start_row+1, start_col,   "Fecha", fmt_head)
        sheet.write(start_row+1, start_col+1, "Valor", fmt_head)
        r = start_row + 2
        for (dt, val) in pairs:
            sheet.write(r, start_col, fmt_date_str(dt), fmt_date)
            sheet.write_number(r, start_col+1, val)
            r += 1
        
        first = start_row + 2; last = r - 1
        chart = workbook.add_chart({'type': 'line'})
        chart.set_title({'name': title})
        chart.add_series({
            'categories': [sheet_name, first, start_col,   last, start_col],
            'values':     [sheet_name, first, start_col+1, last, start_col+1],
            'line':       {'width': 1.5},
        })
        chart.set_x_axis({'name': 'Fecha'})
        chart.set_y_axis({'name': 'Valor'})
        sheet.insert_chart(chart_anchor, chart)
        return r + 2  

    if do_charts or do_raw:
        ws3 = wb.add_worksheet("Gráficos")
        next_row = 0
        
        usd_last12  = sie_last_n(SIE_SERIES["USD_FIX"], n=12)
        tiie_last12 = sie_last_n(SIE_SERIES["TIIE_28"], n=12)
        fed_last12  = fred_last_n("FEDFUNDS", n=12)
        cpi_last12  = fred_cpi_yoy_series(n=12)

        if do_charts:
            next_row = add_table_and_chart(wb, ws3, "Gráficos", next_row, 0, "USD/MXN (FIX) - Últimos 12", usd_last12, "H2")
            next_row = add_table_and_chart(wb, ws3, "Gráficos", next_row, 0, "TIIE 28d (%) - Últimos 12", tiie_last12, "H18")
            if fed_last12:
                next_row = add_table_and_chart(wb, ws3, "Gráficos", next_row, 0, "Fed Funds (%) - Últimos 12", fed_last12, "H34")
            if cpi_last12:
                next_row = add_table_and_chart(wb, ws3, "Gráficos", next_row, 0, "Inflación EUA YoY (%) - Últimos 12", cpi_last12, "H50")
            ws3.set_column(0, 0, 12); ws3.set_column(1, 1, 14)

        if do_raw:
            ws4 = wb.add_worksheet("Datos crudos")
            ws4.write(0, 0, "Serie", fmt_head)
            ws4.write(0, 1, "Fecha", fmt_head)
            ws4.write(0, 2, "Valor", fmt_head)

            def dump_raw(sheet, start_row, name, series):
                r = start_row
                for (f, v) in series:
                    if v is None: continue
                    dt = parse_any_date(f)
                    sheet.write(r, 0, name)
                    sheet.write(r, 1, fmt_date_str(dt), fmt_date)
                    sheet.write_number(r, 2, v)
                    r += 1
                return r

            rraw = 1
            rraw = dump_raw(ws4, rraw, "USD/MXN (FIX)", usd_last12)
            rraw = dump_raw(ws4, rraw, "TIIE 28d (%)",  tiie_last12)
            if fed_last12: rraw = dump_raw(ws4, rraw, "Fed Funds (%)", fed_last12)
            if cpi_last12: rraw = dump_raw(ws4, rraw, "Inflación EUA YoY (%)", cpi_last12)
            ws4.set_column(0, 0, 22); ws4.set_column(1, 1, 12); ws4.set_column(2, 2, 12)

    
    wb.close()
    st.success("¡Listo! Archivo generado con indicadores.")
    st.download_button(
        "Descargar Excel",
        data=bio.getvalue(),
        file_name=f"indicadores_{today_cdmx()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
