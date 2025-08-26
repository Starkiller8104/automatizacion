
# app.py
# ──────────────────────────────────────────────────────────────────────────────
# Indicadores (6 puntos) + Noticias + (opc) Gráficos y Datos crudos
# Exportación con XlsxWriter y charts robustos. Fechas reales en B2..G2.
# CETES conectados, UMA robusto + fallback manual.
# Branding: logo favicon + encabezado sticky; menú y footer ocultos.
# ──────────────────────────────────────────────────────────────────────────────

import io
import re
import time
import html
import base64
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime
from pathlib import Path

import pytz
import re
import requests
import feedparser
from PIL import Image
from urllib.parse import urlparse
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
import xlsxwriter
from requests.adapters import HTTPAdapter, Retry

import streamlit as st

# ==== LOGIN (agregado) ====
import os, pytz as _pytz_for_login  # _pytz_for_login sólo para asegurar import si no existía

def _get_app_password() -> str:
    try:
        return st.secrets["APP_PASSWORD"]
    except Exception:
        pass
    if os.getenv("APP_PASSWORD"):
        return os.getenv("APP_PASSWORD")
    return "imemsa79"  # por defecto

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
# ==== /LOGIN ====


# =========================
#  Utilidades de tiempo/zonas
# =========================
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

# =========================
#  TOKENS
# =========================
BANXICO_TOKEN = "677aaedf11d11712aa2ccf73da4d77b6b785474eaeb2e092f6bad31b29de6609"
INEGI_TOKEN   = "0146a9ed-b70f-4ea2-8781-744b900c19d1"
FRED_TOKEN    = ""  # opcional para gráficos

TZ_MX = pytz.timezone("America/Mexico_City")

# ── Page config (debe ir antes de cualquier otro st.*)
st.set_page_config(
    page_title="Indicadores Económicos",
    page_icon=logo_image_or_emoji(),
    layout="centered"
)
_check_password()  # <<< Login requerido antes de mostrar la UI

# CSS: ocultar menú y footer + estilos del header sticky
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

# Encabezado sticky con logo
_logo_b64 = logo_base64()
if _logo_b64:
    st.markdown(
        f"""
        <div class="app-header">
          <img class="logo" src="data:image/png;base64,{_logo_b64}" alt="logo"/>
          <div class="titles">
            <h1>Indicadores (últimos 6 días) + Noticias</h1>
            <p>Excel con tu layout (B2..G2 fechas reales), noticias y gráficos con XlsxWriter.</p>
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )
else:
    # Fallback normal si no hay logo
    st.title("📈 Indicadores (últimos 6 días) + Noticias")
    st.caption("Excel con tu layout (B2..G2 fechas reales), noticias y gráficos con XlsxWriter.")

# Logo también en el sidebar (si existe)
if _logo_b64:
    st.sidebar.image(f"data:image/png;base64,{_logo_b64}", use_column_width=True)

# =========================
#  Helpers generales
# =========================
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

# =========================
#  Verificación de tokens
# =========================
def _check_tokens():
    missing = []
    if not BANXICO_TOKEN.strip(): missing.append("BANXICO_TOKEN")
    if not INEGI_TOKEN.strip():   missing.append("INEGI_TOKEN")
    if missing:
        st.error("Faltan tokens: " + ", ".join(missing))
        st.stop()

# =========================
#  Banxico SIE
# =========================
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
    start = end - timedelta(days=2*365)
    obs = sie_range("SF43718", start.isoformat(), end.isoformat())  # USD_FIX
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

# =========================
#  Series SIE / mapeo CETES, USD, etc.
# =========================
SIE_SERIES = {
    "USD_FIX":   "SF43718",
    "EUR_MXN":   "SF46410",
    "JPY_MXN":   "SF46406",
    "UDIS":      "SP68257",
    "TIie_dummy": None,  # placeholder si requieres

    "CETES_28":  "SF43936",
    "CETES_91":  "SF43939",
    "CETES_182": "SF43942",
    "CETES_364": "SF43945",
}

# =========================
#  INEGI UMA – robusto (con fallback y diagnóstico)
# =========================
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
        f"{base}/{ids}/es/00/true/BIE/2.0/{inegi_token}?type=json",  # fallback
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

# =========================
#  Render de estado en sidebar
# =========================
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

# =========================
#  STREAMLIT UI
# =========================
with st.expander("Opciones"):
    movex_win = st.number_input("Ventana MOVEX (días hábiles)", min_value=5, max_value=60, value=20, step=1)
    margen_pct = st.number_input("Margen Compra/Venta sobre FIX ...% por lado)", min_value=0.0, max_value=5.0, value=0.5, step=0.1)
    uma_manual = st.number_input("UMA diaria (manual, si INEGI falla)", min_value=0.0, value=0.0, step=0.01)
    do_charts = st.toggle("Agregar hoja 'Gráficos' (últimos 12)", value=True)
    do_raw    = st.toggle("Agregar hoja 'Datos crudos' (últimos 12)", value=True)

_check_tokens()
_render_sidebar_status()

# =========================
#  Generar Excel (XlsxWriter)
# =========================
if st.button("Generar Excel"):
    def pad6(lst): return ([None]*(6-len(lst)))+lst if len(lst) < 6 else lst[-6:]
    none6 = [None]*6

    # --- FIX USD/MXN (últimos 6)
    fix6 = pad6([v for _, v in sie_last_n(SIE_SERIES["USD_FIX"], n=6)])
    # EUR/MXN, JPY/MXN
    eur6 = pad6([v for _, v in sie_last_n(SIE_SERIES["EUR_MXN"], n=6)])
    jpy6 = pad6([v for _, v in sie_last_n(SIE_SERIES["JPY_MXN"], n=6)])

    # --- MOVEX rolling window
    movex_series = rolling_movex_for_last6(window=movex_win)
    movex6 = pad6(movex_series)

    # --- CETES últimos 6
    cetes28_6 = pad6([v for _, v in sie_last_n(SIE_SERIES["CETES_28"], n=6)])
    cetes91_6 = pad6([v for _, v in sie_last_n(SIE_SERIES["CETES_91"], n=6)])
    cetes182_6 = pad6([v for _, v in sie_last_n(SIE_SERIES["CETES_182"], n=6)])
    cetes364_6 = pad6([v for _, v in sie_last_n(SIE_SERIES["CETES_364"], n=6)])

    # --- UMA con fallback manual
    uma = get_uma(INEGI_TOKEN)
    if uma.get("diaria") is None and uma_manual > 0:
        uma["diaria"]  = uma_manual
        uma["mensual"] = uma_manual * 30.4
        uma["anual"]   = uma["mensual"] * 12

    # ===== Excel completo (reemplazado) =====
    bio = io.BytesIO()
    wb = xlsxwriter.Workbook(bio, {'in_memory': True})

    # ---- Preparar datos a escribir ----
    _fix_pairs = sie_last_n(SIE_SERIES["USD_FIX"], n=6)
    header_dates = [d for d,_ in _fix_pairs]
    if len(header_dates) < 6:
        header_dates = ([""]*(6-len(header_dates))) + header_dates

    def last6_map(key):
        pairs = sie_last_n(SIE_SERIES[key], n=6)
        return {d:v for d,v in pairs}

    m_fix   = last6_map("USD_FIX")
    m_eur   = last6_map("EUR_MXN")
    m_jpy   = last6_map("JPY_MXN")
    m_c28   = last6_map("CETES_28")
    m_c91   = last6_map("CETES_91")
    m_c182  = last6_map("CETES_182")
    m_c364  = last6_map("CETES_364")
    m_udis  = last6_map("UDIS")

    try:
        movex6  # maybe exists
    except NameError:
        _movex_series = rolling_movex_for_last6(window=movex_win)
        movex6 = _movex_series[-6:] if _movex_series else [None]*6

    def aligned_values(mdict):
        return [mdict.get(d) for d in header_dates]

    rowset = [
        ("USD/MXN (FIX)", aligned_values(m_fix)),
        (f"MOVEX {movex_win}", movex6),
        ("EUR/MXN", aligned_values(m_eur)),
        ("JPY/MXN", aligned_values(m_jpy)),
        ("CETES 28d (%)", aligned_values(m_c28)),
        ("CETES 91d (%)", aligned_values(m_c91)),
        ("CETES 182d (%)", aligned_values(m_c182)),
        ("CETES 364d (%)", aligned_values(m_c364)),
        ("UDIS", aligned_values(m_udis)),
    ]

    fmt_title   = wb.add_format({'bold': True, 'font_size': 14})
    fmt_hdr     = wb.add_format({'bold': True, 'align': 'center', 'bg_color': '#F2F2F2', 'border':1})
    fmt_row_lbl = wb.add_format({'bold': True, 'align': 'left'})
    fmt_num4    = wb.add_format({'num_format': '0.0000'})
    fmt_num6    = wb.add_format({'num_format': '0.000000'})
    fmt_pct2    = wb.add_format({'num_format': '0.00%'})
    fmt_wrap    = wb.add_format({'text_wrap': True})
    fmt_center  = wb.add_format({'align': 'center'})

    ws = wb.add_worksheet("Indicadores")
    ws.write(0, 0, "Indicadores (últimos 6 días)", fmt_title)
    ws.write(1, 0, "Fecha", fmt_hdr)
    for i, d in enumerate(header_dates):
        ws.write(1, 1+i, d, fmt_hdr)

    start_row = 2
    for r, (label, values) in enumerate(rowset):
        ws.write(start_row+r, 0, label, fmt_row_lbl)
        for c, val in enumerate(values):
            if "CETES" in label:
                ws.write(start_row+r, 1+c, (val/100 if val is not None else None), fmt_pct2)
            elif label == "UDIS":
                ws.write(start_row+r, 1+c, val, fmt_num6)
            else:
                ws.write(start_row+r, 1+c, val, fmt_num4)

    # UMA
    ws.write(1, 9, "UMA", fmt_hdr)
    ws.write(2, 9, "Fecha", fmt_row_lbl);  ws.write(2,10, uma.get("fecha"))
    ws.write(3, 9, "Diaria", fmt_row_lbl); ws.write(3,10, uma.get("diaria"))
    ws.write(4, 9, "Mensual", fmt_row_lbl);ws.write(4,10, uma.get("mensual"))
    ws.write(5, 9, "Anual", fmt_row_lbl);  ws.write(5,10, uma.get("anual"))
    ws.set_column(0, 0, 22); ws.set_column(1, 6, 14); ws.set_column(9, 10, 14)

    # Datos crudos (opcional)
    try:
        do_raw
    except NameError:
        do_raw = True
    if do_raw:
        ws3 = wb.add_worksheet("Datos crudos")
        ws3.write(0,0,"Serie", fmt_hdr); ws3.write(0,1,"Fecha", fmt_hdr); ws3.write(0,2,"Valor", fmt_hdr)
        row = 1
        def dump_raw(tag, mdict, row_idx):
            for d in header_dates:
                v = mdict.get(d)
                ws3.write(row_idx, 0, tag)
                ws3.write(row_idx, 1, d, fmt_center)
                ws3.write(row_idx, 2, v, fmt_num6)
                row_idx += 1
            return row_idx
        row = dump_raw("USD/MXN (FIX)", m_fix, row)
        row = dump_raw("EUR/MXN",       m_eur, row)
        row = dump_raw("JPY/MXN",       m_jpy, row)
        row = dump_raw("CETES 28d (%)", m_c28, row)
        row = dump_raw("CETES 91d (%)", m_c91, row)
        row = dump_raw("CETES 182d (%)",m_c182, row)
        row = dump_raw("CETES 364d (%)",m_c364, row)
        row = dump_raw("UDIS",          m_udis, row)
        ws3.set_column(0, 0, 18); ws3.set_column(1, 1, 12); ws3.set_column(2, 2, 16)

    # Gráficos (opcional)
    try:
        do_charts
    except NameError:
        do_charts = True
    if do_charts:
        ws4 = wb.add_worksheet("Gráficos")
        chart1 = wb.add_chart({'type': 'line'})
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

        chart2 = wb.add_chart({'type': 'line'})
        for off in [4,5,6,7]:
            r = start_row + off
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
    st.success("¡Listo! Archivo generado con datos y hojas completas.")
    st.download_button(
        "⬇️ Descargar Excel",
        data=bio.getvalue(),
        file_name=f"indicadores_{today_cdmx()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
