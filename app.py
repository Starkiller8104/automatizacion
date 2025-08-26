
# -*- coding: utf-8 -*-
"""
IMEMSA – Indicadores + Excel con layout fijo
- Mantiene login (contraseña por defecto: imemsa79; override por st.secrets["APP_PASSWORD"] o env APP_PASSWORD)
- Mantiene captura manual de tokens en el sidebar (BANXICO_TOKEN, INEGI_TOKEN)
- Excel respeta el layout del archivo de referencia:
  Hoja "Indicadores":
    A2="Fecha:", B2..G2 = fechas (YYYY-MM-DD)
    Secciones:
      "TIPOS DE CAMBIO"
      DÓLAR AMERICANO:  B7..G7 USD/MXN (FIX), B8..G8 MONEX, B9..G9 Compra, B10..G10 Venta
      YEN JAPONÉS:      B13..G13 JPY/MXN, B14..G14 USD/JPY
      EURO:             B17..G17 EUR/MXN, B18..G18 EUR/USD
      UDIS:             B22..G22 UDIS
      TIIE:             B27..G27 TIIE 28, B28..G28 TIIE 91, B29..G29 TIIE 182  (repetidas a 6 columnas con el último valor disponible)
      CETES:            B33..G33 28d, B34..G34 91d, B35..G35 182d, B36..G36 364d
      UMA:              B40 Diario, B41 Mensual, B42 Anual
  Hojas "Noticias", "Datos crudos" y "Gráficos" opcionales
"""

import io
import os
from datetime import datetime, timedelta
from pathlib import Path

import requests
from requests.adapters import HTTPAdapter, Retry
import feedparser
import pytz
import streamlit as st
import xlsxwriter

# ============= Configuración básica + Login =================
st.set_page_config(page_title="Indicadores Económicos", page_icon="📊", layout="centered")
CDMX = pytz.timezone("America/Mexico_City")

def _get_app_password() -> str:
    try:
        return st.secrets["APP_PASSWORD"]
    except Exception:
        pass
    if os.getenv("APP_PASSWORD"):
        return os.getenv("APP_PASSWORD")
    return "imemsa79"

def _check_password():
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

_check_password()

# ============= Tokens (valores por defecto, sobre-escribibles en sidebar) ============
BANXICO_TOKEN = ""   # si no se captura, no consulta
INEGI_TOKEN   = ""   # si no se captura, UMA intentará fallback

# ============= Utilidades HTTP y helpers =================
def http_session(timeout=15):
    s = requests.Session()
    retries = Retry(total=3, backoff_factor=0.7,
                    status_forcelist=[429, 500, 502, 503, 504],
                    allowed_methods=frozenset(["GET"]))
    s.mount("https://", HTTPAdapter(max_retries=retries))
    s.mount("http://", HTTPAdapter(max_retries=retries))
    orig = s.request
    def _req(method, url, **kw):
        kw.setdefault("timeout", timeout)
        return orig(method, url, **kw)
    s.request = _req
    return s

def today_cdmx():
    return datetime.now(CDMX).date()

def try_float(x):
    try: return float(str(x).replace(",", "").strip())
    except: return None

# ================= BANXICO (SIE) =================
SIE_SERIES = {
    "USD_FIX":   "SF43718",
    "EUR_MXN":   "SF46410",
    "JPY_MXN":   "SF46406",
    "UDIS":      "SP68257",
    "CETES_28":  "SF43936",
    "CETES_91":  "SF43939",
    "CETES_182": "SF43942",
    "CETES_364": "SF43945",
}

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
    """Devuelve últimos n (fecha ISO, valor float)."""
    end = today_cdmx()
    start = end - timedelta(days=2*365)
    obs = sie_range(series_id, start.isoformat(), end.isoformat(), token)
    vals = []
    for o in obs:
        f = o.get("fecha"); v = try_float(o.get("dato"))
        if f and (v is not None):
            vals.append((f, v))
    vals.sort(key=lambda x: x[0])
    return vals[-n:]

def rolling_movex_for_last6(window:int, token:str):
    """Promedio móvil (window) sobre FIX; devuelve últimos 6 valores de la serie movex."""
    end = today_cdmx()
    start = end - timedelta(days=2*365)
    obs = sie_range(SIE_SERIES["USD_FIX"], start.isoformat(), end.isoformat(), token)
    seq = [try_float(o.get("dato")) for o in obs if try_float(o.get("dato")) is not None]
    if not seq:
        return [None]*6
    movex = []
    for i in range(len(seq)):
        sub = seq[max(0, i-window+1): i+1]
        movex.append(sum(sub)/len(sub))
    return movex[-6:] if len(movex) >= 6 else [None]*6

# ================= INEGI (UMA) =================
@st.cache_data(ttl=60*60, show_spinner=False)
def get_uma(inegi_token: str):
    """Regresa {'fecha','diaria','mensual','anual'} o None si falla."""
    base = "https://www.inegi.org.mx/app/api/indicadores/desarrolladores/jsonxml/INDICATOR"
    ids = "620706,620707,620708"  # diaria, mensual, anual
    for catalog in ("BISE", "BIE"):  # intenta 2 catálogos
        try:
            url = f"{base}/{ids}/es/00/true/{catalog}/2.0/{inegi_token}?type=json"
            r = http_session(20).get(url)
            if r.status_code != 200:
                continue
            data = r.json()
            series = data.get("Series") or data.get("series") or []
            if not series:
                continue
            def last_obs(s):
                obs = s.get("OBSERVATIONS") or s.get("observations") or []
                return obs[-1] if obs else None
            d = last_obs(series[0]); m = last_obs(series[1]); a = last_obs(series[2])
            def _v(o): return try_float(o.get("OBS_VALUE") or o.get("value")) if o else None
            def _f(o): return (o.get("TIME_PERIOD") or o.get("periodo") or o.get("time_period") or o.get("fecha")) if o else None
            return {"fecha": _f(d) or _f(m) or _f(a), "diaria": _v(d), "mensual": _v(m), "anual": _v(a)}
        except Exception:
            continue
    return {"fecha": None, "diaria": None, "mensual": None, "anual": None}

# ================= TIIE (último valor) =================
def fetch_tiie_dof():
    """Intenta obtener TIIE 28/91/182 del DOF/Banxico. Si falla, regresa None."""
    try:
        # Simple: tomamos oportuno de SIE si existen estas series (si no, usa DOF scrape; aquí mantenemos simple).
        # Series SIE oficiales: SF43783 (TIIE 28), SF43784 (TIIE 91), SF43785 (TIIE 182) – en %.
        ids = {"tiie28":"SF43783", "tiie91":"SF43784", "tiie182":"SF43785"}
        end = today_cdmx(); start = end - timedelta(days=400)
        out = {}
        for k, sid in ids.items():
            data = sie_range(sid, start.isoformat(), end.isoformat(), BANXICO_TOKEN)
            if data:
                v = try_float(data[-1].get("dato"))
                out[k] = v
        if out:
            return out
    except Exception:
        pass
    return {"tiie28": None, "tiie91": None, "tiie182": None}

# ================= Noticias (para hoja "Noticias") =================
def build_news_bullets(n=12):
    feeds = [
        "https://www.reuters.com/markets/americas/mexico/feed/?rpc=401&",
        "https://www.eleconomista.com.mx/rss/economia",
        "https://www.elfinanciero.com.mx/rss/finanzas/",
        "https://www.bloomberglinea.com/mexico/rss/",
    ]
    keys = ["México","Banxico","inflación","tasa","CETES","dólar","tipo de cambio","Pemex","Fed","FOMC","nearshoring"]
    rows = []
    for u in feeds:
        try:
            fp = feedparser.parse(u)
            for e in fp.entries[:40]:
                title = (e.get("title","") or "").strip()
                summary = (e.get("summary","") or "")
                link = (e.get("link","") or "").strip()
                text = (title + " " + summary).lower()
                if any(k.lower() in text for k in keys):
                    rows.append((e.get("published",""), f"• {title} — {link}"))
        except Exception:
            pass
    rows.sort(reverse=True, key=lambda x: x[0])
    return "\n".join([t for _, t in rows[:n]]) or "Sin novedades."

# ================= Sidebar: captura manual de tokens =================
st.sidebar.header("🔑 Tokens de APIs")
ban_t = st.sidebar.text_input("BANXICO_TOKEN", value="", type="password", help="Se usará en esta sesión.")
ine_t = st.sidebar.text_input("INEGI_TOKEN", value="", type="password", help="Se usará en esta sesión.")
if ban_t.strip(): BANXICO_TOKEN = ban_t.strip()
if ine_t.strip(): INEGI_TOKEN   = ine_t.strip()

# ================= Panel de opciones (igual que antes) =================
with st.expander("Opciones"):
    movex_win   = st.number_input("Ventana MONEX (días hábiles)", min_value=5, max_value=60, value=20, step=1)
    margen_pct  = st.number_input("Margen en Compra/Venta sobre FIX (% por lado)", min_value=0.0, max_value=5.0, value=0.50, step=0.10)
    uma_manual  = st.number_input("UMA diaria (manual, si INEGI falla)", min_value=0.0, value=0.0, step=0.01)
    do_charts   = st.toggle("Agregar hoja 'Gráficos' (últimos 12)", value=True)
    do_raw      = st.toggle("Agregar hoja 'Datos crudos' (últimos 12)", value=True)

st.title("Indicadores Económicos.  📰 + Excel (layout IMEMSA).")

# ================= Generador de Excel (solo se cambia esta parte) =================
if st.button("Generar Excel"):
    if not BANXICO_TOKEN.strip():
        st.error("Falta BANXICO_TOKEN (ingrésalo en la barra lateral)."); st.stop()

    # --- Series base
    fix6  = sie_last_n(SIE_SERIES["USD_FIX"], 6, BANXICO_TOKEN)    # [(fechaISO, valor)]
    eur6  = sie_last_n(SIE_SERIES["EUR_MXN"], 6, BANXICO_TOKEN)
    jpy6  = sie_last_n(SIE_SERIES["JPY_MXN"], 6, BANXICO_TOKEN)
    udis6 = sie_last_n(SIE_SERIES["UDIS"],    6, BANXICO_TOKEN)
    c28_6 = sie_last_n(SIE_SERIES["CETES_28"],6, BANXICO_TOKEN)
    c91_6 = sie_last_n(SIE_SERIES["CETES_91"],6, BANXICO_TOKEN)
    c182_6= sie_last_n(SIE_SERIES["CETES_182"],6, BANXICO_TOKEN)
    c364_6= sie_last_n(SIE_SERIES["CETES_364"],6, BANXICO_TOKEN)

    # Fechas encabezado: de FIX
    header_dates = [d for d,_ in fix6]
    if len(header_dates) < 6:
        header_dates = ([""]*(6-len(header_dates))) + header_dates

    # Mapas por fecha
    def as_map(pairs): return {d:v for d,v in pairs}
    m_fix, m_eur, m_jpy, m_udis = as_map(fix6), as_map(eur6), as_map(jpy6), as_map(udis6)
    m_c28, m_c91, m_c182, m_c364 = as_map(c28_6), as_map(c91_6), as_map(c182_6), as_map(c364_6)

    # MONEX (promedio móvil)
    movex6 = rolling_movex_for_last6(window=movex_win, token=BANXICO_TOKEN)
    # Compra/Venta con margen
    compra = [(x*(1 - margen_pct/100) if x is not None else None) for x in movex6]
    venta  = [(x*(1 + margen_pct/100) if x is not None else None) for x in movex6]

    # USD/JPY y EUR/USD derivados
    usd_jpy = []
    eur_usd = []
    for d in header_dates:
        u = m_fix.get(d); j = m_jpy.get(d); e = m_eur.get(d)
        usd_jpy.append((u/j) if (u and j) else None)
        eur_usd.append((e/u) if (e and u) else None)

    # UMA
    uma = get_uma(INEGI_TOKEN)
    if (uma.get("diaria") is None) and uma_manual>0:
        uma["diaria"] = uma_manual
        uma["mensual"] = uma_manual * 30.4
        uma["anual"] = (uma["mensual"] or 0) * 12

    # TIIE (último valor repetido 6 veces; si no hay, deja en blanco)
    tiie = fetch_tiie_dof()
    tiie28 = [tiie.get("tiie28")] * 6
    tiie91 = [tiie.get("tiie91")] * 6
    tiie182= [tiie.get("tiie182")] * 6

    # ================= Armado del Excel con EL MISMO LAYOUT =================
    bio = io.BytesIO()
    wb = xlsxwriter.Workbook(bio, {'in_memory': True})

    # formatos simples
    fmt_bold  = wb.add_format({'bold': True})
    fmt_hdr   = wb.add_format({'bold': True, 'bg_color': '#F2F2F2', 'align':'center'})
    fmt_num4  = wb.add_format({'num_format': '0.0000'})
    fmt_num6  = wb.add_format({'num_format': '0.000000'})
    fmt_pct2  = wb.add_format({'num_format': '0.00%'})
    fmt_wrap  = wb.add_format({'text_wrap': True})

    ws = wb.add_worksheet("Indicadores")

    # A2 "Fecha:" y B2..G2 fechas ISO
    ws.write(1, 0, "Fecha:", fmt_bold)
    for i, d in enumerate(header_dates):
        ws.write(1, 1+i, d)

    # ---- TIPO DE CAMBIO (misma estructura del layout que adjuntaste) ----
    ws.write(3, 0, "TIPOS DE CAMBIO:", fmt_bold)
    ws.write(5, 0, "DÓLAR AMERICANO.", fmt_bold)

    # Fila 7: Dólar/Pesos (USD/MXN FIX)
    ws.write(6, 0, "Dólar/Pesos:")
    for i, d in enumerate(header_dates):
        ws.write(6, 1+i, m_fix.get(d), fmt_num4)

    # Fila 8: MONEX (promedio móvil)
    ws.write(7, 0, "MONEX:")
    for i, val in enumerate(movex6):
        ws.write(7, 1+i, val, fmt_num6)

    # Fila 9-10: Compra/Venta con margen
    ws.write(8, 0, "Compra:")
    for i, val in enumerate(compra):
        ws.write(8, 1+i, val, fmt_num6)
    ws.write(9, 0, "Venta:")
    for i, val in enumerate(venta):
        ws.write(9, 1+i, val, fmt_num6)

    # ---- YEN ----
    ws.write(11, 0, "YEN JAPONÉS.", fmt_bold)
    ws.write(12, 0, "Yen Japonés/Peso:")
    for i, d in enumerate(header_dates):
        ws.write(12, 1+i, m_jpy.get(d), fmt_num6)
    ws.write(13, 0, "Dólar/Yen Japonés:")
    for i, v in enumerate(usd_jpy):
        ws.write(13, 1+i, v, fmt_num6)

    # ---- EURO ----
    ws.write(15, 0, "EURO.", fmt_bold)
    ws.write(16, 0, "Euro/Peso:")
    for i, d in enumerate(header_dates):
        ws.write(16, 1+i, m_eur.get(d), fmt_num6)
    ws.write(17, 0, "Euro/Dólar:")
    for i, v in enumerate(eur_usd):
        ws.write(17, 1+i, v, fmt_num6)

    # ---- UDIS ----
    ws.write(19, 0, "UDIS:", fmt_bold)
    ws.write(21, 0, "UDIS: ")
    for i, d in enumerate(header_dates):
        ws.write(21, 1+i, m_udis.get(d), fmt_num6)

    # ---- TIIE ----
    ws.write(23, 0, "TASAS TIIE:", fmt_bold)
    ws.write(25, 0, "TIIE objetivo:")  # (si quieres, aquí puedes escribir el objetivo más adelante)
    ws.write(26, 0, "TIIE 28 Días:")
    ws.write(27, 0, "TIIE 91 Días:")
    ws.write(28, 0, "TIIE 182 Días:")
    for i in range(6):
        ws.write(26, 1+i, tiie28[i] if tiie28[i] is not None else None)
        ws.write(27, 1+i, tiie91[i] if tiie91[i] is not None else None)
        ws.write(28, 1+i, tiie182[i] if tiie182[i] is not None else None)

    # ---- CETES (en % como en tu layout, no en fracción) ----
    ws.write(30, 0, "CETES:", fmt_bold)
    ws.write(32, 0, "CETES 28 Días:")
    ws.write(33, 0, "CETES 91 Días:")
    ws.write(34, 0, "Cetes 182 Días:")
    ws.write(35, 0, "Cetes 364 Días:")
    for i, d in enumerate(header_dates):
        ws.write(32, 1+i, m_c28.get(d))   # valores como 7.65 (no 0.0765)
        ws.write(33, 1+i, m_c91.get(d))
        ws.write(34, 1+i, m_c182.get(d))
        ws.write(35, 1+i, m_c364.get(d))

    # ---- UMA (valores en B40..B42) ----
    ws.write(37, 0, "UMA:", fmt_bold)
    ws.write(39, 0, "Diario:");  ws.write(39, 1, uma.get("diaria"))
    ws.write(40, 0, "Mensual:"); ws.write(40, 1, uma.get("mensual"))
    ws.write(41, 0, "Anual:");   ws.write(41, 1, uma.get("anual"))

    # ---------- Hoja Noticias ----------
    wsN = wb.add_worksheet("Noticias")
    wsN.write(0, 0, "Noticias financieras recientes", fmt_bold)
    wsN.write(1, 0, build_news_bullets(12), fmt_wrap)
    wsN.set_column(0, 0, 120)

    # ---------- Hoja Datos crudos (opcional) ----------
    if do_raw:
        wsR = wb.add_worksheet("Datos crudos")
        wsR.write(0,0,"Serie", fmt_hdr); wsR.write(0,1,"Fecha", fmt_hdr); wsR.write(0,2,"Valor", fmt_hdr)
        row = 1
        def dump_raw(tag, pairs):
            nonlocal row
            for d, v in pairs:
                wsR.write(row, 0, tag)
                wsR.write(row, 1, d)
                wsR.write(row, 2, v)
                row += 1
        dump_raw("USD/MXN (FIX)", fix6)
        dump_raw("EUR/MXN",       eur6)
        dump_raw("JPY/MXN",       jpy6)
        dump_raw("UDIS",          udis6)
        dump_raw("CETES 28d (%)", c28_6)
        dump_raw("CETES 91d (%)", c91_6)
        dump_raw("CETES 182d (%)",c182_6)
        dump_raw("CETES 364d (%)",c364_6)
        wsR.set_column(0, 0, 18); wsR.set_column(1, 1, 12); wsR.set_column(2, 2, 14)

    # ---------- Hoja Gráficos (opcional, simple) ----------
    if do_charts:
        wsG = wb.add_worksheet("Gráficos")
        # Línea USD/MXN
        chart1 = wb.add_chart({'type': 'line'})
        chart1.add_series({
            'name':       "USD/MXN (FIX)",
            'categories': "=Indicadores!$B$2:$G$2",
            'values':     "=Indicadores!$B$7:$G$7",
        })
        chart1.set_title({'name': 'USD/MXN (FIX)'})
        wsG.insert_chart('B2', chart1, {'x_scale': 1.3, 'y_scale': 1.2})

        # Línea CETES
        chart2 = wb.add_chart({'type': 'line'})
        for row in (33,34,35,36):  # filas 1-based en tu layout (CETES)
            chart2.add_series({
                'name':       f"=Indicadores!$A${row+1}",
                'categories': "=Indicadores!$B$2:$G$2",
                'values':     f"=Indicadores!$B${row+1}:$G${row+1}",
            })
        chart2.set_title({'name': 'CETES (%)'})
        wsG.insert_chart('B18', chart2, {'x_scale': 1.3, 'y_scale': 1.2})

    # Cerrar y servir
    wb.close()
    bio.seek(0)
    st.success("¡Listo! Excel generado con el layout IMEMSA.")
    st.download_button("⬇️ Descargar Excel", data=bio.getvalue(),
                       file_name=f"indicadores_{today_cdmx()}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
