
import io
from datetime import datetime, timedelta
import requests
import streamlit as st
import xlsxwriter
import feedparser
from email.utils import parsedate_to_datetime

# ==========================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================
st.set_page_config(page_title="Panel Económico MX", page_icon="📈", layout="wide")
st.title("Panel Económico MX")
st.caption("Genera un Excel con hojas opcionales: Indicadores, FRED_v2 y Noticias_RSS")

# ==========================
# LECTURA DE TOKENS (secrets)
# ==========================
def get_secret(key: str) -> str:
    try:
        val = st.secrets.get(key, "")
        if isinstance(val, str):
            return val.strip()
    except Exception:
        pass
    return ""

BANXICO_TOKEN = get_secret("BANXICO_TOKEN")
FRED_API_KEY  = get_secret("FRED_API_KEY")

# ==========================
# SIDEBAR CONTROLES
# ==========================
with st.sidebar.expander("📄 Hojas del Excel", expanded=True):
    st.caption("Activa/desactiva hojas opcionales del archivo Excel")
    want_fred   = st.checkbox("Agregar hoja FRED_v2 (si hay FRED_API_KEY)", value=st.session_state.get("want_fred", True))
    want_news   = st.checkbox("Agregar hoja Noticias_RSS (MX)", value=st.session_state.get("want_news", True))
    want_charts = st.checkbox("Agregar hoja 'Gráficos' (demo)", value=st.session_state.get("want_charts", False))
    want_raw    = st.checkbox("Agregar hoja 'Datos crudos' (demo)", value=st.session_state.get("want_raw", False))
    st.session_state["want_fred"] = want_fred
    st.session_state["want_news"] = want_news
    st.session_state["want_charts"] = want_charts
    st.session_state["want_raw"] = want_raw

# ==========================
# HELPERS: FRED
# ==========================
def _fred_req():
    s = requests.Session()
    try:
        from requests.adapters import HTTPAdapter, Retry
        rty = Retry(total=3, backoff_factor=0.5, status_forcelist=[429, 500, 502, 503, 504])
        s.mount("https://", HTTPAdapter(max_retries=rty))
    except Exception:
        pass
    return s

def fred_fetch_series(series_id: str, start: str, end: str, api_key: str):
    """Devuelve lista de (fecha_iso, valor_float)."""
    url = "https://api.stlouisfed.org/fred/series/observations"
    params = {
        "series_id": series_id, "api_key": api_key, "file_type": "json",
        "observation_start": start, "observation_end": end
    }
    r = _fred_req().get(url, params=params, timeout=30)
    r.raise_for_status()
    js = r.json()
    out = []
    for o in js.get("observations", []):
        v = o.get("value")
        if v not in (None, ".", ""):
            try:
                out.append((o["date"], float(v)))
            except Exception:
                pass
    return out

def fred_write_sheet_and_charts(wb, series_dict, sheet_name="FRED_v2"):
    """Crea hoja con datos y un gráfico por serie."""
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

# ==========================
# HELPERS: Noticias MX (RSS)
# ==========================
def mx_news_get(max_items=12):
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
    for source, url in feeds:
        try:
            fp = feedparser.parse(url)
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

def mx_news_write_sheet(wb, news_list, sheet_name="Noticias_RSS"):
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

# ==========================
# BOTÓN PRINCIPAL
# ==========================
col1, col2 = st.columns([1,1])
with col1:
    gen = st.button("Generar Excel", use_container_width=True)
with col2:
    pass

def today_cdmx():
    # nombre de archivo legible
    return datetime.now().strftime("%Y%m%d_%H%M")

if gen:
    with st.spinner("Generando archivo de Excel…"):
        bio = io.BytesIO()
        wb = xlsxwriter.Workbook(bio, {'in_memory': True})

        # Hoja simple "Indicadores" (puedes reemplazarla por tu lógica actual)
        ws = wb.add_worksheet("Indicadores")
        fmt_bold = wb.add_format({"bold": True})
        ws.write_row(0, 0, ["Indicador", "Fecha", "Valor"], fmt_bold)
        ws.write_row(1, 0, ["Demo", datetime.now().strftime("%Y-%m-%d"), 1.0])

        # FRED_v2 (si toggle y API key)
        if st.session_state.get("want_fred", True) and FRED_API_KEY:
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
                    fred_data[label] = fred_fetch_series(
                        sid,
                        start_dt.strftime("%Y-%m-%d"),
                        end_dt.strftime("%Y-%m-%d"),
                        FRED_API_KEY
                    )
                except Exception:
                    fred_data[label] = []
            if any(len(v) > 0 for v in fred_data.values()):
                fred_write_sheet_and_charts(wb, fred_data, sheet_name="FRED_v2")

        # Noticias_RSS (MX) si toggle
        if st.session_state.get("want_news", True):
            try:
                news = mx_news_get(max_items=12)
            except Exception:
                news = []
            if news:
                mx_news_write_sheet(wb, news, sheet_name="Noticias_RSS")

        # Hojas demo opcionales (para respetar tus toggles previos)
        if st.session_state.get("want_raw", False):
            wsraw = wb.add_worksheet("Datos crudos")
            wsraw.write(0,0,"(Demo) Aquí irían los datos detallados.")

        if st.session_state.get("want_charts", False):
            wsg = wb.add_worksheet("Gráficos")
            wsg.write(0,0,"(Demo) Aquí irían los gráficos adicionales.")

        # Cerrar y preparar descarga
        try:
            wb.close()
            st.session_state["xlsx_bytes"] = bio.getvalue()
        except Exception:
            st.session_state["xlsx_bytes"] = bio.getvalue()

    st.success("Excel generado. Puedes descargarlo abajo.")

# Botón de descarga (visible solo si hay bytes)
xbytes = st.session_state.get("xlsx_bytes")
if xbytes:
    st.download_button(
        "Descargar Excel",
        data=xbytes,
        file_name=f"indicadores_{today_cdmx()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
