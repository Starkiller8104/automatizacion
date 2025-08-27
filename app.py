# app.py
import io
import os
from datetime import datetime, date
import pytz
import requests
import pandas as pd
import streamlit as st
import xlsxwriter

# ============ Apagador de depuración ============
DEBUG = False   # ⇐ Déjalo en False para NO mostrar leyendas/prints

def dwrite(*args, **kwargs):
    if DEBUG:
        st.write(*args, **kwargs)

def djson(*args, **kwargs):
    if DEBUG:
        st.json(*args, **kwargs)

def dprint(*args, **kwargs):
    if DEBUG:
        print(*args, **kwargs)
# ===============================================

# ------------------ Configuración de página ------------------
st.set_page_config(page_title="IMEMSA - Indicadores", layout="wide")

# ------------------ CSS corporativo ------------------
st.markdown("""
<style>
.block-container { padding-top: 1.2rem; }
.imemsa-header { display: flex; gap: 1.25rem; align-items: center; margin-bottom: 0.75rem; }
.imemsa-logo img { max-height: 70px; width: auto; border-radius: 10px; }
.imemsa-title { line-height: 1.15; }
.imemsa-title h1 { margin: 0 0 0.15rem 0; font-size: clamp(1.6rem, 2.4vw, 2.2rem); font-weight: 800; }
.imemsa-title h3 { margin: 0; font-weight: 500; opacity: 0.95; }
.imemsa-divider { height: 6px; width: 100%; border-radius: 999px; margin: 0.6rem 0 1rem 0;
  background: linear-gradient(90deg, #0A4FA3 0%, #0A4FA3 40%, #E32028 40%, #E32028 60%, #0A4FA3 60%, #0A4FA3 100%); }
.imemsa-spacer { height: 8px; }
</style>
""", unsafe_allow_html=True)

# ------------------ Encabezado ------------------
st.markdown(
    """
    <div class="imemsa-header">
      <div class="imemsa-logo">
        <img src="logo.png" alt="IMEMSA logo">
      </div>
      <div class="imemsa-title">
        <h1>Indicadores de Tipo de Cambio</h1>
        <h3>Excel con tu layout (B2..G2 fechas reales), noticias y gráficos con XlsxWriter.</h3>
      </div>
    </div>
    <div class="imemsa-divider"></div>
    <div class="imemsa-spacer"></div>
    """, unsafe_allow_html=True
)

# ------------------ Utilidades ------------------
def today_cdmx(fmt: str = "%Y%m%d") -> str:
    tz = pytz.timezone("America/Mexico_City")
    return datetime.now(tz).strftime(fmt)

def _get_secret(name: str) -> str | None:
    try:
        return st.secrets[name]
    except Exception:
        return os.getenv(name)

# ------------------ Banxico (FIX) ------------------
def fetch_banxico(tipo: str = "SF43718", dias: int = 6):
    token = _get_secret("BANXICO_TOKEN")
    if not token:
        dprint("Sin BANXICO_TOKEN; devolviendo lista vacía.")
        return []
    try:
        url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{tipo}/datos/ultimos/{dias}"
        r = requests.get(url, headers={"Bmx-Token": token}, timeout=15)
        if r.status_code != 200:
            return []
        serie = r.json().get("bmx", {}).get("series", [{}])[0].get("datos", [])
        out = []
        for d in serie:
            fecha, valor = d.get("fecha"), d.get("dato")
            if fecha and valor and valor not in ("N/E","NR"):
                try:
                    dt = datetime.strptime(fecha, "%d/%m/%Y").date().isoformat()
                    out.append((dt, float(valor)))
                except Exception:
                    pass
        return out
    except Exception as e:
        dprint("fetch_banxico error:", e)
        return []

# ------------------ FRED ------------------
def fred_fetch_series(series_id: str, start: str | None = None, end: str | None = None, units: str = "lin") -> pd.DataFrame:
    api_key = _get_secret("FRED_API_KEY")
    if not api_key:
        return pd.DataFrame(columns=["date", "value"])
    params = {"series_id": series_id, "api_key": api_key, "file_type": "json", "units": units}
    if start: params["observation_start"] = start
    if end:   params["observation_end"]   = end
    try:
        r = requests.get("https://api.stlouisfed.org/fred/series/observations", params=params, timeout=20)
        if r.status_code != 200:
            return pd.DataFrame(columns=["date", "value"])
        data = r.json().get("observations", [])
        df = pd.DataFrame(data)[["date","value"]] if data else pd.DataFrame(columns=["date","value"])
        if not df.empty:
            df["value"] = pd.to_numeric(df["value"], errors="coerce")
        return df
    except Exception:
        return pd.DataFrame(columns=["date","value"])

# ------------------ Construcción de Excel ------------------
def build_excel(buffer: io.BytesIO,
                datos_fix: list[tuple[str, float]],
                fred_df: pd.DataFrame | None = None,
                fred_series_id: str = "DGS10"):
    wb = xlsxwriter.Workbook(buffer, {"in_memory": True})
    # Hoja Banxico
    ws = wb.add_worksheet("Indicadores")
    bold, money = wb.add_format({"bold": True}), wb.add_format({"num_format": "#,##0.0000"})
    ws.write("A1", "IMEMSA - Indicadores de Tipo de Cambio", bold)
    ws.write("A2", f"Generado: {today_cdmx('%Y-%m-%d %H:%M')} (CDMX)")
    ws.write_row("A4", ["Fecha", "FIX"], bold)
    row = 4
    for fecha_iso, valor in datos_fix:
        ws.write(row, 0, fecha_iso)
        ws.write_number(row, 1, valor, money)
        row += 1
    ws.set_column("A:A", 14); ws.set_column("B:B", 12)
    if len(datos_fix) >= 2:
        chart = wb.add_chart({"type": "line"})
        chart.add_series({"name":"Tipo de cambio FIX",
                          "categories":f"=Indicadores!$A$5:$A${row}",
                          "values":f"=Indicadores!$B$5:$B${row}"})
        chart.set_title({"name":"Últimos días"})
        chart.set_y_axis({"num_format":"#,##0.0000"})
        ws.insert_chart("D4", chart)
    # Hoja FRED
    if fred_df is not None and not fred_df.empty:
        wsname = f"FRED_{fred_series_id[:25]}"
        ws2 = wb.add_worksheet(wsname)
        fmt_bold, fmt_num, fmt_date = wb.add_format({"bold": True}), wb.add_format({"num_format": "#,##0.0000"}), wb.add_format({"num_format": "yyyy-mm-dd"})
        ws2.write("A1", f"FRED – {fred_series_id}", fmt_bold)
        ws2.write("A2", f"Generado: {today_cdmx('%Y-%m-%d %H:%M')} (CDMX)")
        ws2.write_row("A4", ["date", fred_series_id], fmt_bold)
        r0, r = 4, 4
        for _, rr in fred_df.iterrows():
            try:
                d = pd.to_datetime(rr["date"]).to_pydatetime()
                ws2.write_datetime(r, 0, d, fmt_date)
            except Exception:
                ws2.write(r, 0, str(rr.get("date","")))
            val = rr.get("value", None)
            if pd.isna(val):
                ws2.write_blank(r, 1, None)
            else:
                ws2.write_number(r, 1, float(val), fmt_num)
            r += 1
        ws2.set_column("A:A", 12); ws2.set_column("B:B", 16)
        if len(fred_df.dropna()) >= 2:
            chart2 = wb.add_chart({"type":"line"})
            chart2.add_series({"name":fred_series_id,
                               "categories":f"={wsname}!$A${r0+1}:$A${r}",
                               "values":f"={wsname}!$B${r0+1}:$B${r}"})
            chart2.set_title({"name":f"{fred_series_id} (FRED)"})
            chart2.set_y_axis({"num_format":"#,##0.0000"})
            ws2.insert_chart("D4", chart2)
    wb.close(); buffer.seek(0)

# ------------------ UI Opciones ------------------
with st.expander("Opciones FRED", expanded=False):
    add_fred   = st.checkbox("Agregar hoja FRED al Excel", value=True)
    fred_id    = st.text_input("Serie FRED", value="DGS10")
    start_date = st.date_input("Inicio FRED", value=date(2015,1,1))
    end_date   = st.date_input("Fin FRED", value=date.today())
    units      = st.selectbox("Unidades", ["lin","chg","ch1","pch","pc1","pca","cch","cca","log"], index=0)

# ------------------ Generar y Descargar ------------------
if st.button("Generar Excel"):
    datos_fix = fetch_banxico(dias=6)
    fred_df = None
    if add_fred and fred_id.strip():
        fred_df = fred_fetch_series(series_id=fred_id.strip(),
                                    start=start_date.isoformat(),
                                    end=end_date.isoformat(),
                                    units=units)
    bio = io.BytesIO()
    build_excel(bio, datos_fix, fred_df, fred_id.strip() or "DGS10")
    st.download_button("Descargar Excel",
                       data=bio.getvalue(),
                       file_name=f"indicadores_{today_cdmx()}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
