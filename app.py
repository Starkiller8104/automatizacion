
import io
import os
import datetime as dt
import pytz
import requests
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

APP_TITLE = "USD FIX → Plantilla (botón único)"
TEMPLATE_PATH = os.environ.get("PLANTILLA_TC_PATH", "plantilla_tc.xlsx")  # ajusta a tu ruta real
SIE_SERIES_FIX = os.environ.get("SIE_USD_FIX", "SF43718")
BANXICO_TOKEN = os.environ.get("BANXICO_TOKEN", "677aaedf11d11712aa2ccf73da4d77b6b785474eaeb2e092f6bad31b29de6609")

MX_TZ = pytz.timezone("America/Mexico_City")
SIE_URL = "https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series}/datos/{start}/{end}"

def today_mx() -> dt.date:
    return dt.datetime.now(MX_TZ).date()

def iso_date(x) -> str:
    if x is None:
        return ""
    if isinstance(x, dt.datetime):
        x = x.date()
    return x.strftime("%Y-%m-%d")

def fetch_series_df(series_id: str, token: str, start: dt.date, end: dt.date) -> pd.DataFrame:
    url = SIE_URL.format(series=series_id, start=iso_date(start), end=iso_date(end))
    headers = {"Bmx-Token": token.strip()}
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    js = r.json()
    series = js.get("bmx", {}).get("series", [])
    datos = series[0].get("datos", []) if series else []
    df = pd.DataFrame(datos)
    if df.empty:
        return pd.DataFrame(columns=["fecha","dato"])
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce").dt.date
    df["dato"] = pd.to_numeric(df["dato"].replace({"N/E": None}), errors="coerce")
    return df.dropna(subset=["fecha"]).sort_values("fecha")

def value_last_leq(series_id: str, token: str, target_date: dt.date, lookback_days: int = 10):
    df = fetch_series_df(series_id, token, target_date - dt.timedelta(days=lookback_days), target_date)
    df = df.dropna(subset=["dato"])
    if df.empty:
        return None, None
    row = df.tail(1).iloc[0]
    return row["fecha"], float(row["dato"])

def value_exact_on(series_id: str, token: str, date_obj: dt.date):
    df = fetch_series_df(series_id, token, date_obj, date_obj)
    row = df.loc[df["fecha"] == date_obj].dropna(subset=["dato"])
    if row.empty:
        return None, None
    r = row.iloc[0]
    return r["fecha"], float(r["dato"])

def write_to_template_bytes(template_path: str, fecha_c, fecha_d, fix_c, fix_d) -> bytes:
    with open(template_path, "rb") as f:
        raw = f.read()
    wb = load_workbook(io.BytesIO(raw))
    ws = wb.active  # primera hoja

    # Fechas (ISO)
    ws["C2"] = iso_date(fecha_c)
    ws["D2"] = iso_date(fecha_d)

    # Valores FIX
    ws["C5"] = float(fix_c) if fix_c is not None else None
    # D5: solo si hay dato EXACTO hoy; si no, deja en blanco
    ws["D5"] = float(fix_d) if fix_d is not None else None

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="💵", layout="centered")
    st.title(APP_TITLE)
    st.caption("C2/C5 = AYER (<= ayer); D2/D5 = HOY (solo si Banxico publicó hoy).")

    if st.button("Generar y descargar Excel"):
        if not os.path.exists(TEMPLATE_PATH):
            st.error(f"No encontré la plantilla en: {TEMPLATE_PATH}. Ajusta TEMPLATE_PATH o PLANTILLA_TC_PATH.")
            return
        if not BANXICO_TOKEN.strip():
            st.error("BANXICO_TOKEN no configurado.")
            return

        hoy = today_mx()
        ayer = hoy - dt.timedelta(days=1)

        with st.spinner("Consultando Banxico y actualizando plantilla..."):
            # C (ayer): último disponible <= AYER
            fecha_c, fix_c = value_last_leq(SIE_SERIES_FIX, BANXICO_TOKEN, ayer)

            # D (hoy): valor SOLO si existe EXACTAMENTE en HOY
            fecha_d, fix_d = value_exact_on(SIE_SERIES_FIX, BANXICO_TOKEN, hoy)

            xlsx_bytes = write_to_template_bytes(TEMPLATE_PATH, ayer, hoy, fix_c, fix_d)

        st.success("Plantilla generada.")
        st.download_button(
            "Descargar plantilla actualizada",
            data=xlsx_bytes,
            file_name=f"Plantilla_USD_FIX_{iso_date(hoy)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.write({
            "serie": SIE_SERIES_FIX,
            "ayer_solicitado": iso_date(ayer),
            "ayer_con_dato": iso_date(fecha_c),
            "fix_ayer": fix_c,
            "hoy_solicitado": iso_date(hoy),
            "hoy_con_dato": iso_date(fecha_d),
            "fix_hoy": fix_d,
        })

if __name__ == "__main__":
    main()
