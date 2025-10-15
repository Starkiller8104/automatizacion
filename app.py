
import io
import os
import datetime as dt
import pytz
import requests
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# =======================
# CONFIGURACIÓN MÍNIMA
# =======================
APP_TITLE = "USD FIX → Plantilla (botón único)"
TEMPLATE_PATH = os.environ.get("PLANTILLA_TC_PATH", "plantilla_tc.xlsx")  # ajusta a tu ruta real
SIE_SERIES_FIX = os.environ.get("SIE_USD_FIX", "SF43718")
BANXICO_TOKEN = os.environ.get("BANXICO_TOKEN", "677aaedf11d11712aa2ccf73da4d77b6b785474eaeb2e092f6bad31b29de6609")  # cambia a secrets en producción

MX_TZ = pytz.timezone("America/Mexico_City")
SIE_URL = "https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series}/datos/{start}/{end}"

def today_mx() -> dt.date:
    return dt.datetime.now(MX_TZ).date()

def iso(d: dt.date) -> str:
    return d.strftime("%Y-%m-%d")

def last_value_leq_date(series_id: str, token: str, target_date: dt.date, lookback_days: int = 10):
    start = target_date - dt.timedelta(days=lookback_days)
    url = SIE_URL.format(series=series_id, start=iso(start), end=iso(target_date))
    headers = {"Bmx-Token": token.strip()}
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    js = r.json()
    series = js.get("bmx", {}).get("series", [])
    datos = series[0].get("datos", []) if series else []
    if not datos:
        return None, None
    df = pd.DataFrame(datos)
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce").dt.date
    df["dato"] = pd.to_numeric(df["dato"].replace({"N/E": None}), errors="coerce")
    df = df.dropna(subset=["dato"]).sort_values("fecha")
    if df.empty:
        return None, None
    last = df.tail(1).iloc[0]
    return last["fecha"], float(last["dato"])

def write_to_template_bytes(template_path: str, fecha_c: dt.date, fecha_d: dt.date, fix_c: float, fix_d: float) -> bytes:
    with open(template_path, "rb") as f:
        raw = f.read()
    wb = load_workbook(io.BytesIO(raw))
    ws = wb.active  # primera hoja

    # Fechas
    ws["C2"] = iso(fecha_c)
    ws["D2"] = iso(fecha_d)

    # Valores FIX
    ws["C5"] = float(fix_c) if fix_c is not None else None
    ws["D5"] = float(fix_d) if fix_d is not None else None

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="💵", layout="centered")
    st.title(APP_TITLE)
    st.caption("Botón único: genera y descarga la plantilla con C2/D2 (fechas) y C5/D5 (USD FIX).")

    if st.button("Generar y descargar Excel"):
        if not os.path.exists(TEMPLATE_PATH):
            st.error(f"No encontré la plantilla en: {TEMPLATE_PATH}. Ajusta TEMPLATE_PATH o variable de entorno PLANTILLA_TC_PATH.")
            return
        if not BANXICO_TOKEN.strip():
            st.error("BANXICO_TOKEN no configurado.")
            return

        hoy = today_mx()
        ayer = hoy - dt.timedelta(days=1)

        with st.spinner("Consultando Banxico y actualizando plantilla..."):
            fecha_c, fix_c = last_value_leq_date(SIE_SERIES_FIX, BANXICO_TOKEN, ayer)
            fecha_d, fix_d = last_value_leq_date(SIE_SERIES_FIX, BANXICO_TOKEN, hoy)

            if fix_c is None and fix_d is None:
                st.error("No se obtuvieron datos de Banxico para las fechas solicitadas.")
                return

            # Si alguna fecha/valor es None, usa las fechas seleccionadas como fallback (y deja celda vacía si valor None)
            fecha_c_final = fecha_c or ayer
            fecha_d_final = fecha_d or hoy

            xlsx_bytes = write_to_template_bytes(TEMPLATE_PATH, fecha_c_final, fecha_d_final, fix_c, fix_d)

        st.success("Plantilla generada.")
        st.download_button(
            "Descargar plantilla actualizada",
            data=xlsx_bytes,
            file_name=f"Plantilla_USD_FIX_{iso(hoy)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Muestra mínima de validación (texto, sin widgets adicionales)
        st.write({
            "serie": SIE_SERIES_FIX,
            "ayer_solicitado": iso(ayer),
            "ayer_con_dato": iso(fecha_c) if fecha_c else None,
            "fix_ayer": fix_c,
            "hoy_solicitado": iso(hoy),
            "hoy_con_dato": iso(fecha_d) if fecha_d else None,
            "fix_hoy": fix_d,
        })

if __name__ == "__main__":
    main()
