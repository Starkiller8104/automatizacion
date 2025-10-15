
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
BANXICO_TOKEN = os.environ.get("BANXICO_TOKEN", "677aaedf11d11712aa2ccf73da4d77b6b785474eaeb2e092f6bad31b29de6609")  # usa secrets en prod

MX_TZ = pytz.timezone("America/Mexico_City")
SIE_URL = "https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series}/datos/{start}/{end}"

def today_mx() -> dt.date:
    return dt.datetime.now(MX_TZ).date()

def is_na(x) -> bool:
    try:
        return pd.isna(x)
    except Exception:
        return x is None

def to_pydate(x):
    """Convierte a datetime.date y maneja None/NaT/Timestamp/Datetime."""
    if x is None:
        return None
    if isinstance(x, dt.date) and not isinstance(x, dt.datetime):
        return x
    if isinstance(x, dt.datetime):
        return x.date()
    # pandas Timestamp o NaT
    try:
        if isinstance(x, pd.Timestamp):
            if pd.isna(x):
                return None
            return x.to_pydatetime().date()
    except Exception:
        pass
    # numpy datetime64
    try:
        return pd.to_datetime(x, errors="coerce").date()
    except Exception:
        return None

def iso_date(x) -> str:
    d = to_pydate(x)
    return d.strftime("%Y-%m-%d") if d else ""

def last_value_leq_date(series_id: str, token: str, target_date: dt.date, lookback_days: int = 10):
    start = target_date - dt.timedelta(days=lookback_days)
    url = SIE_URL.format(series=series_id, start=iso_date(start), end=iso_date(target_date))
    headers = {"Bmx-Token": token.strip()}
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    js = r.json()
    series = js.get("bmx", {}).get("series", [])
    datos = series[0].get("datos", []) if series else []
    if not datos:
        return None, None
    df = pd.DataFrame(datos)
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df["dato"] = pd.to_numeric(df["dato"].replace({"N/E": None}), errors="coerce")
    df = df.dropna(subset=["dato", "fecha"]).sort_values("fecha")
    if df.empty:
        return None, None
    last = df.tail(1).iloc[0]
    return to_pydate(last["fecha"]), float(last["dato"])

def write_to_template_bytes(template_path: str, fecha_c, fecha_d, fix_c, fix_d) -> bytes:
    with open(template_path, "rb") as f:
        raw = f.read()
    wb = load_workbook(io.BytesIO(raw))
    ws = wb.active  # primera hoja

    # Fechas (como texto ISO para evitar problemas de formato regional)
    ws["C2"] = iso_date(fecha_c)
    ws["D2"] = iso_date(fecha_d)

    # Valores FIX
    ws["C5"] = float(fix_c) if fix_c is not None else None
    ws["D5"] = float(fix_d) if fix_d is not None else None

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def coalesce_date(primary, fallback):
    """Si primary es None o NaT, regresa fallback."""
    if is_na(primary):
        return fallback
    return to_pydate(primary) or fallback

def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="💵", layout="centered")
    st.title(APP_TITLE)
    st.caption("Botón único: C2/D2 (fechas ISO) y C5/D5 (USD FIX) escritos directo en tu plantilla.")

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
            fecha_c, fix_c = last_value_leq_date(SIE_SERIES_FIX, BANXICO_TOKEN, ayer)
            fecha_d, fix_d = last_value_leq_date(SIE_SERIES_FIX, BANXICO_TOKEN, hoy)

            fecha_c_final = coalesce_date(fecha_c, ayer)
            fecha_d_final = coalesce_date(fecha_d, hoy)

            xlsx_bytes = write_to_template_bytes(TEMPLATE_PATH, fecha_c_final, fecha_d_final, fix_c, fix_d)

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
