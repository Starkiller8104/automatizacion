
import io, os, tempfile
from pathlib import Path
from datetime import datetime, timedelta, timezone
import streamlit as st

# =========================
# Utiles de fecha / zona MX
# =========================
def today_cdmx():
    try:
        import pytz
        tz = pytz.timezone("America/Mexico_City")
        return datetime.now(tz).replace(tzinfo=None)
    except Exception:
        return datetime.now()

def last_6_business_days(end=None):
    end = end or today_cdmx().date()
    days = []
    d = end
    while len(days) < 6:
        if d.weekday() < 5:
            days.append(d)
        d -= timedelta(days=1)
    return list(reversed(days))

# =============================================
# Exportador que usa SIEMPRE la plantilla nueva
# =============================================
TEMPLATE_PATH = str((Path(__file__).parent / "Indicadores_template.xlsx").resolve())

def export_indicadores_template_bytes():
    # Importar el writer v3 (debe estar en el repo)
    from app_writer_layout_v3 import write_layout_v3

    # Cabecera de fechas B..G (G=hoy)
    header_dates = last_6_business_days()

    # Aquí podrías mapear tus datos reales; por ahora solo None para no romper formato
    payload = {
        "DOLAR": {},
        "YEN": {},
        "EURO": {},
        "UDIS": {},
        "TIIE": {},
        "CETES": {},
        "UMA": {},
    }

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.close()
    write_layout_v3(TEMPLATE_PATH, tmp.name, header_dates=header_dates, payload=payload)
    with open(tmp.name, "rb") as f:
        content = f.read()
    try:
        os.unlink(tmp.name)
    except Exception:
        pass
    return content

# ==============
# UI minimalista
# ==============
st.set_page_config(page_title="Indicadores", layout="wide")
st.title("Indicadores (plantilla nueva)")

if "xlsx_bytes" not in st.session_state:
    st.session_state["xlsx_bytes"] = None

col1, col2 = st.columns([1,1])
with col1:
    if st.button("Generar Excel"):
        with st.spinner("Generando a partir de la PLANTILLA nueva..."):
            st.session_state["xlsx_bytes"] = export_indicadores_template_bytes()
        st.success("Listo. Descarga abajo.")

st.markdown("---")
st.subheader("Descarga (plantilla nueva)")
st.download_button(
    "Descargar Excel",
    data=(st.session_state["xlsx_bytes"] or b""),
    file_name="Indicadores " + today_cdmx().strftime("%Y-%m-%d %H%M%S") + ".xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    disabled=(st.session_state["xlsx_bytes"] is None),
    key="dl_new_template_only",
)

st.caption("Plantilla usada: " + TEMPLATE_PATH)

