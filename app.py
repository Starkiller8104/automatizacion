
import os, io, json, requests, streamlit as st
from datetime import datetime, timedelta, date

# ======================
# Zona horaria y corte 12:00 CDMX
# ======================
try:
    from zoneinfo import ZoneInfo
    CDMX_TZ = ZoneInfo("America/Mexico_City")
except Exception:
    CDMX_TZ = None

def now_cdmx():
    try:
        if CDMX_TZ:
            return datetime.now(CDMX_TZ)
    except Exception:
        pass
    return datetime.now()

def is_business(d: date) -> bool:
    return d.weekday() < 5

def prev_business(d: date) -> date:
    if d.weekday() == 0:  # lunes
        return d - timedelta(days=3)
    while d.weekday() >= 5:  # fin de semana
        d -= timedelta(days=1)
    if d.weekday() in (1,2,3,4):
        return d - timedelta(days=1)
    return d

def header_dates_effective():
    """
    (d_prev, d_latest) con corte 12:00 CDMX:
    - Si hora < 12: d_latest = día hábil anterior
    - Si hora >= 12: d_latest = hoy (si es hábil; si no, retrocede hasta hábil)
    - d_prev = hábil anterior a d_latest
    """
    now = now_cdmx()
    d = now.date()
    while not is_business(d):
        d -= timedelta(days=1)
    d_latest = prev_business(d) if now.hour < 12 else d
    d_prev = prev_business(d_latest)
    return d_prev, d_latest

# ======================
# BANXICO_TOKEN (una sola lectura, temprano)
# ======================
def _get_secret_any(*names):
    try:
        for n in names:
            if n in st.secrets and st.secrets[n]:
                return str(st.secrets[n]).strip()
    except Exception:
        pass
    return None

BANXICO_TOKEN = (
    _get_secret_any("banxico_token", "BANXICO_TOKEN")
    or os.environ.get("BANXICO_TOKEN")
    or os.environ.get("banxico_token")
    or ""
)

# ======================
# Cliente SIE
# ======================
SESSION = requests.Session()
SESSION.headers.update({"User-Agent":"Mozilla/5.0"})

def sie_range(series_id: str, start_iso: str, end_iso: str):
    url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series_id}/datos/{start_iso}/{end_iso}"
    headers = {}
    if BANXICO_TOKEN:
        headers["Bmx-Token"] = BANXICO_TOKEN
    r = SESSION.get(url, headers=headers, timeout=25)
    r.raise_for_status()
    data = r.json().get("bmx", {}).get("series", [])
    return data[0].get("datos", []) if data else []

def parse_date_any(s):
    try:
        if isinstance(s, str) and "/" in s:
            return datetime.strptime(s, "%d/%m/%Y").date()
        return datetime.fromisoformat(str(s)).date()
    except Exception:
        return None

def obs_to_map_as_datekey(obs_list):
    m = {}
    for o in obs_list or []:
        d = parse_date_any(o.get("fecha"))
        try:
            v = float(str(o.get("dato")).replace(",","").strip())
        except Exception:
            v = None
        if d and (v is not None):
            m[d] = v
    return m

def value_asof(m: dict, d: date):
    keys = sorted([k for k in m.keys() if k <= d])
    return (m[keys[-1]] if keys else None)

# ======================
# Series oficiales + normalización
# ======================
SIE = {
    "USD_FIX":   "SF43718",
    "EUR_MXN":   "SF46410",
    "JPY_MXN":   "SF46406",
    "UDIS":      "SP68257",
    "CETES_28":  "SF43936",
    "CETES_91":  "SF43939",
    "CETES_182": "SF43942",
    "CETES_364": "SF43945",
    "TIIE_28":   "SF60648",
    "TIIE_91":   "SF60649",
    "TIIE_182":  "SF60650",
    "OBJETIVO":  "SF61745",
}

def fetch_values_two_dates(d_prev: date, d_latest: date):
    start = (d_prev - timedelta(days=450)).isoformat()
    end   = d_latest.isoformat()

    def m_for(key):
        sid = SIE[key]
        return obs_to_map_as_datekey(sie_range(sid, start, end))

    m_fix  = m_for("USD_FIX")
    m_eur  = m_for("EUR_MXN")
    m_jpy  = m_for("JPY_MXN")
    m_udis = m_for("UDIS")

    m_tobj = m_for("OBJETIVO")
    m_t28  = m_for("TIIE_28")
    m_t91  = m_for("TIIE_91")
    m_t182 = m_for("TIIE_182")

    m_c28  = m_for("CETES_28")
    m_c91  = m_for("CETES_91")
    m_c182 = m_for("CETES_182")
    m_c364 = m_for("CETES_364")

    def two(m, pct=False):
        a = value_asof(m, d_prev)
        b = value_asof(m, d_latest)
        if pct:
            def norm(x):
                if x is None: return None
                return (x/100.0) if abs(x) > 1.0 else x
            a = norm(a); b = norm(b)
        return a, b

    return {
        "fix": two(m_fix),
        "eur": two(m_eur),
        "jpy": two(m_jpy),
        "udis": two(m_udis),
        "tobj": two(m_tobj, pct=True),
        "t28": two(m_t28, pct=True),
        "t91": two(m_t91, pct=True),
        "t182": two(m_t182, pct=True),
        "c28": two(m_c28, pct=True),
        "c91": two(m_c91, pct=True),
        "c182": two(m_c182, pct=True),
        "c364": two(m_c364, pct=True),
    }

def _has_any(values: dict) -> bool:
    for k,(a,b) in values.items():
        if (a is not None) or (b is not None):
            return True
    return False

# ======================
# PLANTILLA desde repo (ruta fija o env TEMPLATE_PATH)
# ======================
CANDIDATE_TPL_PATHS = [
    os.environ.get("TEMPLATE_PATH", ""),
    "plantillas/Indicadores.xlsx",
    "plantilla/Indicadores.xlsx",
    "templates/Indicadores.xlsx",
    "Indicadores.xlsx",
]
TEMPLATE_PATH = next((p for p in CANDIDATE_TPL_PATHS if p and os.path.exists(p)), None)

# ======================
# UI
# ======================
st.set_page_config(page_title="TC Oficial – PLANTILLA del Repo", layout="centered")
st.title("Tipo de Cambio Oficial — usando la PLANTILLA del Repo")

with st.expander("Diagnóstico", expanded=False):
    st.write({"token_present": bool(BANXICO_TOKEN), "now_cdmx": now_cdmx().isoformat(), "template_path": TEMPLATE_PATH})
    if not BANXICO_TOKEN:
        st.warning("No se detectó BANXICO_TOKEN (secrets/env). Pégalo aquí para esta sesión:")
        tmp = st.text_input("Bmx-Token", type="password")
        if tmp:
            BANXICO_TOKEN = tmp.strip()
            st.success("Token cargado temporalmente.")

if not TEMPLATE_PATH:
    st.error("No se encontró la PLANTILLA en el repo. Define TEMPLATE_PATH o coloca el archivo en 'plantillas/Indicadores.xlsx'.")
else:
    if st.button("Generar Excel (desde PLANTILLA del Repo)"):
        try:
            if not BANXICO_TOKEN:
                st.error("Falta BANXICO_TOKEN. Configúralo o pégalo en el panel de Diagnóstico.")
                st.stop()

            d_prev, d_latest = header_dates_effective()
            vals = fetch_values_two_dates(d_prev, d_latest)
            if not _has_any(vals):
                st.error("SIE no regresó datos (token inválido o red bloqueada).")
                st.stop()

            from openpyxl import load_workbook
            wb = load_workbook(TEMPLATE_PATH, data_only=False)
            # Si tu plantilla tiene un nombre de hoja específico, ponlo aquí:
            # ws = wb["Indicadores"]
            ws = wb.active

            # Fechas
            ws["C2"].value = d_prev
            ws["D2"].value = d_latest
            ws["C2"].number_format = "dd/mm/yyyy"
            ws["D2"].number_format = "dd/mm/yyyy"

            # Filas esperadas en la plantilla (1-based)
            rows = {
                "fix":   5,
                "jpy":   10,
                "eur":   14,
                "udis":  18,
                "tobj":  21,
                "t28":   22,
                "t91":   23,
                "t182":  24,
                "c28":   27,
                "c91":   28,
                "c182":  29,
                "c364":  30,
            }

            def write_pair(key, num_fmt=None):
                r = rows[key]
                a,b = vals.get(key, (None,None))
                ws[f"C{r}"].value = a
                ws[f"D{r}"].value = b
                if num_fmt:
                    ws[f"C{r}"].number_format = num_fmt
                    ws[f"D{r}"].number_format = num_fmt

            # FX / UDIS / Tasas
            write_pair("fix",  "0.0000")
            write_pair("eur",  "0.0000")
            write_pair("jpy",  "0.0000")
            write_pair("udis", "0.000000")
            for k in ["tobj","t28","t91","t182","c28","c91","c182","c364"]:
                write_pair(k, "0.00%")

            out = io.BytesIO()
            wb.save(out); out.seek(0)

            st.success(f"Fechas: anterior={d_prev} · actual={d_latest}")
            st.download_button(
                "Descargar Excel (desde PLANTILLA del Repo)",
                data=out.getvalue(),
                file_name=f"Indicadores {now_cdmx().strftime('%Y-%m-%d %H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except requests.HTTPError as e:
            st.error(f"HTTPError SIE: {e}")
        except Exception as e:
            st.error(f"Error: {e}")
