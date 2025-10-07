
import os
import io
import json
import requests
import streamlit as st
from datetime import datetime, timedelta, date
try:
    from zoneinfo import ZoneInfo
    MX_TZ = ZoneInfo("America/Mexico_City")
except Exception:
    MX_TZ = None

# ======================
# Helpers de tiempo/fechas
# ======================
def now_cdmx():
    try:
        if MX_TZ:
            return datetime.now(MX_TZ)
    except Exception:
        pass
    return datetime.now()

def is_business_day(d: date) -> bool:
    return d.weekday() < 5  # 0..4 = Lun..Vie

def prev_business_day(d: date) -> date:
    if d.weekday() == 0:  # lunes -> viernes anterior
        return d - timedelta(days=3)
    # si cae sábado/domingo, regresar hasta viernes
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    # si es mar..vie, regresar 1 día
    if d.weekday() in (1,2,3,4):
        return d - timedelta(days=1)
    return d

def header_dates_today():
    d_latest = now_cdmx().date()
    if not is_business_day(d_latest):
        while not is_business_day(d_latest):
            d_latest -= timedelta(days=1)
    d_prev = prev_business_day(d_latest)
    return d_prev, d_latest

# ======================
# Tokens / Config SIE
# ======================
def get_env_any(*names, default=None):
    for n in names:
        v = os.environ.get(n)
        if v:
            return v
    return default

def get_secret_any(*names, default=None):
    # Streamlit secrets si existen
    try:
        for n in names:
            if n in st.secrets:
                v = st.secrets[n]
                if v:
                    return str(v)
    except Exception:
        pass
    return default

BANXICO_TOKEN = (
    get_secret_any("banxico_token", "BANXICO_TOKEN")
    or get_env_any("BANXICO_TOKEN", "banxico_token", default="")
)

SIE_SERIES = {
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

# ======================
# Cliente SIE
# ======================
def sie_range(series_id: str, start: str, end: str):
    """Regresa lista de observaciones [{'fecha': 'YYYY-MM-DD', 'dato': '...'}] o []"""
    url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series_id}/datos/{start}/{end}"
    headers = {"User-Agent":"Mozilla/5.0"}
    if BANXICO_TOKEN:
        headers["Bmx-Token"] = BANXICO_TOKEN
    r = requests.get(url, headers=headers, timeout=25)
    r.raise_for_status()
    data = r.json().get("bmx",{}).get("series",[])
    return data[0].get("datos",[]) if data else []

def obs_to_map(obs_list):
    m = {}
    for o in obs_list or []:
        f = o.get("fecha")
        d = None
        try:
            # Banxico suele regresar 'dd/MM/yyyy' en algunas series; soportemos ambos
            if "/" in f:
                d = datetime.strptime(f, "%d/%m/%Y").date()
            else:
                d = datetime.fromisoformat(f).date()
        except Exception:
            continue
        try:
            v = float(str(o.get("dato")).replace(",","").strip())
        except Exception:
            continue
        m[d] = v
    return m

def value_asof(m: dict, d: date):
    keys = sorted([k for k in m.keys() if k <= d])
    return (m[keys[-1]] if keys else None)

# ======================
# Lógica indicadores
# ======================
def fetch_values(d_prev: date, d_latest: date):
    start = (d_prev - timedelta(days=450)).isoformat()
    end   = d_latest.isoformat()

    used = {}
    def m_for(series_key):
        sid = SIE_SERIES[series_key]
        obs = sie_range(sid, start, end)
        used[series_key] = sid
        return obs_to_map(obs)

    m_fix = m_for("USD_FIX")
    m_eur = m_for("EUR_MXN")
    m_jpy = m_for("JPY_MXN")
    m_udis = m_for("UDIS")

    m_tobj = m_for("OBJETIVO")
    m_t28  = m_for("TIIE_28")
    m_t91  = m_for("TIIE_91")
    m_t182 = m_for("TIIE_182")

    m_c28  = m_for("CETES_28")
    m_c91  = m_for("CETES_91")
    m_c182 = m_for("CETES_182")
    m_c364 = m_for("CETES_364")

    def two(m, scale_pct=False):
        a = value_asof(m, d_prev)
        b = value_asof(m, d_latest)
        if scale_pct:
            def norm(x):
                if x is None: return None
                # Normaliza a fracción: si viene como 11.25 → 0.1125
                return (x/100.0) if abs(x) > 1.0 else x
            a = norm(a); b = norm(b)
        return a, b

    fix_prev,   fix_latest   = two(m_fix, scale_pct=False)
    eur_prev,   eur_latest   = two(m_eur, scale_pct=False)
    jpy_prev,   jpy_latest   = two(m_jpy, scale_pct=False)
    udis_prev,  udis_latest  = two(m_udis, scale_pct=False)

    tobj_prev,  tobj_latest  = two(m_tobj, scale_pct=True)
    t28_prev,   t28_latest   = two(m_t28,  scale_pct=True)
    t91_prev,   t91_latest   = two(m_t91,  scale_pct=True)
    t182_prev,  t182_latest  = two(m_t182, scale_pct=True)

    c28_prev,   c28_latest   = two(m_c28,  scale_pct=True)
    c91_prev,   c91_latest   = two(m_c91,  scale_pct=True)
    c182_prev,  c182_latest  = two(m_c182, scale_pct=True)
    c364_prev,  c364_latest  = two(m_c364, scale_pct=True)

    # Compra/Venta MONEX: opcional, si no hay scraping configurado, usa margen sobre FIX.
    margen_pct = 0.85  # ejemplo
    compra_prev   = fix_prev   * (1 - margen_pct/100.0) if fix_prev   else None
    compra_latest = fix_latest * (1 - margen_pct/100.0) if fix_latest else None
    venta_prev    = fix_prev   * (1 + margen_pct/100.0) if fix_prev   else None
    venta_latest  = fix_latest * (1 + margen_pct/100.0) if fix_latest else None

    return {
        "fix": (fix_prev, fix_latest),
        "eur": (eur_prev, eur_latest),
        "jpy": (jpy_prev, jpy_latest),
        "udis": (udis_prev, udis_latest),
        "tobj": (tobj_prev, tobj_latest),
        "t28": (t28_prev, t28_latest),
        "t91": (t91_prev, t91_latest),
        "t182": (t182_prev, t182_latest),
        "c28": (c28_prev, c28_latest),
        "c91": (c91_prev, c91_latest),
        "c182": (c182_prev, c182_latest),
        "c364": (c364_prev, c364_latest),
        "monex_compra": (compra_prev, compra_latest),
        "monex_venta": (venta_prev, venta_latest),
    }

def has_any_value(vals: dict) -> bool:
    for k, (a,b) in vals.items():
        if (a is not None) or (b is not None):
            return True
    return False

# ======================
# Excel
# ======================
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

def build_excel(d_prev: date, d_latest: date, values: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Indicadores"

    # Encabezados simples (tu plantilla real los puede sobrescribir)
    ws["B2"]="Fecha anterior"; ws["C2"]=d_prev;   ws["C2"].number_format="dd/mm/yyyy"
    ws["B3"]="Fecha actual";   ws["D2"]=d_latest; ws["D2"].number_format="dd/mm/yyyy"

    # Filas objetivo (mantengo tus filas base)
    rows = {
        "fix":   5,
        "monex_compra": 6,
        "monex_venta":  7,
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

    def write_pair(key, round_to=None):
        r = rows[key]
        v_prev, v_latest = values.get(key, (None, None))
        if round_to is not None:
            v_prev   = (round(v_prev, round_to)   if v_prev   is not None else None)
            v_latest = (round(v_latest, round_to) if v_latest is not None else None)
        ws[f"C{r}"] = v_prev
        ws[f"D{r}"] = v_latest

    # Escribir valores
    write_pair("fix")
    write_pair("monex_compra")
    write_pair("monex_venta")
    write_pair("jpy")
    write_pair("eur")
    write_pair("udis")
    write_pair("tobj")
    write_pair("t28")
    write_pair("t91")
    write_pair("t182")
    write_pair("c28")
    write_pair("c91")
    write_pair("c182")
    write_pair("c364")

    # Formatos: FX 4 dec; UDIS 6; Tasas en % (2)
    fx_rows = [rows["fix"], rows["eur"], rows["jpy"]]
    udis_row = rows["udis"]
    tasa_rows = [rows["tobj"], rows["t28"], rows["t91"], rows["t182"], rows["c28"], rows["c91"], rows["c182"], rows["c364"]]
    for r in fx_rows:
        ws[f"C{r}"].number_format = "0.0000"
        ws[f"D{r}"].number_format = "0.0000"
    ws[f"C{udis_row}"].number_format = "0.000000"
    ws[f"D{udis_row}"].number_format = "0.000000"
    for r in tasa_rows:
        ws[f"C{r}"].number_format = "0.00%"
        ws[f"D{r}"].number_format = "0.00%"

    
    # Rangos con nombre (compatibles con distintas versiones de openpyxl)
    try:
        from openpyxl.workbook.defined_name import DefinedName
        def add_name(name, ref):
            dn_container = wb.defined_names
            dn = DefinedName(name=name, attr_text=f"'{ws.title}'!{ref}")
            # eliminar existentes con el mismo nombre si el contenedor lo permite
            try:
                # openpyxl >=3.1: dn_container is list-like
                existing = [d for d in getattr(dn_container, "definedName", []) if getattr(d, "name", None) == name]
                for d in existing:
                    dn_container.definedName.remove(d)
            except Exception:
                pass
            if hasattr(dn_container, "append"):
                dn_container.append(dn)
            elif hasattr(dn_container, "add"):
                dn_container.add(dn)
            else:
                # como último recurso, omitir sin romper
                pass

        add_name("RANGO_FECHAS", "$C$2:$D$2")
        add_name("RANGO_USDMXN", f"$C${rows['fix']}:$D${rows['fix']}")
        add_name("RANGO_EURMXN", f"$C${rows['eur']}:$D${rows['eur']}")
        add_name("RANGO_JPYMXN", f"$C${rows['jpy']}:$D${rows['jpy']}")
        add_name("RANGO_UDIS",   f"$C${rows['udis']}:$D${rows['udis']}")
        add_name("RANGO_TOBJ",   f"$C${rows['tobj']}:$D${rows['tobj']}")
        add_name("RANGO_TIIE28", f"$C${rows['t28']}:$D${rows['t28']}")
        add_name("RANGO_TIIE91", f"$C${rows['t91']}:$D${rows['t91']}")
        add_name("RANGO_TIIE182",f"$C${rows['t182']}:$D${rows['t182']}")
        add_name("RANGO_C28",    f"$C${rows['c28']}:$D${rows['c28']}")
        add_name("RANGO_C91",    f"$C${rows['c91']}:$D${rows['c91']}")
        add_name("RANGO_C182",   f"$C${rows['c182']}:$D${rows['c182']}")
        add_name("RANGO_C364",   f"$C${rows['c364']}:$D${rows['c364']}")
    except Exception:
        # si la versión de openpyxl no soporta esta API, no interrumpimos la exportación
        pass


    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ======================
# UI
# ======================
st.set_page_config(page_title="Indicadores IMEMSA - TC Oficial", layout="centered")
st.title("Indicadores IMEMSA – Tipo de Cambio Oficial")

with st.expander("🔎 Diagnóstico rápido", expanded=False):
    st.write({
        "token_present": bool(BANXICO_TOKEN),
        "hoy_cdmx": now_cdmx().isoformat(),
    })
    if not BANXICO_TOKEN:
        st.error("No se detectó **BANXICO_TOKEN**. Configúralo en *Secrets* o variable de entorno.")
        st.stop()

# Ejecución principal
try:
    d_prev, d_latest = header_dates_today()
    vals = fetch_values(d_prev, d_latest)
    if not has_any_value(vals):
        st.error("No se recibieron datos de SIE. Revisa el token o la conectividad (firewall/proxy).")
        st.stop()
    st.success(f"Fechas: anterior={d_prev} · actual={d_latest}")
    if st.checkbox("Ver valores obtenidos"):
        st.json(vals)
    xlsx = build_excel(d_prev, d_latest, vals)
    st.download_button(
        "Descargar Excel",
        data=xlsx,
        file_name=f"Indicadores {now_cdmx().strftime('%Y-%m-%d %H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
except requests.HTTPError as e:
    st.error(f"HTTPError: {e}")
except Exception as e:
    st.error(f"Error: {e}")
