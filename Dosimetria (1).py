import streamlit as st
import pandas as pd
import requests
from datetime import datetime
import re

st.set_page_config(page_title="Sistema de Gestión de Dosimetría", layout="centered")

# ---- CREDENCIALES NINOX ----
NINOX_TOKEN  = "d3c82d50-60d4-11f0-9dd2-0154422825e5"
TEAM_ID      = "6dA5DFvfDTxCQxpDF"
DATABASE_ID  = "vlw6nql24oek"
TABLE_NAME   = "BASE DE DATOS"

# ---- LOGIN SIMPLE ----
USUARIOS = {"Mispanama": "Maxilo2000", "usuario1": "password123"}
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if not st.session_state["autenticado"]:
    st.markdown("<h2 style='text-align:center; color:#1c6758'>Acceso al Sistema de Gestión de Dosimetría</h2>", unsafe_allow_html=True)
    usuario  = st.text_input("Usuario")
    password = st.text_input("Contraseña", type="password")
    if st.button("Ingresar"):
        if usuario in USUARIOS and password == USUARIOS[usuario]:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos.")
    st.stop()

if st.sidebar.button("Cerrar sesión"):
    st.session_state["autenticado"] = False
    st.rerun()

st.markdown(
    "<h1 style='color:#1c6758;text-align:center'>Sistema de Gestión de Dosimetría</h1>"
    "<hr style='border-top:2px solid #3d8361'>",
    unsafe_allow_html=True
)
st.write("1) Participantes desde Ninox → 2) Sube **Dosis** → 3) Elige período → 4) Descarga **CSV UTF-8 (Comma delimited)** (solo la tabla).")

# ================== NINOX ==================
BASE_URL = f"https://api.ninox.com/v1/teams/{TEAM_ID}/databases/{DATABASE_ID}"

def _ninox_request(path, params=None):
    r = requests.get(
        f"{BASE_URL}{path}",
        headers={"Authorization": f"Bearer {NINOX_TOKEN}"},
        params=params,
        timeout=30
    )
    r.raise_for_status()
    return r.json()

def resolve_table_id_by_name(name: str) -> str:
    tables = _ninox_request("/tables")
    for t in tables:
        nm = str(t.get("name","")).strip().lower()
        lb = str(t.get("label","")).strip().lower()
        if nm == name.strip().lower() or lb == name.strip().lower():
            return t["id"]
    raise RuntimeError(f"No se encontró la tabla '{name}' en Ninox.")

def fetch_participants_from_ninox(table_name: str) -> pd.DataFrame:
    table_id = resolve_table_id_by_name(table_name)
    limit, offset, all_rows = 500, 0, []
    while True:
        chunk = _ninox_request(f"/tables/{table_id}/records", params={"limit": limit, "offset": offset})
        all_rows.extend(chunk)
        if len(chunk) < limit: break
        offset += limit

    df_raw = pd.DataFrame([r.get("fields", {}) for r in all_rows])
    n = len(df_raw)

    def choose(*opts):
        norm = {str(c).strip().lower(): c for c in df_raw.columns}
        for op in opts:
            if op.lower() in norm:
                return norm[op.lower()]
        return None

    def series_or_empty(colname):
        if colname and colname in df_raw.columns:
            return df_raw[colname]
        return pd.Series([""] * n, index=df_raw.index, dtype="object")

    df = pd.DataFrame({
        "COMPAÑÍA":    series_or_empty(choose("COMPAÑÍA","COMPAÑIA","COMPANIA","EMPRESA","COMPANY")),
        "NOMBRE":      series_or_empty(choose("NOMBRE","NAME")),
        "APELLIDO":    series_or_empty(choose("APELLIDO","APELLIDOS","LASTNAME")),
        "CÉDULA":      series_or_empty(choose("CÉDULA","CEDULA","ID","DOCUMENTO")),
        "DOSIMETRO 1": series_or_empty(choose("DOSÍMETRO 1","DOSIMETRO 1","DOSIMETRO1","DOS1")),
        "DOSIMETRO 2": series_or_empty(choose("DOSÍMETRO 2","DOSIMETRO 2","DOSIMETRO2","DOS2")),
        "PERIODO 1":   series_or_empty(choose("PERÍODO 1","PERIODO 1","PERIODO1")),
        "PERIODO 2":   series_or_empty(choose("PERÍODO 2","PERIODO 2","PERIODO2")),
    })

    for c in ["DOSIMETRO 1","DOSIMETRO 2","COMPAÑÍA","NOMBRE","APELLIDO","CÉDULA"]:
        df[c] = df[c].astype(str).str.strip()

    return df

# ================== CARGA DE PARTICIPANTES ==================
with st.spinner("Leyendo participantes de Ninox…"):
    try:
        dfp = fetch_participants_from_ninox(TABLE_NAME)
        st.success(f"Participantes: {len(dfp)} filas.")
    except Exception as e:
        st.error(f"No pude leer participantes: {e}")
        st.stop()

# ================== UI PERÍODO / ARCHIVO DOSIS ==================
periodos = set()
for c in ["PERIODO 1","PERIODO 2"]:
    if c in dfp.columns:
        periodos.update(dfp[c].dropna().astype(str).str.strip().str.upper())
periodo_opciones = sorted([p for p in periodos if p and p.lower() != "nan"])

col1, col2 = st.columns(2)
with col1:
    periodo_seleccionado = st.selectbox("Periodo", options=periodo_opciones) if periodo_opciones else None
with col2:
    dosis_file = st.file_uploader("Dosis (.csv UTF-8, .xlsx, .xls)", type=["csv","xlsx","xls"])

nombre_reporte = st.text_input("Nombre (sin extensión):", value=f"ReporteDosimetria_{datetime.now().strftime('%Y-%m-%d')}")

# ================== LECTURA DE DOSIS (ROBUSTA) ==================
def leer_dosis(f):
    # CSV: detecta ; o , ; Excel: lee directo
    if f.name.lower().endswith(".csv"):
        f.seek(0)
        try:
            df = pd.read_csv(f, sep=None, engine="python")
        except Exception:
            df = None
        if df is None or (len(df.columns) == 1 and ";" in str(df.columns[0])):
            f.seek(0); df = pd.read_csv(f, delimiter=";")
        if len(df.columns) == 1 and "," in str(df.columns[0]):
            f.seek(0); df = pd.read_csv(f, delimiter=",")
    else:
        df = pd.read_excel(f)

    # 🔧 QUITA BOM + normaliza encabezados
    df.columns = (df.columns.astype(str)
                  .str.replace("\ufeff", "", regex=False)  # ← elimina BOM del primer header
                  .str.strip().str.lower()
                  .str.replace(" ","",regex=False)
                  .str.replace("(","",regex=False)
                  .str.replace(")","",regex=False))

    # Renombra columnas a estándar
    if "dosimeter" not in df.columns:
        for alt in ["dosimeterid","badgeid","serial","detector","detectorid",
                    "codigodedosimetro","dosimetro","dosímetro"]:
            if alt in df.columns:
                df = df.rename(columns={alt:"dosimeter"})
                break

    rename = {}
    if "hp10dosecorr." in df.columns:   rename["hp10dosecorr."]   = "hp10dose"
    if "hp0.07dosecorr." in df.columns: rename["hp0.07dosecorr."] = "hp0.07dose"
    if "hp3dosecorr." in df.columns:    rename["hp3dosecorr."]    = "hp3dose"
    for t_alt in ["timestamp","time","date","readingdate","readdate",
                  "readoutdate","fechadelectura","fecha"]:
        if t_alt in df.columns:
            rename[t_alt] = "timestamp"; break
    df = df.rename(columns=rename)

    # Coalesce serialno -> dosimeter
    if "serialno" in df.columns:
        if "dosimeter" in df.columns:
            df["dosimeter"] = df["dosimeter"].astype(str).str.strip()
            df["dosimeter"] = df["dosimeter"].mask(df["dosimeter"] == "", df["serialno"])
        else:
            df = df.rename(columns={"serialno":"dosimeter"})

    if "dosimeter" not in df.columns:
        raise RuntimeError(f"No encontré la columna del dosímetro. Encabezados: {list(df.columns)}")

    # Limpieza de valores
    df["dosimeter"] = df["dosimeter"].astype(str).str.strip().str.upper()
    for k in ["hp10dose","hp0.07dose","hp3dose"]:
        if k not in df.columns: df[k] = 0.0
        df[k] = pd.to_numeric(df[k].astype(str).str.replace(",",".",regex=False),
                              errors="coerce").fillna(0.0)

    # Guarda SIEMPRE el valor textual original del Timestamp
    if "timestamp" in df.columns:
        s = df["timestamp"]
        df["timestamp_raw"] = s.astype(str)  # ← EXACTO como viene en el archivo

        # Opcional: versión datetime solo para ordenar (si se puede)
        t_txt = pd.to_datetime(s, errors="coerce", dayfirst=True)
        n = pd.to_numeric(s, errors="coerce")
        t_xls = pd.to_datetime("1899-12-30") + pd.to_timedelta(n, unit="D")
        t_xls[pd.isna(n)] = pd.NaT
        df["timestamp_dt"] = t_txt.where(t_txt.notna(), t_xls)
    else:
        df["timestamp_raw"] = ""
        df["timestamp_dt"]  = pd.NaT

    return df

def clean_ts_text(s: str) -> str:
    """Evita 'NaN' / 'NaT' como texto en la salida."""
    s = str(s).strip()
    return "" if s.lower() in ("nan","nat","none","") else s

HEADERS = [
    'PERIODO DE LECTURA','COMPAÑÍA','CÓDIGO DE DOSÍMETRO','NOMBRE','CÉDULA',
    'FECHA DE LECTURA','TIPO DE DOSÍMETRO','Hp(10)','Hp(0.07)','Hp(3)'
]

def generar_csv(df_final: pd.DataFrame) -> bytes:
    # CSV con BOM y CRLF para Excel (Comma delimited)
    return df_final[HEADERS].to_csv(index=False, lineterminator="\r\n").encode("utf-8-sig")

# ================== GENERAR CSV ==================
st.divider()
if st.button("✅ Generar CSV UTF-8 (Comma delimited)", disabled=not(dosis_file and periodo_seleccionado)):
    try:
        dfd = leer_dosis(dosis_file)

        # Normaliza para empatar
        dfp['DOSIMETRO 1'] = dfp['DOSIMETRO 1'].astype(str).str.strip().str.upper()
        if 'DOSIMETRO 2' in dfp.columns:
            dfp['DOSIMETRO 2'] = dfp['DOSIMETRO 2'].astype(str).str.strip().str.upper()
        dfd['dosimeter'] = dfd['dosimeter'].astype(str).str.strip().str.upper()

        registros = []
        for _, fila in dfp.iterrows():
            for cod in [fila.get('DOSIMETRO 1',''), fila.get('DOSIMETRO 2','')]:
                cod = (cod or "").strip().upper()
                if not cod or cod == 'NAN':
                    continue

                match = dfd[dfd['dosimeter'] == cod]
                if match.empty:
                    continue

                if 'timestamp_dt' in match.columns:
                    match = match.sort_values('timestamp_dt', ascending=False)

                dosis = match.iloc[0]

                # 🔴 FECHA DE LECTURA = texto EXACTO del archivo (Timestamp)
                fecha_str = clean_ts_text(dosis.get('timestamp_raw', ""))

                nombre_raw = f"{fila.get('NOMBRE','')} {fila.get('APELLIDO','')}".strip()
                nombre = "CONTROL" if "CONTROL" in nombre_raw.upper() else nombre_raw

                registros.append({
                    'PERIODO DE LECTURA': periodo_seleccionado,
                    'COMPAÑÍA': fila.get('COMPAÑÍA',''),
                    'CÓDIGO DE DOSÍMETRO': cod,
                    'NOMBRE': nombre,
                    'CÉDULA': fila.get('CÉDULA',''),
                    'FECHA DE LECTURA': fecha_str,
                    'TIPO DE DOSÍMETRO': 'CE',
                    'Hp(10)': float(dosis.get('hp10dose',0)),
                    'Hp(0.07)': float(dosis.get('hp0.07dose',0)),
                    'Hp(3)': float(dosis.get('hp3dose',0))
                })

        if not registros:
            st.warning("No se encontraron coincidencias entre participantes (Ninox) y dosis.")
        else:
            # Ajuste por CONTROL
            control_idx = next((i for i, r in enumerate(registros) if str(r['NOMBRE']).upper() == 'CONTROL'), 0)
            base10 = float(registros[control_idx]['Hp(10)'])
            base07 = float(registros[control_idx]['Hp(0.07)'])
            base3  = float(registros[control_idx]['Hp(3)'])
            for i, r in enumerate(registros):
                if i == control_idx:
                    r['Hp(10)'] = f"{base10:.2f}"
                    r['Hp(0.07)'] = f"{base07:.2f}"
                    r['Hp(3)'] = f"{base3:.2f}"
                else:
                    for k,b in [('Hp(10)',base10),('Hp(0.07)',base07),('Hp(3)',base3)]:
                        diff = float(r[k]) - b
                        r[k] = "PM" if diff < 0.005 else f"{diff:.2f}"

            df_final = pd.DataFrame(registros, columns=HEADERS)
            if not nombre_reporte.strip():
                st.error("El nombre del archivo es obligatorio.")
            elif re.search(r'[\\/:*?\"<>|]', nombre_reporte):
                st.error("El nombre contiene caracteres no permitidos.")
            else:
                st.success(f"CSV generado con {len(df_final)} registros.")
                st.download_button(
                    label="Descargar CSV UTF-8 (Comma delimited)",
                    data=generar_csv(df_final),
                    file_name=f"{nombre_reporte.strip()}.csv",
                    mime="text/csv; charset=utf-8"
                )
    except Exception as e:
        st.error(f"Error generando CSV: {e}")

st.markdown("<div style='color:#6c757d;text-align:center;font-size:12px;'>Sistema de Gestión de Dosimetría - MicroSievert</div>", unsafe_allow_html=True)










