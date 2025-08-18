# app.py
import io
import re
import math
import requests
import pandas as pd
import streamlit as st
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# ========== NINOX API CONFIG ==========
API_TOKEN   = "0b3a1130-785a-11f0-ace0-3fb1fcb242e2"   # tu API key
TEAM_ID     = "ihp8o8AaLzfodwc4J"
DATABASE_ID = "ksqzvuts5aq0"
BASE_URL    = "https://api.ninox.com/v1"

# ========== STREAMLIT CONFIG ==========
st.set_page_config(page_title="Microsievert - Dosimetría", page_icon="🧪", layout="wide")
st.title("🧪 Sistema de Gestión de Dosimetría — Microsievert")
st.caption("Ninox + Procesamiento VALOR − CONTROL + Exportación y Carga a Ninox")

COLOR_HEADER = "DDDDDD"

# --- Session state (persistir entre clics) ---
if "df_final" not in st.session_state:
    st.session_state.df_final = None

# ------------------------ Helpers Ninox ------------------------
def ninox_headers():
    return {"Authorization": f"Bearer {API_TOKEN}", "Content-Type": "application/json"}

@st.cache_data(show_spinner=False, ttl=300)
def ninox_list_tables(team_id: str, db_id: str):
    url = f"{BASE_URL}/teams/{team_id}/databases/{db_id}/tables"
    r = requests.get(url, headers=ninox_headers(), timeout=30)
    r.raise_for_status()
    return r.json()

@st.cache_data(show_spinner=False, ttl=300)
def ninox_fetch_records(team_id: str, db_id: str, table_id: str, per_page: int = 1000, max_pages: int = 20):
    url = f"{BASE_URL}/teams/{team_id}/databases/{db_id}/tables/{table_id}/records"
    results, offset, page = [], 0, 0
    while page < max_pages:
        r = requests.get(url, headers=ninox_headers(), params={"perPage": per_page, "offset": offset}, timeout=60)
        r.raise_for_status()
        batch = r.json()
        if not batch:
            break
        results.extend(batch)
        if len(batch) < per_page:
            break
        offset += per_page
        page += 1
    rows = [rec.get("fields", {}) for rec in results]
    df = pd.DataFrame(rows) if rows else pd.DataFrame()
    df.columns = [str(c).upper().strip() for c in df.columns]
    return df

def ninox_insert_records(team_id: str, db_id: str, table_id: str, rows: list, batch_size: int = 400):
    """Inserta registros en lotes."""
    url = f"{BASE_URL}/teams/{team_id}/databases/{db_id}/tables/{table_id}/records"
    headers = ninox_headers()
    n = len(rows)
    if n == 0:
        return {"ok": True, "inserted": 0}
    total_batches, inserted = math.ceil(n / batch_size), 0
    for b in range(total_batches):
        chunk = rows[b*batch_size:(b+1)*batch_size]
        r = requests.post(url, headers=headers, json=chunk, timeout=60)
        if r.status_code != 200:
            return {"ok": False, "inserted": inserted, "error": f"{r.status_code} {r.text}"}
        inserted += len(chunk)
    return {"ok": True, "inserted": inserted}

# ------------------------ Lectura/normalización dosis ------------------------
def leer_dosis(upload):
    if upload is None:
        return None
    name = upload.name.lower()
    if name.endswith(".csv"):
        try:
            df = pd.read_csv(upload, delimiter=';', engine='python')
        except Exception:
            upload.seek(0)
            df = pd.read_csv(upload)
    else:
        df = pd.read_excel(upload)

    # Normalizar nombres
    norm = (df.columns.astype(str).str.strip().str.lower()
           .str.replace(' ', '', regex=False)
           .str.replace('(', '').str.replace(')', '')
           .str.replace('.', '', regex=False))
    df.columns = norm

    # Mapear 'dosimeter'
    if 'dosimeter' not in df.columns:
        for alt in ['dosimetro', 'codigo', 'codigo_dosimetro', 'codigodosimetro']:
            if alt in df.columns:
                df.rename(columns={alt: 'dosimeter'}, inplace=True)
                break

    # Mapear dosis
    for cand in ['hp10dosecorr', 'hp10dose', 'hp10']:
        if cand in df.columns: df.rename(columns={cand: 'hp10dose'}, inplace=True); break
    for cand in ['hp007dosecorr', 'hp007dose', 'hp007']:
        if cand in df.columns: df.rename(columns={cand: 'hp0.07dose'}, inplace=True); break
    for cand in ['hp3dosecorr', 'hp3dose', 'hp3']:
        if cand in df.columns: df.rename(columns={cand: 'hp3dose'}, inplace=True); break

    # Asegurar numéricos
    for k in ['hp10dose', 'hp0.07dose', 'hp3dose']:
        if k in df.columns: df[k] = pd.to_numeric(df[k], errors='coerce').fillna(0.0)
        else: df[k] = 0.0

    if 'dosimeter' in df.columns:
        df['dosimeter'] = df['dosimeter'].astype(str).str.strip().str.upper()

    # timestamp opcional
    if 'timestamp' in df.columns:
        df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')

    return df

# ------------------------ Cruce y procesamiento ------------------------
def construir_registros(dfp, dfd, periodo_filtro="— TODOS —"):
    """Cruza BASE DE DATOS (Ninox) con archivo de dosis."""
    registros = []
    expected = ["NOMBRE","APELLIDO","CÉDULA","COMPAÑÍA"] + \
               [f"DOSIMETRO {i}" for i in range(1,6)] + \
               [f"PERIODO {i}" for i in range(1,6)]
    for col in expected:
        if col not in dfp.columns:
            dfp[col] = ""

    for _, fila in dfp.iterrows():
        nombre_raw = f"{str(fila.get('NOMBRE','')).strip()} {str(fila.get('APELLIDO','')).strip()}".strip()
        for i in range(1, 6):
            cod = str(fila.get(f'DOSIMETRO {i}', '')).strip().upper()
            raw_periodo = str(fila.get(f'PERIODO {i}', '')).upper()
            if not cod or cod == "NAN":
                continue

            # Normalizar periodo (CONTROL..., CONTROL.... -> CONTROL)
            if re.match(r'^\s*CONTROL\b', raw_periodo):
                periodo_i = "CONTROL"
            else:
                periodo_i = re.sub(r'\.+', '.', raw_periodo).strip()

            # Filtro por periodo
            pf = (periodo_filtro or "").strip().upper()
            if pf not in ("", "— TODOS —") and periodo_i != pf:
                continue

            row = dfd.loc[dfd['dosimeter'] == cod]
            if row.empty:
                continue

            r0 = row.iloc[0]
            fecha = r0.get('timestamp', pd.NaT)
            fecha_str = ""
            try:
                if pd.notna(fecha):
                    fecha_str = pd.to_datetime(fecha).strftime('%d/%m/%Y %H:%M')
            except Exception:
                fecha_str = ""

            registros.append({
                'PERIODO DE LECTURA': periodo_i,
                'COMPAÑÍA': fila.get('COMPAÑÍA',''),
                'CÓDIGO DE DOSÍMETRO': cod,
                'NOMBRE': nombre_raw,
                'CÉDULA': fila.get('CÉDULA',''),
                'FECHA DE LECTURA': fecha_str,
                'TIPO DE DOSÍMETRO': 'CE',
                'Hp(10)': float(r0.get('hp10dose', 0.0)),
                'Hp(0.07)': float(r0.get('hp0.07dose', 0.0)),
                'Hp(3)': float(r0.get('hp3dose', 0.0))
            })
    return registros

def procesar_valor_menos_control(registros):
    """
    Primera fila = CONTROL (base).
    Regla PM: diff < 0.005 (incluye negativos).
    Redondeo a 2 decimales solo para mostrar.
    """
    if not registros:
        return registros
    base10 = float(registros[0]['Hp(10)'])
    base07 = float(registros[0]['Hp(0.07)'])
    base3  = float(registros[0]['Hp(3)'])
    for i, r in enumerate(registros):
        if i == 0:
            r['PERIODO DE LECTURA'] = "CONTROL"
            r['NOMBRE'] = "CONTROL"
            r['Hp(10)']  = f"{base10:.2f}"
            r['Hp(0.07)'] = f"{base07:.2f}"
            r['Hp(3)']   = f"{base3:.2f}"
        else:
            for key, base in [('Hp(10)', base10), ('Hp(0.07)', base07), ('Hp(3)', base3)]:
                diff = float(r[key]) - base     # VALOR - CONTROL
                r[key] = "PM" if diff < 0.005 else f"{diff:.2f}"
    return registros

# ------------------------ Exportar Excel ------------------------
def exportar_excel_formato(df_final: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "REPORTE DE DOSIS"
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'),  bottom=Side(style='thin'))

    # Fecha de emisión
    ws['I1'] = f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y')}"
    ws['I1'].font = Font(size=10, italic=True)
    ws['I1'].alignment = Alignment(horizontal='right', vertical='top')

    # Título
    ws.merge_cells('A5:J5')
    c = ws['A5']; c.value = 'REPORTE DE DOSIMETRÍA'
    c.font = Font(bold=True, size=14); c.alignment = Alignment(horizontal='center')

    # Encabezados
    headers = [
        'PERIODO DE LECTURA', 'COMPAÑÍA', 'CÓDIGO DE DOSÍMETRO',
        'NOMBRE', 'CÉDULA', 'FECHA DE LECTURA',
        'TIPO DE DOSÍMETRO', 'Hp(10)', 'Hp(0.07)', 'Hp(3)'
    ]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=7, column=i, value=h)
        cell.font = Font(bold=True); cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill('solid', fgColor=COLOR_HEADER); cell.border = border

    # Datos
    start_row = 8
    for idx, row in df_final.iterrows():
        for col_idx, val in enumerate(row, 1):
            c = ws.cell(row=start_row + idx, column=col_idx, value=val)
            c.alignment = Alignment(horizontal='center', wrap_text=True)
            c.font = Font(size=10); c.border = border

    # Ajuste de anchos
    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value else 0 for c in col) + 2
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len

    bio = io.BytesIO()
    wb.save(bio); bio.seek(0)
    return bio.read()

# ------------------------ UI Sidebar ------------------------
with st.sidebar:
    st.header("⚙️ Configuración")
    st.markdown("1) Conecta a Ninox.\n2) Sube **dosis**.\n3) Procesa.\n4) (Opcional) Sube a **REPORTE**.")
    manual_base_table_id   = st.text_input("Table ID BASE DE DATOS", value="E")  # module/E
    manual_report_table_id = st.text_input("Table ID REPORTE", value="C")       # module/C
    periodo_filtro = st.text_input("Filtro PERIODO (opcional)", value="— TODOS —")
    show_tables = st.checkbox("Mostrar tablas Ninox (debug)", value=False)

# ------------------------ Conexión Ninox (BASE DE DATOS) ------------------------
df_participantes = None
ninox_error = None
try:
    base_table_id = manual_base_table_id.strip() or "E"
    if show_tables:
        st.subheader("Debug Ninox - Tablas")
        st.json(ninox_list_tables(TEAM_ID, DATABASE_ID))
    df_participantes = ninox_fetch_records(TEAM_ID, DATABASE_ID, base_table_id)
    if df_participantes is None or df_participantes.empty:
        ninox_error = "No hay datos de participantes desde Ninox (tabla BASE DE DATOS)."
except Exception as e:
    ninox_error = f"Error conectando a Ninox: {e}"

if ninox_error:
    st.error(ninox_error)
else:
    st.success(f"Conectado a Ninox. Tabla BASE DE DATOS id: **{base_table_id}**")
    st.caption("Vista previa de participantes (Ninox):")
    st.dataframe(df_participantes.head(20), use_container_width=True)

# ------------------------ Cargar Dosis ------------------------
st.subheader("📤 Cargar archivo de Dosis (CSV/Excel)")
upload = st.file_uploader("Selecciona tu archivo de dosis", type=["csv", "xls", "xlsx"])
df_dosis = leer_dosis(upload) if upload else None
if df_dosis is not None:
    st.caption("Vista previa de dosis (normalizada):")
    st.dataframe(df_dosis.head(20), use_container_width=True)

# ------------------------ Procesar ------------------------
colA, colB = st.columns([1,1])
with colA:
    nombre_reporte = st.text_input("Nombre del archivo (sin extensión)",
                                   value=f"ReporteDosimetria_{datetime.now().strftime('%Y-%m-%d')}")
with colB:
    run_btn = st.button("✅ Procesar", type="primary", use_container_width=True)

if run_btn:
    if ninox_error:
        st.error(ninox_error)
    elif df_participantes is None or df_participantes.empty:
        st.error("No hay participantes desde Ninox.")
    elif df_dosis is None or df_dosis.empty:
        st.error("No hay datos de dosis. Sube un archivo.")
    elif 'dosimeter' not in df_dosis.columns:
        st.error("El archivo de dosis debe incluir una columna 'dosimeter'.")
    else:
        with st.spinner("Procesando..."):
            registros = construir_registros(df_participantes, df_dosis, periodo_filtro=periodo_filtro)
            if not registros:
                st.warning("No se encontraron coincidencias DOSÍMETRO ↔ dosis (revisa filtro o códigos).")
            else:
                registros = procesar_valor_menos_control(registros)
                df_final = pd.DataFrame(registros)

                # Normalizar "CONTROL..." → "CONTROL" en período y nombre
                if 'PERIODO DE LECTURA' in df_final.columns:
                    df_final['PERIODO DE LECTURA'] = (
                        df_final['PERIODO DE LECTURA'].astype(str).str.upper()
                        .str.replace(r'^\s*CONTROL.*$', 'CONTROL', regex=True)
                        .str.replace(r'\.+$', '', regex=True).str.strip()
                    )
                if 'NOMBRE' in df_final.columns:
                    df_final.loc[df_final.index.min(), 'NOMBRE'] = 'CONTROL'
                    df_final['NOMBRE'] = (
                        df_final['NOMBRE'].astype(str)
                        .str.replace(r'^\s*CONTROL.*$', 'CONTROL', regex=True)
                        .str.replace(r'\.+$', '', regex=True).str.strip()
                    )

                st.session_state.df_final = df_final  # persistir

                st.success(f"¡Listo! Registros generados: {len(df_final)}")
                st.dataframe(df_final, use_container_width=True)

                try:
                    xlsx_bytes = exportar_excel_formato(df_final)
                    st.download_button(
                        label="⬇️ Descargar Excel",
                        data=xlsx_bytes,
                        file_name=f"{(nombre_reporte.strip() or 'ReporteDosimetria')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"No se pudo generar el Excel formateado: {e}")

# ------------------------ Subir a Ninox (REPORTE) ------------------------
st.markdown("---")
st.subheader("⬆️ Subir el resultado a Ninox (tabla **REPORTE**)")

col_upload1, col_upload2 = st.columns(2)
with col_upload1:
    report_table_id = st.text_input("Table ID REPORTE", value=manual_report_table_id or "C")
with col_upload2:
    debug_uno = st.checkbox("Enviar 1 registro (debug)")

def _hp_to_num_or_none(v):
    """'PM' -> None; número en texto -> float; otro -> None."""
    if v is None:
        return None
    if isinstance(v, str) and v.strip().upper() == "PM":
        return None
    try:
        return float(v)
    except Exception:
        return None

def _fecha_to_iso(s):
    """Convierte '13/08/2025 13:22' -> '2025-08-13T13:22:00' cuando se pueda."""
    if not s:
        return ""
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            dt = pd.to_datetime(s, errors="coerce")
        if pd.isna(dt):
            return str(s)
        return dt.strftime("%Y-%m-%dT%H:%M:%S")
    except Exception:
        return str(s)

if st.button("Subir a Ninox (tabla REPORTE)"):
    df_final = st.session_state.df_final  # recuperar DF persistido
    if df_final is None or df_final.empty:
        st.error("Primero genera el reporte (Procesar).")
    else:
        rows = []
        iterable = df_final.head(1).iterrows() if debug_uno else df_final.iterrows()
        for _, row in iterable:
            rows.append({
                "fields": {
                    "PERIODO DE LECTURA": str(row.get("PERIODO DE LECTURA", "")),
                    "COMPAÑÍA": str(row.get("COMPAÑÍA", "")),
                    "CÓDIGO DE DOSÍMETRO": str(row.get("CÓDIGO DE DOSÍMETRO", "")),
                    "NOMBRE": str(row.get("NOMBRE", "")),
                    "CÉDULA": str(row.get("CÉDULA", "")),
                    "FECHA DE LECTURA": _fecha_to_iso(row.get("FECHA DE LECTURA", "")),
                    "TIPO DE DOSÍMETRO": str(row.get("TIPO DE DOSÍMETRO", "")),
                    "Hp(10) ACTUAL": _hp_to_num_or_none(row.get("Hp(10)", "")),
                    "Hp(0.07) ACTUAL": _hp_to_num_or_none(row.get("Hp(0.07)", "")),
                    "Hp(3) ACTUAL": _hp_to_num_or_none(row.get("Hp(3)", ""))
                }
            })

        with st.spinner("Subiendo a Ninox..."):
            url = f"{BASE_URL}/teams/{TEAM_ID}/databases/{DATABASE_ID}/tables/{report_table_id}/records"
            r = requests.post(url, headers=ninox_headers(), json=rows, timeout=60)

        if r.status_code == 200:
            st.success(f"✅ Subido a Ninox: {len(rows)} registro(s) en tabla REPORTE (id {report_table_id}).")
        else:
            st.error(f"❌ Error al subir: {r.status_code} {r.text}")
            st.caption("Activa 'Enviar 1 registro (debug)' para aislar el problema.")


