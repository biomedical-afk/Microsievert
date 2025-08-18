import os
import re
import io
import json
import time
import requests
import pandas as pd
import streamlit as st
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# =========================
# Configuración de página
# =========================
st.set_page_config(
    page_title="Dosimetría - Microsievert",
    page_icon="🧪",
    layout="wide"
)

st.title("🧪 Sistema de Gestión de Dosimetría — Microsievert")
st.caption("Conexión Ninox + Procesamiento VALOR − CONTROL + Exportación a Excel")

# =========================
# Helpers/Constantes
# =========================
NINOX_TEAM   = "ihp8o8AaLzfodwc4J"
NINOX_DB     = "ksqzvuts5aq0"
TARGET_TABLE_NAME = "BASE DE DATOS"  # nombre visible en Ninox

COLOR_HEADER = "DDDDDD"

def ninox_headers():
    api_key = st.secrets.get("NINOX_API_KEY", "")  # <- añade tu key en .streamlit/secrets.toml
    return {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }

# =========================
# Conexión Ninox
# =========================
@st.cache_data(show_spinner=False, ttl=300)
def ninox_list_tables(team_id: str, db_id: str):
    url = f"https://api.ninox.com/v1/teams/{team_id}/databases/{db_id}/tables"
    r = requests.get(url, headers=ninox_headers(), timeout=30)
    r.raise_for_status()
    return r.json()

@st.cache_data(show_spinner=False, ttl=300)
def ninox_get_table_id_by_name(team_id: str, db_id: str, target_name: str):
    tables = ninox_list_tables(team_id, db_id)
    for t in tables:
        # Cada 't' típicamente trae keys: id, name
        if str(t.get("name", "")).strip().lower() == target_name.strip().lower():
            return t.get("id")
    return None

@st.cache_data(show_spinner=False, ttl=300)
def ninox_fetch_records(team_id: str, db_id: str, table_id: str, per_page: int = 1000, max_pages: int = 10):
    """
    Descarga registros de una tabla Ninox. Maneja paginado simple por offset.
    Devuelve lista de dicts con campos en 'fields'.
    """
    url = f"https://api.ninox.com/v1/teams/{team_id}/databases/{db_id}/tables/{table_id}/records"
    results = []
    offset = 0
    page = 0
    while page < max_pages:
        params = {"perPage": per_page, "offset": offset}
        r = requests.get(url, headers=ninox_headers(), params=params, timeout=30)
        r.raise_for_status()
        batch = r.json()
        if not batch:
            break
        results.extend(batch)
        if len(batch) < per_page:
            break
        offset += per_page
        page += 1
    return results

# =========================
# Normalización de dosis
# =========================
def leer_dosis(upload):
    """
    Lee archivo de dosis (CSV/Excel) y normaliza columnas:
    - 'dosimeter' (código)
    - 'hp10dose', 'hp0.07dose', 'hp3dose'
    """
    if upload is None:
        return None

    # Leer en DataFrame
    name = upload.name.lower()
    if name.endswith(".csv"):
        # Intentar ; y luego ,
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

    # Mapear variantes
    # Dosimeter
    if 'dosimeter' not in df.columns:
        # tratar alternativas típicas
        for alt in ['dosimetro', 'codigo', 'codigo_dosimetro']:
            if alt in df.columns:
                df.rename(columns={alt: 'dosimeter'}, inplace=True)
                break

    # Hp(10)
    for cand in ['hp10dosecorr', 'hp10dose', 'hp10']:
        if cand in df.columns:
            df.rename(columns={cand: 'hp10dose'}, inplace=True)
            break
    # Hp(0.07)
    for cand in ['hp007dosecorr', 'hp007dose', 'hp007']:
        if cand in df.columns:
            df.rename(columns={cand: 'hp0.07dose'}, inplace=True)
            break
    # Hp(3)
    for cand in ['hp3dosecorr', 'hp3dose', 'hp3']:
        if cand in df.columns:
            df.rename(columns={cand: 'hp3dose'}, inplace=True)
            break

    # Asegurar columnas
    for k in ['hp10dose', 'hp0.07dose', 'hp3dose']:
        if k in df.columns:
            df[k] = pd.to_numeric(df[k], errors='coerce').fillna(0.0)
        else:
            df[k] = 0.0

    if 'dosimeter' in df.columns:
        df['dosimeter'] = df['dosimeter'].astype(str).str.strip().str.upper()

    # Si trae timestamp
    if 'timestamp' in df.columns:
        df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')

    return df

# =========================
# Procesamiento VALOR − CONTROL
# =========================
def procesar_valor_menos_control(registros):
    """
    registros: lista de dicts con keys: Hp(10), Hp(0.07), Hp(3)
    Primera fila es CONTROL (base)
    Regla PM: diff < 0.005  (cubre negativos)
    Redondeo a 2 decimales solo para mostrar
    """
    if not registros:
        return []

    base10 = float(registros[0]['Hp(10)'])
    base07 = float(registros[0]['Hp(0.07)'])
    base3  = float(registros[0]['Hp(3)'])

    for i, r in enumerate(registros):
        if i == 0:
            r['PERIODO DE LECTURA'] = "CONTROL"
            r['Hp(10)'] = f"{base10:.2f}"
            r['Hp(0.07)'] = f"{base07:.2f}"
            r['Hp(3)'] = f"{base3:.2f}"
        else:
            for key, base in [('Hp(10)', base10), ('Hp(0.07)', base07), ('Hp(3)', base3)]:
                diff = float(r[key]) - base  # VALOR - CONTROL
                r[key] = "PM" if diff < 0.005 else f"{diff:.2f}"

    return registros

# =========================
# Cruzar Ninox <-> Dosis
# =========================
def construir_registros(dfp, dfd, periodo_filtro="— TODOS —"):
    """
    dfp: participantes desde Ninox (DataFrame)
    dfd: dosis (DataFrame) con 'dosimeter', 'hp10dose', 'hp0.07dose', 'hp3dose'
    """
    registros = []
    for _, fila in dfp.iterrows():
        nombre_raw = f"{str(fila.get('NOMBRE','')).strip()} {str(fila.get('APELLIDO','')).strip()}".strip()

        for i in range(1, 6):
            col_d = f'DOSIMETRO {i}'
            col_p = f'PERIODO {i}'
            cod = str(fila.get(col_d, '')).strip().upper()
            raw_periodo = str(fila.get(col_p, '')).upper()

            if not cod or cod == 'NAN':
                continue

            # Normalizar periodo
            if re.match(r'^\s*CONTROL\b', raw_periodo):
                periodo_i = "CONTROL"
            else:
                periodo_i = re.sub(r'\.+', '.', raw_periodo).strip()

            # Filtrado por periodo
            if periodo_filtro not in ("", "— TODOS —") and periodo_i != periodo_filtro:
                continue

            # Buscar dosis
            row = dfd.loc[dfd['dosimeter'] == cod]
            if not row.empty:
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

# =========================
# Exportar a Excel con formato
# =========================
def exportar_excel_formato(df_final: pd.DataFrame, nombre_base: str = None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "REPORTE DE DOSIS"

    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

    # Fecha emisión
    ws['I1'] = f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y')}"
    ws['I1'].font = Font(size=10, italic=True)
    ws['I1'].alignment = Alignment(horizontal='right', vertical='top')

    # Título
    ws.merge_cells('A5:J5')
    c = ws['A5']
    c.value = 'REPORTE DE DOSIMETRÍA'
    c.font = Font(bold=True, size=14)
    c.alignment = Alignment(horizontal='center')

    # Encabezados
    headers = [
        'PERIODO DE LECTURA', 'COMPAÑÍA', 'CÓDIGO DE DOSÍMETRO',
        'NOMBRE', 'CÉDULA', 'FECHA DE LECTURA',
        'TIPO DE DOSÍMETRO', 'Hp(10)', 'Hp(0.07)', 'Hp(3)'
    ]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=7, column=i, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill('solid', fgColor=COLOR_HEADER)
        cell.border = border

    # Datos
    start_row = 8
    for idx, row in df_final.iterrows():
        for col_idx, val in enumerate(row, 1):
            c = ws.cell(row=start_row + idx, column=col_idx, value=val)
            c.alignment = Alignment(horizontal='center', wrap_text=True)
            c.font = Font(size=10)
            c.border = border

    # Ajuste ancho
    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value else 0 for c in col) + 2
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len

    # A binario
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.read()

# =========================
# UI: Panel izquierdo
# =========================
with st.sidebar:
    st.header("⚙️ Configuración")
    st.markdown("1. Añade tu **NINOX_API_KEY** a `st.secrets`.\n2. Carga tu archivo de **dosis**.\n3. Pulsa **Procesar**.")

    show_tables = st.checkbox("Mostrar tablas de Ninox (debug)", value=False)
    periodo_filtro = st.text_input("Filtrar por PERIODO (opcional)", value="— TODOS —")

# =========================
# Paso 1: Verificar API Key
# =========================
if not st.secrets.get("NINOX_API_KEY"):
    st.warning("Agrega tu **NINOX_API_KEY** en `.streamlit/secrets.toml` para conectar con Ninox.")
    st.code(
        '[global]\n'
        'disableWatchdogWarning = true\n\n'
        '[secrets]\n'
        'NINOX_API_KEY = "0b3a1130-785a-11f0-ace0-3fb1fcb242e2"\n',
        language="toml"
    )

# =========================
# Paso 2: Conectar a Ninox
# =========================
table_id = None
df_participantes = None
ninox_error = None

try:
    if st.secrets.get("NINOX_API_KEY"):
        if show_tables:
            st.subheader("Tablas disponibles en Ninox")
            _tables = ninox_list_tables(NINOX_TEAM, NINOX_DB)
            st.json(_tables)

        table_id = ninox_get_table_id_by_name(NINOX_TEAM, NINOX_DB, TARGET_TABLE_NAME)
        if not table_id:
            ninox_error = f"No se encontró la tabla '{TARGET_TABLE_NAME}'. Revisa el nombre."
        else:
            raw_records = ninox_fetch_records(NINOX_TEAM, NINOX_DB, table_id)
            # Ninox devuelve cada item con 'fields'
            rows = []
            for rec in raw_records:
                fields = rec.get("fields", {})
                rows.append(fields)
            df_participantes = pd.DataFrame(rows) if rows else pd.DataFrame()
            # Normalizar columnas de interés a mayúsculas uniformes
            df_participantes.columns = [str(c).upper().strip() for c in df_participantes.columns]
            # Si falta alguna, aún así seguimos; el procesamiento valida
except Exception as e:
    ninox_error = f"Error conectando a Ninox: {e}"

if ninox_error:
    st.error(ninox_error)
else:
    st.success(f"Conectado a Ninox. Tabla: **{TARGET_TABLE_NAME}** (id: `{table_id}`)")

# =========================
# Paso 3: Subir archivo de Dosis
# =========================
st.subheader("📤 Cargar archivo de Dosis (CSV/Excel)")
upload = st.file_uploader("Selecciona tu archivo de dosis", type=["csv", "xls", "xlsx"])
df_dosis = leer_dosis(upload) if upload else None

if df_dosis is not None:
    st.caption("Vista previa de dosis (normalizada):")
    st.dataframe(df_dosis.head(20), use_container_width=True)

# =========================
# Paso 4: Procesar
# =========================
col_a, col_b = st.columns([1, 1])
with col_a:
    nombre_reporte = st.text_input("Nombre del archivo de reporte (sin extensión)", 
                                   value=f"ReporteDosimetria_{datetime.now().strftime('%Y-%m-%d')}")
with col_b:
    run_btn = st.button("✅ Procesar", type="primary", use_container_width=True)

if run_btn:
    if df_participantes is None or df_participantes.empty:
        st.error("No hay datos de participantes desde Ninox. Verifica la conexión/tabla.")
    elif df_dosis is None or df_dosis.empty:
        st.error("No hay datos de dosis. Sube un archivo CSV/Excel válido.")
    else:
        # Validaciones mínimas
        required_part_cols = ["NOMBRE","APELLIDO","CÉDULA","COMPAÑÍA"] + \
                             [f"DOSIMETRO {i}" for i in range(1,6)] + \
                             [f"PERIODO {i}" for i in range(1,6)]
        faltantes = [c for c in required_part_cols if c not in df_participantes.columns]
        if faltantes:
            st.warning(f"Faltan columnas en participantes (Ninox): {faltantes}")

        if 'dosimeter' not in df_dosis.columns:
            st.error("El archivo de dosis debe incluir una columna de **dosimeter** (código).")
        else:
            # Construir registros y aplicar lógica VALOR − CONTROL
            with st.spinner("Procesando..."):
                registros = construir_registros(df_participantes, df_dosis, periodo_filtro=periodo_filtro.strip().upper())
                if not registros:
                    st.warning("No se encontraron coincidencias DOSÍMETRO ↔ dosis (con este filtro).")
                else:
                    registros = procesar_valor_menos_control(registros)
                    df_final = pd.DataFrame(registros)
                    # Normalizar "CONTROL..." por si algo quedó
                    if 'PERIODO DE LECTURA' in df_final.columns:
                        df_final['PERIODO DE LECTURA'] = (
                            df_final['PERIODO DE LECTURA']
                            .astype(str).str.upper()
                            .str.replace(r'^\s*CONTROL.*$', 'CONTROL', regex=True)
                            .str.replace(r'\.+$', '', regex=True)
                            .str.strip()
                        )

                    st.success(f"¡Listo! Registros generados: {len(df_final)}")
                    st.dataframe(df_final, use_container_width=True)

                    # Exportar a Excel con formato
                    try:
                        xlsx_bytes = exportar_excel_formato(df_final, nombre_base=nombre_reporte.strip())
                        st.download_button(
                            label="⬇️ Descargar Excel",
                            data=xlsx_bytes,
                            file_name=f"{nombre_reporte.strip() or 'ReporteDosimetria'}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"No se pudo generar el Excel formateado: {e}")

