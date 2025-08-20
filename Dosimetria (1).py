import io
import re
import requests
import pandas as pd
import streamlit as st
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# ===================== NINOX CONFIG =====================
API_TOKEN   = "0b3a1130-785a-11f0-ace0-3fb1fcb242e2"   # <-- tu API key
TEAM_ID     = "ihp8o8AaLzfodwc4J"
DATABASE_ID = "ksqzvuts5aq0"
BASE_URL    = "https://api.ninox.com/v1"

# IDs por defecto (puedes cambiarlos en la sidebar)
DEFAULT_BASE_TABLE_ID   = "E"   # BASE DE DATOS
DEFAULT_REPORT_TABLE_ID = "C"   # REPORTE

# ===================== STREAMLIT =====================
st.set_page_config(page_title="Microsievert - Dosimetría", page_icon="🧪", layout="wide")
st.title("🧪 Sistema de Gestión de Dosimetría — Microsievert")
st.caption("Ninox + Procesamiento VALOR − CONTROL + Exportación y Carga a Ninox")

if "df_final" not in st.session_state:
    st.session_state.df_final = None

# ===================== Ninox helpers =====================
def ninox_headers():
    return {"Authorization": f"Bearer {API_TOKEN}", "Content-Type": "application/json"}

@st.cache_data(ttl=300, show_spinner=False)
def ninox_list_tables(team_id: str, db_id: str):
    url = f"{BASE_URL}/teams/{team_id}/databases/{db_id}/tables"
    r = requests.get(url, headers=ninox_headers(), timeout=30)
    r.raise_for_status()
    return r.json()

@st.cache_data(ttl=300, show_spinner=False)
def ninox_fetch_records(team_id: str, db_id: str, table_id: str, per_page: int = 1000):
    url = f"{BASE_URL}/teams/{team_id}/databases/{db_id}/tables/{table_id}/records"
    out, offset = [], 0
    while True:
        r = requests.get(url, headers=ninox_headers(), params={"perPage": per_page, "offset": offset}, timeout=60)
        r.raise_for_status()
        batch = r.json()
        if not batch: break
        out.extend(batch)
        if len(batch) < per_page: break
        offset += per_page
    rows = [x.get("fields", {}) for x in out]
    df = pd.DataFrame(rows) if rows else pd.DataFrame()
    df.columns = [str(c) for c in df.columns]  # conservar acentos/espacios
    return df

def ninox_insert_records(team_id: str, db_id: str, table_id: str, rows: list, batch_size: int = 400):
    url = f"{BASE_URL}/teams/{team_id}/databases/{db_id}/tables/{table_id}/records"
    n = len(rows)
    if n == 0:
        return {"ok": True, "inserted": 0}
    inserted = 0
    for i in range(0, n, batch_size):
        chunk = rows[i:i+batch_size]
        r = requests.post(url, headers=ninox_headers(), json=chunk, timeout=60)
        if r.status_code != 200:
            return {"ok": False, "inserted": inserted, "error": f"{r.status_code} {r.text}"}
        inserted += len(chunk)
    return {"ok": True, "inserted": inserted}

@st.cache_data(ttl=120, show_spinner=False)
def ninox_get_table_fields(team_id: str, db_id: str, table_id: str):
    """Devuelve el conjunto de nombres de campos existentes en la tabla Ninox."""
    url = f"{BASE_URL}/teams/{team_id}/databases/{db_id}/tables"
    r = requests.get(url, headers=ninox_headers(), timeout=30)
    r.raise_for_status()
    info = r.json()
    fields = set()
    for t in info:
        if str(t.get("id")) == str(table_id):
            cols = t.get("fields") or t.get("columns") or []
            for c in cols:
                name = c.get("name") if isinstance(c, dict) else None
                if name:
                    fields.add(name)
            break
    return fields

# ===================== Dosis =====================
def leer_dosis(upload):
    if not upload:
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

    # normalizar columnas
    norm = (df.columns.astype(str).str.strip().str.lower()
            .str.replace(' ', '', regex=False)
            .str.replace('(', '').str.replace(')', '')
            .str.replace('.', '', regex=False))
    df.columns = norm

    # mapear
    if 'dosimeter' not in df.columns:
        for alt in ['dosimetro', 'codigo', 'codigodosimetro', 'codigo_dosimetro']:
            if alt in df.columns:
                df.rename(columns={alt: 'dosimeter'}, inplace=True); break

    for cand in ['hp10dosecorr', 'hp10dose', 'hp10']:
        if cand in df.columns: df.rename(columns={cand: 'hp10dose'}, inplace=True); break
    for cand in ['hp007dosecorr', 'hp007dose', 'hp007']:
        if cand in df.columns: df.rename(columns={cand: 'hp0.07dose'}, inplace=True); break
    for cand in ['hp3dosecorr', 'hp3dose', 'hp3']:
        if cand in df.columns: df.rename(columns={cand: 'hp3dose'}, inplace=True); break

    # tipos
    for k in ['hp10dose', 'hp0.07dose', 'hp3dose']:
        if k in df.columns: df[k] = pd.to_numeric(df[k], errors='coerce').fillna(0.0)
        else: df[k] = 0.0

    if 'dosimeter' in df.columns:
        df['dosimeter'] = df['dosimeter'].astype(str).str.strip().str.upper()

    if 'timestamp' in df.columns:
        df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')

    return df

# ===================== Utilidades de período =====================
def periodo_desde_fecha(periodo_str: str, fecha_str: str) -> str:
    """
    Si el periodo es 'CONTROL' (o vacío), devuelve 'MES YYYY' usando FECHA DE LECTURA.
    Si ya viene un texto distinto a CONTROL, lo limpia y lo devuelve.
    """
    per = (periodo_str or "").strip().upper()
    per = re.sub(r'\.+$', '', per).strip()

    if per and per != "CONTROL":
        return per

    if not fecha_str:
        return per or ""

    try:
        fecha = pd.to_datetime(fecha_str, dayfirst=True, errors="coerce")
        if pd.isna(fecha):
            return per or ""
        meses = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
                 "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
        mes = meses[fecha.month - 1]
        return f"{mes} {fecha.year}"
    except Exception:
        return per or ""

# ===================== Cruce y cálculo =====================
def construir_registros(dfp, dfd, periodo_filtro="— TODOS —"):
    registros = []
    needed = ["NOMBRE","APELLIDO","CÉDULA","COMPAÑÍA"] + \
             [f"DOSIMETRO {i}" for i in range(1,6)] + \
             [f"PERIODO {i}" for i in range(1,6)]
    for c in needed:
        if c not in dfp.columns: dfp[c] = ""

    for _, fila in dfp.iterrows():
        nombre_raw = f"{str(fila.get('NOMBRE','')).strip()} {str(fila.get('APELLIDO','')).strip()}".strip()
        for i in range(1, 6):
            cod = str(fila.get(f'DOSIMETRO {i}', '')).strip().upper()
            per = str(fila.get(f'PERIODO {i}', '')).upper()
            if not cod or cod == "NAN": continue

            periodo_i = "CONTROL" if re.match(r'^\s*CONTROL\b', per) else re.sub(r'\.+', '.', per).strip()

            pf = (periodo_filtro or "").strip().upper()
            if pf not in ("", "— TODOS —") and periodo_i != pf:
                continue

            row = dfd.loc[dfd['dosimeter'] == cod]
            if row.empty: continue

            r0 = row.iloc[0]
            fecha = r0.get('timestamp', pd.NaT)
            fecha_str = ""
            try:
                if pd.notna(fecha): fecha_str = pd.to_datetime(fecha).strftime('%d/%m/%Y %H:%M')
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

def aplicar_valor_menos_control(registros):
    if not registros: return registros

    # Bases del primer registro (CONTROL)
    base10 = float(registros[0]['Hp(10)'])
    base07 = float(registros[0]['Hp(0.07)'])
    base3  = float(registros[0]['Hp(3)'])

    for i, r in enumerate(registros):
        # Normalizar el PERIODO usando FECHA (evita 'CONTROL')
        r['PERIODO DE LECTURA'] = periodo_desde_fecha(
            r.get('PERIODO DE LECTURA', ''),
            r.get('FECHA DE LECTURA', '')
        )

        if i == 0:
            # El primer registro solo marca NOMBRE como CONTROL; no tocar el periodo
            r['NOMBRE']  = "CONTROL"
            r['Hp(10)']  = f"{base10:.2f}"
            r['Hp(0.07)'] = f"{base07:.2f}"
            r['Hp(3)']   = f"{base3:.2f}"
        else:
            # VALOR - CONTROL y PM si diff < 0.005 (luego formateo 0.00)
            for key, base in [('Hp(10)', base10), ('Hp(0.07)', base07), ('Hp(3)', base3)]:
                diff = float(r[key]) - base
                r[key] = "PM" if diff < 0.005 else f"{diff:.2f}"

    return registros

# ===================== Excel =====================
def exportar_excel(df_final: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "REPORTE DE DOSIS"
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'),  bottom=Side(style='thin'))

    ws['I1'] = f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y')}"
    ws['I1'].font = Font(size=10, italic=True)
    ws['I1'].alignment = Alignment(horizontal='right', vertical='top')
    
    ws["I2"] = "Cliente: ____________________________"
    ws["I2"].font = Font(size=10, italic=True)
    ws["I2"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells('A5:J5')
    c = ws['A5']; c.value = 'REPORTE DE DOSIMETRÍA'
    c.font = Font(bold=True, size=14); c.alignment = Alignment(horizontal='center')

    headers = [
        'PERIODO DE LECTURA','COMPAÑÍA','CÓDIGO DE DOSÍMETRO','NOMBRE',
        'CÉDULA','FECHA DE LECTURA','TIPO DE DOSÍMETRO','Hp(10)','Hp(0.07)','Hp(3)'
    ]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=7, column=i, value=h)
        cell.font = Font(bold=True); cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill('solid', fgColor='DDDDDD'); cell.border = border

    start = 8
    for ridx, row in df_final.iterrows():
        for cidx, val in enumerate(row, 1):
            cell = ws.cell(row=start + ridx, column=cidx, value=val)
            cell.alignment = Alignment(horizontal='center', wrap_text=True)
            cell.font = Font(size=10); cell.border = border

    for col in ws.columns:
        mx = max(len(str(c.value)) if c.value else 0 for c in col) + 2
        ws.column_dimensions[get_column_letter(col[0].column)].width = mx

    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio.read()

# ===================== Sidebar =====================
with st.sidebar:
    st.header("⚙️ Configuración")
    base_table_id   = st.text_input("Table ID BASE DE DATOS", value=DEFAULT_BASE_TABLE_ID)
    report_table_id = st.text_input("Table ID REPORTE", value=DEFAULT_REPORT_TABLE_ID)
    periodo_filtro  = st.text_input("Filtro PERIODO (opcional)", value="— TODOS —")
    subir_pm_como_texto = st.checkbox("Subir 'PM' como TEXTO (si campos Hp son Texto en Ninox)", value=True)
    debug_uno = st.checkbox("Enviar 1 registro (debug)", value=False)
    show_tables = st.checkbox("Mostrar tablas Ninox (debug)", value=False)

# ===================== Conexión Ninox BASE =====================
try:
    if show_tables:
        st.expander("Tablas Ninox (debug)").json(ninox_list_tables(TEAM_ID, DATABASE_ID))
    df_participantes = ninox_fetch_records(TEAM_ID, DATABASE_ID, base_table_id)
    if df_participantes.empty:
        st.warning("No hay datos en BASE DE DATOS (Ninox).")
    else:
        st.success(f"Conectado a Ninox. Tabla BASE DE DATOS: {base_table_id}")
        st.dataframe(df_participantes.head(15), use_container_width=True)
except Exception as e:
    st.error(f"Error leyendo BASE DE DATOS: {e}")
    df_participantes = None

# ===================== Cargar Dosis =====================
st.subheader("📤 Cargar archivo de Dosis")
upload = st.file_uploader("Selecciona CSV/XLS/XLSX", type=["csv","xls","xlsx"])
df_dosis = leer_dosis(upload) if upload else None
if df_dosis is not None:
    st.caption("Vista previa dosis (normalizada):")
    st.dataframe(df_dosis.head(15), use_container_width=True)

# ===================== Procesar =====================
col1, col2 = st.columns([1,1])
with col1:
    nombre_reporte = st.text_input("Nombre archivo (sin extensión)",
                                   value=f"ReporteDosimetria_{datetime.now().strftime('%Y-%m-%d')}")
with col2:
    btn_proc = st.button("✅ Procesar", type="primary", use_container_width=True)

if btn_proc:
    if df_participantes is None or df_participantes.empty:
        st.error("No hay participantes desde Ninox.")
    elif df_dosis is None or df_dosis.empty:
        st.error("No hay datos de dosis.")
    elif 'dosimeter' not in df_dosis.columns:
        st.error("El archivo de dosis debe tener la columna 'dosimeter'.")
    else:
        with st.spinner("Procesando..."):
            registros = construir_registros(df_participantes, df_dosis, periodo_filtro=periodo_filtro)
            if not registros:
                st.warning("No hay coincidencias DOSÍMETRO ↔ dosis (revisa filtro/códigos).")
            else:
                registros = aplicar_valor_menos_control(registros)
                df_final = pd.DataFrame(registros)

                # Limpieza suave del PERIODO (sin forzar 'CONTROL')
                df_final['PERIODO DE LECTURA'] = (
                    df_final['PERIODO DE LECTURA'].astype(str)
                    .str.replace(r'\.+$', '', regex=True).str.strip()
                )
                # Aseguramos que el primer registro diga NOMBRE=CONTROL
                df_final.loc[df_final.index.min(), 'NOMBRE'] = 'CONTROL'
                df_final['NOMBRE'] = (
                    df_final['NOMBRE'].astype(str)
                    .str.replace(r'\.+$', '', regex=True).str.strip()
                )

                st.session_state.df_final = df_final
                st.success(f"¡Listo! Registros generados: {len(df_final)}")
                st.dataframe(df_final, use_container_width=True)

                try:
                    xlsx = exportar_excel(df_final)
                    st.download_button("⬇️ Descargar Excel", data=xlsx,
                        file_name=f"{(nombre_reporte.strip() or 'ReporteDosimetria')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"No se pudo generar Excel: {e}")

# ===================== Subir TODO a Ninox REPORTE =====================
st.markdown("---")
st.subheader("⬆️ Subir TODO a Ninox (tabla REPORTE)")

CUSTOM_MAP = {
    "PERIODO DE LECTURA": "PERIODO DE LECTURA",
    "COMPAÑÍA": "COMPAÑÍA",
    "CÓDIGO DE DOSÍMETRO": "CÓDIGO DE DOSÍMETRO",
    "NOMBRE": "NOMBRE",
    "CÉDULA": "CÉDULA",
    "FECHA DE LECTURA": "FECHA DE LECTURA",
    "TIPO DE DOSÍMETRO": "TIPO DE DOSÍMETRO",
}
SPECIAL_MAP = {"Hp(10)": "Hp (10)", "Hp(0.07)": "Hp (0.07)", "Hp(3)": "Hp (3)"}

def resolve_dest_name(col_name: str) -> str:
    if col_name in SPECIAL_MAP: return SPECIAL_MAP[col_name]
    if col_name in CUSTOM_MAP:  return CUSTOM_MAP[col_name]
    return col_name

def _hp_value(v, as_text_pm=True):
    if isinstance(v, str) and v.strip().upper() == "PM":
        return "PM" if as_text_pm else None
    try:
        return float(v)
    except Exception:
        return v if v is not None else None

def _to_str(v):
    if pd.isna(v): return ""
    if isinstance(v, (pd.Timestamp, )):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    return str(v)

if st.button("Subir TODO a Ninox (tabla REPORTE)"):
    df_final = st.session_state.df_final
    if df_final is None or df_final.empty:
        st.error("Primero pulsa 'Procesar'.")
    else:
        try:
            ninox_fields = ninox_get_table_fields(TEAM_ID, DATABASE_ID, report_table_id)
            if not ninox_fields:
                st.warning("No pude leer los campos de la tabla en Ninox. Verifica el ID de tabla.")
        except Exception as e:
            st.error(f"No se pudo leer el esquema de la tabla Ninox: {e}")
            ninox_fields = set()

        with st.expander("Campos detectados en Ninox"):
            st.write(sorted(ninox_fields))

        rows, skipped_cols = [], set()
        iterator = df_final.head(1).iterrows() if debug_uno else df_final.iterrows()

        for _, row in iterator:
            fields_payload = {}
            for col in df_final.columns:
                dest = resolve_dest_name(col)
                if ninox_fields and dest not in ninox_fields:
                    skipped_cols.add(dest); continue
                val = row[col]
                if dest in {"Hp (10)", "Hp (0.07)", "Hp (3)"}:
                    val = _hp_value(val, as_text_pm=subir_pm_como_texto)
                else:
                    val = _to_str(val)
                fields_payload[dest] = val
            rows.append({"fields": fields_payload})

        if debug_uno:
            st.caption("Payload (primer registro):")
            st.json(rows[:1])

        with st.spinner("Subiendo a Ninox..."):
            res = ninox_insert_records(TEAM_ID, DATABASE_ID, report_table_id, rows, batch_size=300)

        if res.get("ok"):
            st.success(f"✅ Subido a Ninox: {res.get('inserted', 0)} registro(s).")
            if skipped_cols:
                st.info("Columnas omitidas por no existir en Ninox:\n- " + "\n- ".join(sorted(skipped_cols)))
            try:
                df_check = ninox_fetch_records(TEAM_ID, DATABASE_ID, report_table_id)
                st.caption("Contenido reciente en REPORTE:")
                st.dataframe(df_check.tail(len(rows)), use_container_width=True)
            except Exception:
                pass
        else:
            st.error(f"❌ Error al subir: {res.get('error')}")
            if skipped_cols:
                st.info("Revisa/crea en Ninox los campos omitidos:\n- " + "\n- ".join(sorted(skipped_cols)))




