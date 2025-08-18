# app.py
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
API_TOKEN   = "0b3a1130-785a-11f0-ace0-3fb1fcb242e2"
TEAM_ID     = "ihp8o8AaLzfodwc4J"
DATABASE_ID = "ksqzvuts5aq0"
BASE_URL    = "https://api.ninox.com/v1"

DEFAULT_BASE_TABLE_ID   = "E"   # BASE DE DATOS
DEFAULT_REPORT_TABLE_ID = "C"   # REPORTE

# ===================== STREAMLIT BASE =====================
st.set_page_config(page_title="Microsievert - Dosimetría", page_icon="🧪", layout="wide")
st.title("🧪 Sistema de Gestión de Dosimetría — Microsievert")

if "df_final" not in st.session_state:
    st.session_state.df_final = None
if "df_reporte_ninox" not in st.session_state:
    st.session_state.df_reporte_ninox = None

# ===================== Helpers Ninox =====================
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
    df.columns = [str(c) for c in df.columns]
    return df

@st.cache_data(ttl=300, show_spinner=False)
def ninox_fetch_records_with_ids(team_id: str, db_id: str, table_id: str, per_page: int = 1000):
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
    rows = []
    for rec in out:
        row = {"_id": rec.get("id")}
        row.update(rec.get("fields", {}))
        rows.append(row)
    df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["_id"])
    df.columns = [str(c) for c in df.columns]
    return df

def ninox_insert_records(team_id: str, db_id: str, table_id: str, rows: list, batch_size: int = 400):
    url = f"{BASE_URL}/teams/{team_id}/databases/{db_id}/tables/{table_id}/records"
    if not rows:
        return {"ok": True, "inserted": 0}
    inserted = 0
    for i in range(0, len(rows), batch_size):
        chunk = rows[i:i+batch_size]
        r = requests.post(url, headers=ninox_headers(), json=chunk, timeout=60)
        if r.status_code != 200:
            return {"ok": False, "inserted": inserted, "error": f"{r.status_code} {r.text}"}
        inserted += len(chunk)
    return {"ok": True, "inserted": inserted}

def ninox_update_records(team_id: str, db_id: str, table_id: str, rows: list, batch_size: int = 300):
    url = f"{BASE_URL}/teams/{team_id}/databases/{db_id}/tables/{table_id}/records"
    if not rows:
        return {"ok": True, "updated": 0}
    updated = 0
    for i in range(0, len(rows), batch_size):
        chunk = rows[i:i+batch_size]
        r = requests.post(url, headers=ninox_headers(), json=chunk, timeout=60)
        if r.status_code != 200:
            return {"ok": False, "updated": updated, "error": f"{r.status_code} {r.text}"}
        updated += len(chunk)
    return {"ok": True, "updated": updated}

@st.cache_data(ttl=120, show_spinner=False)
def ninox_get_table_fields(team_id: str, db_id: str, table_id: str):
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
                if name: fields.add(name)
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

    norm = (df.columns.astype(str).str.strip().str.lower()
            .str.replace(' ', '', regex=False)
            .str.replace('(', '').str.replace(')', '')
            .str.replace('.', '', regex=False))
    df.columns = norm

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

    for k in ['hp10dose', 'hp0.07dose', 'hp3dose']:
        if k in df.columns: df[k] = pd.to_numeric(df[k], errors='coerce').fillna(0.0)
        else: df[k] = 0.0

    if 'dosimeter' in df.columns:
        df['dosimeter'] = df['dosimeter'].astype(str).str.strip().str.upper()

    if 'timestamp' in df.columns:
        df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')

    return df

# ===================== Cruce + cálculo VALOR − CONTROL =====================
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
                diff = float(r[key]) - base  # VALOR - CONTROL
                r[key] = "PM" if diff < 0.005 else f"{diff:.2f}"
    return registros

# ===================== Excel reporte =====================
def exportar_excel(df_final: pd.DataFrame) -> bytes:
    wb = Workbook(); ws = wb.active; ws.title = "REPORTE DE DOSIS"
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'),  bottom=Side(style='thin'))
    ws['I1'] = f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y')}"
    ws['I1'].font = Font(size=10, italic=True)
    ws['I1'].alignment = Alignment(horizontal='right', vertical='top')

    ws.merge_cells('A5:P5')
    c = ws['A5']; c.value = 'REPORTE DE DOSIMETRÍA'
    c.font = Font(bold=True, size=14); c.alignment = Alignment(horizontal='center')

    headers = [
        'PERIODO DE LECTURA','COMPAÑÍA','CÓDIGO DE DOSÍMETRO','NOMBRE','CÉDULA','FECHA DE LECTURA',
        'TIPO DE DOSÍMETRO','Hp (10)','Hp (0.07)','Hp (3)',
        'Hp (10) ANUAL','Hp (0.07) ANUAL','Hp (3) ANUAL',
        'Hp (10) DE POR VIDA','Hp (0.07) DE POR VIDA','Hp (3) DE POR VIDA'
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
    subir_pm_como_texto = st.checkbox("Subir 'PM' como TEXTO (si campos Hp en Ninox son Texto)", value=True)
    debug_uno = st.checkbox("Enviar 1 registro (debug)", value=False)
    show_tables = st.checkbox("Mostrar tablas Ninox (debug)", value=False)

# ===================== TABS =====================
tab1, tab2 = st.tabs(["Procesar y subir reporte", "Actualizar acumulados (Ninox)"])

# -------------------------------------------------------------------
# TAB 1: Procesar y subir a Ninox (incluye acumulados y orden final)
# -------------------------------------------------------------------
with tab1:
    st.subheader("Procesar archivo de dosis y subir a REPORTE")

    # Conexión a BASE DE DATOS
    try:
        if show_tables:
            st.expander("Tablas Ninox (debug)").json(ninox_list_tables(TEAM_ID, DATABASE_ID))
        df_participantes = ninox_fetch_records(TEAM_ID, DATABASE_ID, base_table_id)
        if df_participantes.empty:
            st.warning("No hay datos en BASE DE DATOS (Ninox).")
        else:
            st.success(f"Conectado a Ninox. BASE DE DATOS ({base_table_id})")
            st.dataframe(df_participantes.head(12), use_container_width=True)
    except Exception as e:
        st.error(f"Error leyendo BASE DE DATOS: {e}")
        df_participantes = None

    # Cargar dosis
    upload = st.file_uploader("📤 Selecciona CSV/XLS/XLSX de Dosis", type=["csv","xls","xlsx"])
    df_dosis = leer_dosis(upload) if upload else None
    if df_dosis is not None:
        st.caption("Vista previa dosis (normalizada):")
        st.dataframe(df_dosis.head(12), use_container_width=True)

    col1, col2 = st.columns([1,1])
    with col1:
        nombre_reporte = st.text_input("Nombre archivo (sin extensión)",
                                       value=f"ReporteDosimetria_{datetime.now().strftime('%Y-%m-%d')}")
    with col2:
        btn_proc = st.button("✅ Procesar", use_container_width=True)

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

                    # Normalizar CONTROL (quitar puntos)
                    df_final['PERIODO DE LECTURA'] = (
                        df_final['PERIODO DE LECTURA'].astype(str).str.upper()
                        .str.replace(r'^\s*CONTROL.*$', 'CONTROL', regex=True)
                        .str.replace(r'\.+$', '', regex=True).str.strip()
                    )
                    df_final.loc[df_final.index.min(), 'NOMBRE'] = 'CONTROL'
                    df_final['NOMBRE'] = (
                        df_final['NOMBRE'].astype(str)
                        .str.replace(r'^\s*CONTROL.*$', 'CONTROL', regex=True)
                        .str.replace(r'\.+$', '', regex=True).str.strip()
                    )

                    # --------- AÑADIR ACUMULADOS Y ORDEN FINAL ---------
                    def _hp_to_num(x):
                        if isinstance(x, str) and x.strip().upper() == "PM": return 0.0
                        try: return float(x)
                        except Exception: return 0.0

                    def _parse_fecha_dmy(s):
                        try: return pd.to_datetime(s, dayfirst=True, errors='coerce')
                        except Exception: return pd.NaT

                    df_final['__fecha_dt'] = df_final['FECHA DE LECTURA'].apply(_parse_fecha_dmy)
                    df_final['__year'] = df_final['__fecha_dt'].dt.year

                    # Númericos auxiliares
                    df_final['__hp10'] = df_final['Hp(10)'].apply(_hp_to_num)
                    df_final['__hp07'] = df_final['Hp(0.07)'].apply(_hp_to_num)
                    df_final['__hp3']  = df_final['Hp(3)'].apply(_hp_to_num)

                    # Fuente para acumulados: tabla REPORTE (si ya cargaste) o el propio lote
                    df_src = st.session_state.get("df_reporte_ninox")
                    if df_src is not None and not df_src.empty:
                        # detectar nombres en Ninox
                        nombre_col  = next((c for c in df_src.columns if c.strip().upper() == "NOMBRE"), "NOMBRE")
                        cedula_col  = next((c for c in df_src.columns if c.strip().upper() in {"CÉDULA","CEDULA"}), "CÉDULA")
                        fecha_col   = next((c for c in df_src.columns if c.strip().upper() == "FECHA DE LECTURA"), "FECHA DE LECTURA")
                        hp10_col    = next((c for c in df_src.columns if c.strip().lower() in {"hp (10)","hp(10)","hp 10"}), "Hp (10)")
                        hp07_col    = next((c for c in df_src.columns if c.strip().lower() in {"hp (0.07)","hp(0.07)","hp 0.07"}), "Hp (0.07)")
                        hp3_col     = next((c for c in df_src.columns if c.strip().lower() in {"hp (3)","hp(3)","hp 3"}), "Hp (3)")

                        tmp = df_src.copy()
                        tmp["_hp10"] = tmp[hp10_col].apply(_hp_to_num)
                        tmp["_hp07"] = tmp[hp07_col].apply(_hp_to_num)
                        tmp["_hp3"]  = tmp[hp3_col].apply(_hp_to_num)
                        tmp["_fecha"] = pd.to_datetime(tmp[fecha_col], dayfirst=True, errors='coerce')
                        tmp["_year"]  = tmp["_fecha"].dt.year

                        key = [nombre_col, cedula_col]
                        # DE POR VIDA
                        life10 = tmp.groupby(key)["_hp10"].sum()
                        life07 = tmp.groupby(key)["_hp07"].sum()
                        life3  = tmp.groupby(key)["_hp3"].sum()
                        # ANUAL (año de cada fila del lote)
                        # para el lote, usamos el año de la fila df_final['__year'] y buscamos en tmp ese mismo año
                        # Preparamos dict {(nombre, cedula, year): suma}
                        year10 = tmp.groupby(key + ["_year"])["_hp10"].sum()
                        year07 = tmp.groupby(key + ["_year"])["_hp07"].sum()
                        year3  = tmp.groupby(key + ["_year"])["_hp3"].sum()

                        def acc_row(n, c, y, serie_life, serie_year):
                            vlife = float(serie_life.get((n, c), 0.0))
                            vyear = float(serie_year.get((n, c, y), 0.0))
                            return vyear, vlife

                        anual10, anual07, anual3, vida10, vida07, vida3 = [], [], [], [], [], []
                        for _, r in df_final.iterrows():
                            n = r['NOMBRE']; c = r['CÉDULA']; y = int(r['__year']) if pd.notna(r['__year']) else None
                            vyear10, vlife10 = acc_row(n, c, y, life10, year10)
                            vyear07, vlife07 = acc_row(n, c, y, life07, year07)
                            vyear3,  vlife3  = acc_row(n, c, y, life3,  year3)
                            anual10.append(round(vyear10, 2)); anual07.append(round(vyear07, 2)); anual3.append(round(vyear3, 2))
                            vida10.append(round(vlife10, 2));   vida07.append(round(vlife07, 2));   vida3.append(round(vlife3, 2))
                    else:
                        # Acumulados solo con el lote
                        key = ['NOMBRE','CÉDULA']
                        life10 = df_final.groupby(key)['__hp10'].sum()
                        life07 = df_final.groupby(key)['__hp07'].sum()
                        life3  = df_final.groupby(key)['__hp3'].sum()
                        year10 = df_final.groupby(key + ['__year'])['__hp10'].sum()
                        year07 = df_final.groupby(key + ['__year'])['__hp07'].sum()
                        year3  = df_final.groupby(key + ['__year'])['__hp3'].sum()

                        anual10, anual07, anual3, vida10, vida07, vida3 = [], [], [], [], [], []
                        for _, r in df_final.iterrows():
                            n = r['NOMBRE']; c = r['CÉDULA']; y = r['__year']
                            vyear10 = float(year10.get((n, c, y), 0.0)); vlife10 = float(life10.get((n, c), 0.0))
                            vyear07 = float(year07.get((n, c, y), 0.0)); vlife07 = float(life07.get((n, c), 0.0))
                            vyear3  = float(year3.get((n, c, y), 0.0));  vlife3  = float(life3.get((n, c), 0.0))
                            anual10.append(round(vyear10, 2)); anual07.append(round(vyear07, 2)); anual3.append(round(vyear3, 2))
                            vida10.append(round(vlife10, 2));   vida07.append(round(vlife07, 2));   vida3.append(round(vlife3, 2))

                    df_final['Hp (10) ANUAL'] = anual10
                    df_final['Hp (0.07) ANUAL'] = anual07
                    df_final['Hp (3) ANUAL'] = anual3
                    df_final['Hp (10) DE POR VIDA'] = vida10
                    df_final['Hp (0.07) DE POR VIDA'] = vida07
                    df_final['Hp (3) DE POR VIDA'] = vida3

                    # Orden FINAL exacto
                    final_cols = [
                        'PERIODO DE LECTURA','COMPAÑÍA','CÓDIGO DE DOSÍMETRO','NOMBRE','CÉDULA','FECHA DE LECTURA',
                        'TIPO DE DOSÍMETRO','Hp (10)','Hp (0.07)','Hp (3)',
                        'Hp (10) ANUAL','Hp (0.07) ANUAL','Hp (3) ANUAL',
                        'Hp (10) DE POR VIDA','Hp (0.07) DE POR VIDA','Hp (3) DE POR VIDA'
                    ]
                    # renombrar Hp para coincidir con "Hp (..)" en salida
                    df_final.rename(columns={"Hp(10)":"Hp (10)", "Hp(0.07)":"Hp (0.07)", "Hp(3)":"Hp (3)"}, inplace=True)
                    df_final = df_final[final_cols].copy()

                    # guardar / mostrar
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

    st.markdown("### Subir TODO a Ninox (tabla REPORTE)")
    SPECIAL_MAP = {
        "Hp (10)": "Hp (10)", "Hp (0.07)": "Hp (0.07)", "Hp (3)": "Hp (3)",
        "Hp (10) ANUAL":"Hp (10) ANUAL", "Hp (0.07) ANUAL":"Hp (0.07) ANUAL", "Hp (3) ANUAL":"Hp (3) ANUAL",
        "Hp (10) DE POR VIDA":"Hp (10) DE POR VIDA", "Hp (0.07) DE POR VIDA":"Hp (0.07) DE POR VIDA", "Hp (3) DE POR VIDA":"Hp (3) DE POR VIDA"
    }
    CUSTOM_MAP = {
        "PERIODO DE LECTURA":"PERIODO DE LECTURA", "COMPAÑÍA":"COMPAÑÍA",
        "CÓDIGO DE DOSÍMETRO":"CÓDIGO DE DOSÍMETRO", "NOMBRE":"NOMBRE",
        "CÉDULA":"CÉDULA", "FECHA DE LECTURA":"FECHA DE LECTURA", "TIPO DE DOSÍMETRO":"TIPO DE DOSÍMETRO"
    }
    def resolve_dest_name(col_name: str, ninox_fields: set) -> str:
        cand = SPECIAL_MAP.get(col_name) or CUSTOM_MAP.get(col_name) or col_name
        return cand

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

    if st.button("⬆️ Subir a Ninox (REPORTE)"):
        df_final = st.session_state.df_final
        if df_final is None or df_final.empty:
            st.error("Primero pulsa 'Procesar'.")
        else:
            try:
                ninox_fields = ninox_get_table_fields(TEAM_ID, DATABASE_ID, report_table_id)
            except Exception as e:
                st.error(f"No pude leer los campos de la tabla en Ninox: {e}")
                ninox_fields = set()

            rows, skipped_cols = [], set()
            iterator = df_final.head(1).iterrows() if debug_uno else df_final.iterrows()
            for _, row in iterator:
                payload = {}
                for col in df_final.columns:
                    dest = resolve_dest_name(col, ninox_fields)
                    if ninox_fields and dest not in ninox_fields:
                        skipped_cols.add(col); continue
                    val = row[col]
                    if dest.startswith("Hp ("):
                        val = _hp_value(val, as_text_pm=subir_pm_como_texto)
                    else:
                        val = _to_str(val)
                    payload[dest] = val
                rows.append({"fields": payload})

            if debug_uno:
                st.caption("Payload (primer registro):"); st.json(rows[:1])

            with st.spinner("Subiendo a Ninox..."):
                res = ninox_insert_records(TEAM_ID, DATABASE_ID, report_table_id, rows, batch_size=300)

            if res.get("ok"):
                st.success(f"✅ Subido a Ninox: {res.get('inserted', 0)} registro(s).")
                if skipped_cols:
                    st.info("Campos NO subidos (no existen en Ninox):\n- " + "\n- ".join(sorted(skipped_cols)))
                try:
                    df_check = ninox_fetch_records(TEAM_ID, DATABASE_ID, report_table_id)
                    st.caption("Contenido reciente en REPORTE:")
                    st.dataframe(df_check.tail(len(rows)), use_container_width=True)
                except Exception:
                    pass
            else:
                st.error(f"❌ Error al subir: {res.get('error')}")

# -------------------------------------------------------------------
# TAB 2: Actualizar acumulados Hp(10/0.07/3) ANUAL y DE POR VIDA en Ninox
# -------------------------------------------------------------------
with tab2:
    st.subheader("Calcular y escribir acumulados en REPORTE (modo seguro)")

    colY, colZ = st.columns(2)
    with colY:
        target_year = st.number_input("Año para el acumulado ANUAL", value=datetime.now().year, step=1)
    with colZ:
        solo_lote = st.checkbox("Filtrar por dosímetros de un archivo de Dosis (opcional)", value=False)

    uploaded_lote = None
    df_lote = None
    codigos_lote = set()
    if solo_lote:
        uploaded_lote = st.file_uploader("Sube el archivo de Dosis del lote a considerar", type=["csv","xls","xlsx"], key="lote")
        if uploaded_lote:
            df_lote = leer_dosis(uploaded_lote)
            if df_lote is not None and "dosimeter" in df_lote.columns:
                codigos_lote = set(df_lote["dosimeter"].astype(str).str.strip().str.upper().dropna().tolist())
                st.success(f"Códigos en lote: {len(codigos_lote)}")
            else:
                st.warning("No se encontró columna 'dosimeter' en el archivo del lote. No se filtrará.")

    if st.button("🔽 Traer REPORTE desde Ninox"):
        try:
            df_rep = ninox_fetch_records_with_ids(TEAM_ID, DATABASE_ID, report_table_id)
            if df_rep.empty:
                st.warning("La tabla REPORTE está vacía.")
            else:
                st.success(f"Leídas {len(df_rep)} filas de REPORTE.")
                st.dataframe(df_rep.head(20), use_container_width=True)
                st.session_state.df_reporte_ninox = df_rep
        except Exception as e:
            st.error(f"Error leyendo REPORTE: {e}")

    # Campos destino configurables
    st.markdown("#### Campos destino en Ninox (deben ser EDITABLES, no fórmulas)")
    colA, colB, colC = st.columns(3)
    with colA:
        anual10_col = st.text_input("Hp (10) ANUAL", value="Hp (10) ANUAL")
    with colB:
        anual07_col = st.text_input("Hp (0.07) ANUAL", value="Hp (0.07) ANUAL")
    with colC:
        anual3_col  = st.text_input("Hp (3) ANUAL", value="Hp (3) ANUAL")
    colA2, colB2, colC2 = st.columns(3)
    with colA2:
        vida10_col = st.text_input("Hp (10) DE POR VIDA", value="Hp (10) DE POR VIDA")
    with colB2:
        vida07_col = st.text_input("Hp (0.07) DE POR VIDA", value="Hp (0.07) DE POR VIDA")
    with colC2:
        vida3_col  = st.text_input("Hp (3) DE POR VIDA", value="Hp (3) DE POR VIDA")

    def _safe_num(x):
        if isinstance(x, str) and x.strip().upper() == "PM": return 0.0
        try:
            v = float(x);  return 0.0 if pd.isna(v) else v
        except Exception:
            return 0.0

    def _parse_fecha(s):
        if pd.isna(s): return pd.NaT
        for fmt in ["%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"]:
            try: return datetime.strptime(str(s), fmt)
            except Exception: pass
        try: return pd.to_datetime(s, errors="coerce")
        except Exception: return pd.NaT

    def ninox_update_one_by_one(team_id, db_id, table_id, updates):
        url = f"{BASE_URL}/teams/{team_id}/databases/{db_id}/tables/{table_id}/records"
        for i, up in enumerate(updates, 1):
            r = requests.post(url, headers=ninox_headers(), json=[up], timeout=60)
            if r.status_code != 200:
                return {"ok": False, "updated": i-1, "error": f"{r.status_code} {r.text}", "failed_payload": up}
        return {"ok": True, "updated": len(updates)}

    if st.button("🧮 Calcular y actualizar acumulados (modo seguro)"):
        df_rep = st.session_state.get("df_reporte_ninox")
        if df_rep is None or df_rep.empty:
            st.error("Primero trae los datos de REPORTE desde Ninox.")
        else:
            df = df_rep.copy()

            hp10_col = next((c for c in df.columns if c.strip().lower() in {"hp (10)","hp(10)","hp 10"}), None)
            hp07_col = next((c for c in df.columns if c.strip().lower() in {"hp (0.07)","hp(0.07)","hp 0.07"}), None)
            hp3_col  = next((c for c in df.columns if c.strip().lower() in {"hp (3)","hp(3)","hp 3"}), None)
            nombre_col  = next((c for c in df.columns if c.strip().upper() == "NOMBRE"), None)
            cedula_col  = next((c for c in df.columns if c.strip().upper() in {"CÉDULA","CEDULA"}), None)
            fecha_col   = next((c for c in df.columns if c.strip().upper() == "FECHA DE LECTURA"), None)

            if not all([hp10_col, hp07_col, hp3_col, nombre_col, cedula_col, fecha_col]):
                st.error("Faltan columnas clave en REPORTE.")
                st.stop()

            if solo_lote:
                codigo_col  = next((c for c in df.columns if "CÓDIGO" in c.upper() and ("DOSÍMETRO" in c.upper() or "DOSIMETRO" in c.upper())), None)
                if codigo_col and len(codigos_lote) > 0:
                    df = df[df[codigo_col].astype(str).str.upper().isin(codigos_lote)].copy()
                    st.info(f"Aplicado filtro de lote: {len(df)} filas consideradas.")

            df["_hp10"] = df[hp10_col].apply(_safe_num)
            df["_hp07"] = df[hp07_col].apply(_safe_num)
            df["_hp3"]  = df[hp3_col].apply(_safe_num)
            df["_fecha"] = df[fecha_col].apply(_parse_fecha)
            df["_year"]  = df["_fecha"].dt.year

            key = [nombre_col, cedula_col]
            life10 = df.groupby(key)["_hp10"].sum()
            life07 = df.groupby(key)["_hp07"].sum()
            life3  = df.groupby(key)["_hp3"].sum()
            year10 = df[df["_year"] == int(target_year)].groupby(key)["_hp10"].sum()
            year07 = df[df["_year"] == int(target_year)].groupby(key)["_hp07"].sum()
            year3  = df[df["_year"] == int(target_year)].groupby(key)["_hp3"].sum()

            acc = pd.concat([
                life10.rename("life10"), life07.rename("life07"), life3.rename("life3"),
                year10.rename("year10"), year07.rename("year07"), year3.rename("year3")
            ], axis=1).fillna(0.0).reset_index()
            st.caption("Acumulados por persona:"); st.dataframe(acc, use_container_width=True)

            ninox_fields = ninox_get_table_fields(TEAM_ID, DATABASE_ID, report_table_id)
            missing = [x for x in [anual10_col, anual07_col, anual3_col, vida10_col, vida07_col, vida3_col] if x not in ninox_fields]
            if missing:
                st.error("Estos campos no existen en REPORTE (o no coinciden exactamente):\n- " + "\n- ".join(missing))
                st.info("Crea campos tipo **Número** o cambia los nombres arriba para que coincidan.")
                st.stop()

            full_df_ids = st.session_state.df_reporte_ninox.copy()
            if "_id" not in full_df_ids.columns:
                st.error("No tengo los IDs de Ninox. Vuelve a pulsar 'Traer REPORTE desde Ninox'.")
                st.stop()

            updates = []
            for _, r in acc.iterrows():
                n, c = r[nombre_col], r[cedula_col]
                yr10, yr07, yr3 = float(r["year10"]), float(r["year07"]), float(r["year3"])
                lf10, lf07, lf3  = float(r["life10"]), float(r["life07"]), float(r["life3"])
                mask = (full_df_ids[nombre_col] == n) & (full_df_ids[cedula_col] == c)
                ids = full_df_ids.loc[mask, "_id"].dropna().tolist()
                for rid in ids:
                    updates.append({"id": rid, "fields": {
                        anual10_col: yr10, anual07_col: yr07, anual3_col: yr3,
                        vida10_col: lf10, vida07_col: lf07, vida3_col: lf3
                    }})

            if not updates:
                st.warning("No hay filas para actualizar.")
            else:
                with st.spinner(f"Actualizando {len(updates)} filas (modo seguro)..."):
                    res = ninox_update_one_by_one(TEAM_ID, DATABASE_ID, report_table_id, updates)
                if res.get("ok"):
                    st.success(f"✅ Actualizadas {res.get('updated', 0)} filas.")
                    try:
                        df_ref = ninox_fetch_records(TEAM_ID, DATABASE_ID, report_table_id)
                        st.caption("Vista rápida de REPORTE (refrescado):")
                        st.dataframe(df_ref.tail(20), use_container_width=True)
                    except Exception:
                        pass
                else:
                    st.error(f"❌ Error al actualizar: {res.get('error')}")
                    st.caption("Payload que falló:"); st.json(res.get("failed_payload"))

