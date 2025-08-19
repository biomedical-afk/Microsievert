
import re
import io
import requests
import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from dateutil.parser import parse as dtparse
from typing import List, Dict, Any, Optional, Set

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.drawing.image import Image as XLImage

# Logo demo (si no subes uno)
try:
    from PIL import Image as PILImage, ImageDraw, ImageFont
except Exception:
    PILImage = None

# ===================== NINOX CONFIG =====================
API_TOKEN   = "0b3a1130-785a-11f0-ace0-3fb1fcb242e2"
TEAM_ID     = "ihp8o8AaLzfodwc4J"
DATABASE_ID = "ksqzvuts5aq0"
BASE_URL    = "https://api.ninox.com/v1"

# Tabla BASE (personas/asignaciones) y REPORTE (lecturas)
TABLE_BASE_ID   = "E"   # BASE DE DATOS
TABLE_REPORTE_ID = "C"  # REPORTE

# ===================== STREAMLIT =====================
st.set_page_config(page_title="Microsievert — Dosimetría", page_icon="🧪", layout="wide")
st.title("🧪 Microsievert — Gestión y Reportes de Dosimetría")

if "df_valor_control" not in st.session_state:
    st.session_state.df_valor_control = None
if "df_reporte_final" not in st.session_state:
    st.session_state.df_reporte_final = None

# ===================== Helpers Ninox =====================
def ninox_headers() -> Dict[str,str]:
    return {"Authorization": f"Bearer {API_TOKEN}", "Content-Type": "application/json"}

@st.cache_data(ttl=300, show_spinner=False)
def ninox_list_tables(team_id: str, db_id: str):
    url = f"{BASE_URL}/teams/{team_id}/databases/{db_id}/tables"
    r = requests.get(url, headers=ninox_headers(), timeout=30)
    r.raise_for_status()
    return r.json()

@st.cache_data(ttl=300, show_spinner=False)
def ninox_fetch_records(team_id: str, db_id: str, table_id: str, per_page: int = 1000) -> pd.DataFrame:
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
    df.columns = [str(c) for c in df.columns]  # conserva acentos/espacios
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
def ninox_get_table_fields(team_id: str, db_id: str, table_id: str) -> Set[str]:
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

# ===================== Utilidades comunes =====================
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
        if pd.isna(fecha): return per or ""
        meses = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
                 "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
        mes = meses[fecha.month - 1]
        return f"{mes} {fecha.year}"
    except Exception:
        return per or ""

def as_value(v: Any):
    if v is None: return ""
    s = str(v).strip().replace(",", ".")
    if s.upper() == "PM": return "PM"
    try: return float(s)
    except Exception: return s

def as_num(v: Any) -> float:
    if v is None: return 0.0
    s = str(v).strip().replace(",", ".")
    if s == "" or s.upper() == "PM": return 0.0
    try: return float(s)
    except Exception: return 0.0

def round2(x: float) -> float:
    return float(f"{x:.2f}")

# =============== Pestañas ===============
tab1, tab2 = st.tabs(["VALOR − CONTROL + Subida a Ninox", "Reporte Actual/Anual/Vida"])

# ======================================================================================
# TAB 1: VALOR − CONTROL + subida a Ninox
# ======================================================================================
with tab1:
    st.subheader("📤 Cargar archivo de Dosis")
    upload = st.file_uploader("Selecciona CSV/XLS/XLSX", type=["csv","xls","xlsx"])

    def leer_dosis(upload):
        if not upload: return None
        name = upload.name.lower()
        if name.endswith(".csv"):
            try:
                df = pd.read_csv(upload, delimiter=';', engine='python')
            except Exception:
                upload.seek(0); df = pd.read_csv(upload)
        else:
            df = pd.read_excel(upload)
        norm = (df.columns.astype(str).str.strip().str.lower()
                .str.replace(' ', '', regex=False)
                .str.replace('(', '').str.replace(')', '')
                .str.replace('.', '', regex=False))
        df.columns = norm
        if 'dosimeter' not in df.columns:
            for alt in ['dosimetro','codigo','codigodosimetro','codigo_dosimetro']:
                if alt in df.columns:
                    df.rename(columns={alt:'dosimeter'}, inplace=True); break
        for cand in ['hp10dosecorr','hp10dose','hp10']:
            if cand in df.columns: df.rename(columns={cand:'hp10dose'}, inplace=True); break
        for cand in ['hp007dosecorr','hp007dose','hp007']:
            if cand in df.columns: df.rename(columns={cand:'hp0.07dose'}, inplace=True); break
        for cand in ['hp3dosecorr','hp3dose','hp3']:
            if cand in df.columns: df.rename(columns={cand:'hp3dose'}, inplace=True); break
        for k in ['hp10dose','hp0.07dose','hp3dose']:
            if k in df.columns: df[k] = pd.to_numeric(df[k], errors='coerce').fillna(0.0)
            else: df[k] = 0.0
        if 'dosimeter' in df.columns:
            df['dosimeter'] = df['dosimeter'].astype(str).str.strip().str.upper()
        if 'timestamp' in df.columns:
            df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')
        return df

    df_dosis = leer_dosis(upload) if upload else None
    if df_dosis is not None:
        st.caption("Vista previa dosis (normalizada):")
        st.dataframe(df_dosis.head(15), use_container_width=True)

    st.markdown("---")
    st.subheader("📥 Leer BASE DE DATOS (Ninox)")
    with st.expander("Opciones de lectura", expanded=True):
        base_table_id = st.text_input("Table ID BASE DE DATOS", value=TABLE_BASE_ID)
        show_tables = st.checkbox("Mostrar tablas Ninox (debug)", value=False)

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

    # ---- construcción de registros VALOR - CONTROL ----
    def construir_registros(dfp: pd.DataFrame, dfd: pd.DataFrame, periodo_filtro="— TODOS —"):
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
                per_norm = "CONTROL" if re.match(r'^\s*CONTROL\b', per) else re.sub(r'\.+', '.', per).strip()

                pf = (periodo_filtro or "").strip().upper()
                if pf not in ("", "— TODOS —") and per_norm != pf:
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

                # periodo final NO CONTROL
                per_final = periodo_desde_fecha(per_norm, fecha_str)

                registros.append({
                    'PERIODO DE LECTURA': per_final,
                    'COMPAÑÍA': fila.get('COMPAÑÍA',''),
                    'CÓDIGO DE DOSÍMETRO': cod,
                    'NOMBRE': nombre_raw,
                    'CÉDULA': fila.get('CÉDULA',''),
                    'FECHA DE LECTURA': fecha_str,
                    'TIPO DE DOSÍMETRO': 'CE',
                    'Hp(10)': float(r0.get('hp10dose', 0.0)),
                    'Hp(0.07)': float(r0.get('hp0.07dose', 0.0)),
                    'Hp(3)': float(r0.get('hp3dose', 0.0)),
                })
        return registros

    def aplicar_valor_menos_control(registros: List[dict]) -> List[dict]:
        if not registros: return registros
        base10 = float(registros[0]['Hp(10)'])
        base07 = float(registros[0]['Hp(0.07)'])
        base3  = float(registros[0]['Hp(3)'])
        for i, r in enumerate(registros):
            if i == 0:
                r['NOMBRE'] = "CONTROL"
                r['Hp(10)']  = f"{base10:.2f}"
                r['Hp(0.07)'] = f"{base07:.2f}"
                r['Hp(3)']   = f"{base3:.2f}"
            else:
                for key, base in [('Hp(10)', base10), ('Hp(0.07)', base07), ('Hp(3)', base3)]:
                    diff = float(r[key]) - base
                    r[key] = "PM" if diff < 0.005 else f"{diff:.2f}"
        return registros

    # ---- export simple excel (VALOR-CONTROL) ----
    def exportar_excel_valor_control(df_final: pd.DataFrame) -> bytes:
        wb = Workbook(); ws = wb.active; ws.title = "REPORTE DE DOSIS"
        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'),  bottom=Side(style='thin'))
        ws['I1'] = f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y')}"
        ws['I1'].font = Font(size=10, italic=True)
        ws['I1'].alignment = Alignment(horizontal='right', vertical='top')
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

    # ---- Controles de proceso ----
    left, right = st.columns([1,1])
    with left:
        periodo_filtro  = st.text_input("Filtro PERIODO (opcional)", value="— TODOS —")
        nombre_reporte = st.text_input("Nombre archivo (sin extensión)",
                                       value=f"ReporteDosimetria_{datetime.now().strftime('%Y-%m-%d')}")
    with right:
        subir_pm_como_texto = st.checkbox("Subir 'PM' como TEXTO (si Hp son Texto en Ninox)", value=True)
        debug_uno = st.checkbox("Enviar 1 registro (debug)", value=False)

    btn_proc = st.button("✅ Procesar VALOR − CONTROL", type="primary")
    if btn_proc:
        if df_participantes is None or df_participantes.empty:
            st.error("No hay participantes desde Ninox.")
        elif df_dosis is None or df_dosis.empty:
            st.error("No hay datos de dosis.")
        elif 'dosimeter' not in df_dosis.columns:
            st.error("El archivo de dosis debe tener la columna 'dosimeter'.")
        else:
            with st.spinner("Procesando…"):
                registros = construir_registros(df_participantes, df_dosis, periodo_filtro=periodo_filtro)
                if not registros:
                    st.warning("No hay coincidencias DOSÍMETRO ↔ dosis (revisa filtro/códigos).")
                else:
                    registros = aplicar_valor_menos_control(registros)
                    df_final = pd.DataFrame(registros)

                    # Asegurar NOMBRE CONTROL solo para el primero (ya puesto)
                    df_final['PERIODO DE LECTURA'] = (
                        df_final['PERIODO DE LECTURA'].astype(str)
                        .str.replace(r'\.+$', '', regex=True).str.strip()
                    )

                    st.session_state.df_valor_control = df_final
                    st.success(f"¡Listo! Registros generados: {len(df_final)}")
                    st.dataframe(df_final, use_container_width=True)

                    try:
                        xlsx = exportar_excel_valor_control(df_final)
                        st.download_button("⬇️ Descargar Excel (VALOR−CONTROL)", data=xlsx,
                            file_name=f"{(nombre_reporte.strip() or 'ReporteDosimetria')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    except Exception as e:
                        st.error(f"No se pudo generar Excel: {e}")

    # ---- Subir a Ninox (tabla REPORTE) ----
    st.markdown("---")
    st.subheader("⬆️ Subir a Ninox (tabla REPORTE)")
    report_table_id = st.text_input("Table ID REPORTE", value=TABLE_REPORTE_ID)

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
        try: return float(v)
        except Exception: return v if v is not None else None

    def _to_str(v):
        if pd.isna(v): return ""
        if isinstance(v, (pd.Timestamp, )):
            return v.strftime("%Y-%m-%d %H:%M:%S")
        return str(v)

    if st.button("Subir TODO a Ninox (REPORTE)"):
        df_final = st.session_state.df_valor_control
        if df_final is None or df_final.empty:
            st.error("Primero procesa el archivo (botón 'Procesar VALOR − CONTROL').")
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
                    if dest in {"Hp (10)","Hp (0.07)","Hp (3)"}:
                        val = _hp_value(val, as_text_pm=subir_pm_como_texto)
                    else:
                        val = _to_str(val)
                    fields_payload[dest] = val
                rows.append({"fields": fields_payload})

            if debug_uno:
                st.caption("Payload (primer registro):")
                st.json(rows[:1])

            with st.spinner("Subiendo a Ninox…"):
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

# ======================================================================================
# TAB 2: Reporte Actual/Anual/Vida (desde Ninox REPORTE) + Excel plantilla
# ======================================================================================
with tab2:
    st.subheader("📥 Cargar registros desde REPORTE (Ninox)")

    # ----- fetch completo de tabla REPORTE -----
    def headers_api() -> Dict[str,str]:
        return {"Authorization": f"Bearer {API_TOKEN}", "Content-Type": "application/json"}

    def fetch_all_records(table_id: str, page_size: int = 1000):
        url = f"{BASE_URL}/teams/{TEAM_ID}/databases/{DATABASE_ID}/tables/{table_id}/records"
        skip, out = 0, []
        while True:
            r = requests.get(url, headers=headers_api(), params={"limit": page_size, "skip": skip}, timeout=60)
            r.raise_for_status()
            chunk = r.json()
            if not chunk: break
            out.extend(chunk)
            if len(chunk) < page_size: break
            skip += page_size
        return out

    def normalize_df(records):
        rows = []
        for r in records:
            f = r.get("fields", {}) or {}
            rows.append({
                "_id": r.get("id"),
                "PERIODO DE LECTURA": f.get("PERIODO DE LECTURA"),
                "COMPAÑÍA": f.get("COMPAÑÍA"),
                "CÓDIGO DE DOSÍMETRO": str(f.get("CÓDIGO DE DOSÍMETRO") or "").strip(),
                "NOMBRE": f.get("NOMBRE"),
                "CÉDULA": f.get("CÉDULA"),
                "FECHA DE LECTURA": f.get("FECHA DE LECTURA"),
                "TIPO DE DOSÍMETRO": f.get("TIPO DE DOSÍMETRO"),
                "Hp10_RAW":  as_value(f.get("Hp (10)")),
                "Hp007_RAW": as_value(f.get("Hp (0.07)")),
                "Hp3_RAW":  as_value(f.get("Hp (3)")),
                "Hp10_NUM":  as_num(f.get("Hp (10)")),
                "Hp007_NUM": as_num(f.get("Hp (0.07)")),
                "Hp3_NUM":  as_num(f.get("Hp (3)")),
            })
        df = pd.DataFrame(rows)
        df["FECHA_DE_LECTURA_DT"] = pd.to_datetime(
            df["FECHA DE LECTURA"].apply(
                lambda x: dtparse(str(x), dayfirst=True) if pd.notna(x) and str(x).strip() != "" else pd.NaT
            ), errors="coerce"
        )
        return df

    with st.spinner("Cargando datos desde Ninox…"):
        base_rep = normalize_df(fetch_all_records(TABLE_REPORTE_ID))

    if base_rep.empty:
        st.warning("No hay registros en la tabla REPORTE.")
        st.stop()

    # --------- filtros en la propia pestaña ---------
    left, right = st.columns([1,1])
    with left:
        files = st.file_uploader("Archivos de dosis para filtrar por CÓDIGO (opcional)", type=["csv","xlsx","xls"], accept_multiple_files=True)
    with right:
        header_line1 = st.text_input("Encabezado línea 1", "MICROSIEVERT, S.A.")
        header_line2 = st.text_input("Línea 2", "PH Conardo")
        header_line3 = st.text_input("Línea 3", "Calle 41 Este, Panamá")
        header_line4 = st.text_input("Línea 4", "PANAMÁ")
        logo_file = st.file_uploader("Logo (PNG/JPG) opcional", type=["png","jpg","jpeg"])

    # Códigos desde archivos
    def read_codes_from_files(files) -> Set[str]:
        codes: Set[str] = set()
        for f in files or []:
            raw = f.read(); f.seek(0)
            name = f.name.lower()
            try:
                if name.endswith((".xlsx",".xls")):
                    df = pd.read_excel(BytesIO(raw))
                else:
                    df = None
                    for enc in ("utf-8-sig","latin-1"):
                        try:
                            df = pd.read_csv(BytesIO(raw), sep=None, engine="python", encoding=enc); break
                        except Exception: continue
                    if df is None: df = pd.read_csv(BytesIO(raw))
            except Exception:
                continue
            if df is None or df.empty: continue
            cand = None
            for c in df.columns:
                cl = str(c).lower()
                if any(k in cl for k in ["dosim","código","codigo","wb","dosímetro","dosimetro"]):
                    cand = c; break
            if cand is None:
                for c in df.columns:
                    if df[c].astype(str).str.contains(r"^WB\d{5,}$", case=False, na=False).any():
                        cand = c; break
            if cand is None: cand = df.columns[0]
            codes |= set(df[cand].astype(str).str.strip())
        return {c for c in codes if c and c.lower() != "nan"}

    codes_filter: Optional[Set[str]] = read_codes_from_files(files) if files else None
    if codes_filter:
        st.success(f"Códigos detectados: {len(codes_filter)}")

    per_order = (base_rep.groupby("PERIODO DE LECTURA")["FECHA_DE_LECTURA_DT"].max()
                 .sort_values(ascending=False).index.astype(str).tolist())
    per_valid = [p for p in per_order if p.strip().upper() != "CONTROL"]
    periodo_actual = st.selectbox("Periodo actual", per_valid, index=0 if per_valid else None)
    periodos_anteriores = st.multiselect(
        "Periodos anteriores (para ANUAL)",
        [p for p in per_valid if p != periodo_actual],
        default=[per_valid[1]] if len(per_valid) > 1 else []
    )

    comp_opts = ["(todas)"] + sorted(base_rep["COMPAÑÍA"].dropna().astype(str).unique().tolist())
    compania = st.selectbox("Compañía", comp_opts, index=0)
    tipo_opts = ["(todos)"] + sorted(base_rep["TIPO DE DOSÍMETRO"].dropna().astype(str).unique().tolist())
    tipo = st.selectbox("Tipo de dosímetro", tipo_opts, index=0)

    df = base_rep.copy()
    if codes_filter: df = df[df["CÓDIGO DE DOSÍMETRO"].isin(codes_filter)]
    if compania != "(todas)": df = df[df["COMPAÑÍA"].astype(str) == compania]
    if tipo != "(todos)": df = df[df["TIPO DE DOSÍMETRO"].astype(str) == tipo]
    if df.empty:
        st.warning("No hay registros que cumplan el filtro.")
        st.stop()

    control_codes = set(df.loc[df["NOMBRE"].astype(str).str.strip().str.upper()=="CONTROL",
                               "CÓDIGO DE DOSÍMETRO"].unique())

    def ultimo_en_periodo(g: pd.DataFrame, periodo: str) -> pd.Series:
        x = g[g["PERIODO DE LECTURA"].astype(str) == str(periodo)].sort_values("FECHA_DE_LECTURA_DT", ascending=False)
        return x.iloc[0] if not x.empty else pd.Series(dtype="object")

    rows = []
    for code, sub in df.groupby("CÓDIGO DE DOSÍMETRO", as_index=False):
        ult = ultimo_en_periodo(sub, periodo_actual)
        if ult.empty: continue
        rows.append({
            "CÓDIGO DE DOSÍMETRO": code,
            "PERIODO DE LECTURA": periodo_actual,
            "COMPAÑÍA": ult.get("COMPAÑÍA"),
            "NOMBRE": ult.get("NOMBRE"),
            "CÉDULA": ult.get("CÉDULA"),
            "FECHA Y HORA DE LECTURA": ult.get("FECHA DE LECTURA"),
            "TIPO DE DOSÍMETRO": ult.get("TIPO DE DOSÍMETRO"),
            "Hp10_ACTUAL_RAW":  ult.get("Hp10_RAW"),
            "Hp007_ACTUAL_RAW": ult.get("Hp007_RAW"),
            "Hp3_ACTUAL_RAW":   ult.get("Hp3_RAW"),
            "Hp10_ACTUAL_NUM":  ult.get("Hp10_NUM", 0.0),
            "Hp007_ACTUAL_NUM": ult.get("Hp007_NUM", 0.0),
            "Hp3_ACTUAL_NUM":   ult.get("Hp3_NUM", 0.0),
        })
    df_actual = pd.DataFrame(rows)

    df_prev = df[df["PERIODO DE LECTURA"].astype(str).isin(periodos_anteriores)]
    prev_sum = (df_prev.groupby("CÓDIGO DE DOSÍMETRO")[["Hp10_NUM","Hp007_NUM","Hp3_NUM"]]
                .sum().rename(columns={"Hp10_NUM":"Hp10_ANT_SUM","Hp007_NUM":"Hp007_ANT_SUM","Hp3_NUM":"Hp3_ANT_SUM"}))

    vida_sum = (df.groupby("CÓDIGO DE DOSÍMETRO")[["Hp10_NUM","Hp007_NUM","Hp3_NUM"]]
                .sum().rename(columns={"Hp10_NUM":"Hp10_VIDA_NUM","Hp007_NUM":"Hp007_VIDA_NUM","Hp3_NUM":"Hp3_VIDA_NUM"}))
    vida_raw = (df.groupby("CÓDIGO DE DOSÍMETRO")[["Hp10_RAW","Hp007_RAW","Hp3_RAW"]]
                .agg(list).rename(columns={"Hp10_RAW":"Hp10_VIDA_RAW","Hp007_RAW":"Hp007_VIDA_RAW","Hp3_RAW":"Hp3_VIDA_RAW"}))

    out = (df_actual.set_index("CÓDIGO DE DOSÍMETRO")
           .join(prev_sum, how="left").join(vida_sum, how="left").join(vida_raw, how="left")).reset_index()

    for c in ["Hp10_ANT_SUM","Hp007_ANT_SUM","Hp3_ANT_SUM","Hp10_VIDA_NUM","Hp007_VIDA_NUM","Hp3_VIDA_NUM"]:
        if c not in out: out[c] = 0.0
        out[c] = out[c].fillna(0.0)

    def pm_or_sum(raws: List[Any], numeric_sum: float) -> Any:
        vals = [str(x).upper() for x in raws if str(x).strip()!=""]
        if vals and all(v == "PM" for v in vals): return "PM"
        return round2(numeric_sum)

    def show_raw_or_num(raw): return raw if str(raw).upper()=="PM" else round2(float(raw))

    out["Hp (10) ACTUAL"]   = out["Hp10_ACTUAL_RAW"].apply(show_raw_or_num)
    out["Hp (0.07) ACTUAL"] = out["Hp007_ACTUAL_RAW"].apply(show_raw_or_num)
    out["Hp (3) ACTUAL"]    = out["Hp3_ACTUAL_RAW"].apply(show_raw_or_num)

    out["Hp (10) ANUAL"]   = out.apply(lambda r: pm_or_sum([r["Hp10_ACTUAL_RAW"]], r["Hp10_ACTUAL_NUM"] + r["Hp10_ANT_SUM"]), axis=1)
    out["Hp (0.07) ANUAL"] = out.apply(lambda r: pm_or_sum([r["Hp007_ACTUAL_RAW"]], r["Hp007_ACTUAL_NUM"] + r["Hp007_ANT_SUM"]), axis=1)
    out["Hp (3) ANUAL"]    = out.apply(lambda r: pm_or_sum([r["Hp3_ACTUAL_RAW"]],  r["Hp3_ACTUAL_NUM"]  + r["Hp3_ANT_SUM"]), axis=1)

    out["Hp (10) VIDA"]   = out.apply(lambda r: pm_or_sum(r.get("Hp10_VIDA_RAW", []) or [], r["Hp10_VIDA_NUM"]), axis=1)
    out["Hp (0.07) VIDA"] = out.apply(lambda r: pm_or_sum(r.get("Hp007_VIDA_RAW", []) or [], r["Hp007_VIDA_NUM"]), axis=1)
    out["Hp (3) VIDA"]    = out.apply(lambda r: pm_or_sum(r.get("Hp3_VIDA_RAW", []) or [],  r["Hp3_VIDA_NUM"]), axis=1)

    out["__is_control"] = out["CÓDIGO DE DOSÍMETRO"].isin(control_codes)
    out = out.sort_values(["__is_control","CÓDIGO DE DOSÍMETRO"], ascending=[False, True])

    FINAL_COLS = [
        "PERIODO DE LECTURA","COMPAÑÍA","CÓDIGO DE DOSÍMETRO","NOMBRE","CÉDULA",
        "FECHA Y HORA DE LECTURA","TIPO DE DOSÍMETRO",
        "Hp (10) ACTUAL","Hp (0.07) ACTUAL","Hp (3) ACTUAL",
        "Hp (10) ANUAL","Hp (0.07) ANUAL","Hp (3) ANUAL",
        "Hp (10) VIDA","Hp (0.07) VIDA","Hp (3) VIDA",
    ]
    for c in FINAL_COLS:
        if c not in out.columns: out[c] = ""
    out = out[FINAL_COLS]

    st.subheader("Reporte final (vista previa)")
    st.dataframe(out, use_container_width=True, hide_index=True)

    # --------- Descargas simples ---------
    csv_bytes = out.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Descargar CSV (UTF-8 con BOM)", data=csv_bytes,
                       file_name=f"reporte_dosimetria_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                       mime="text/csv")

    def to_excel_simple(df: pd.DataFrame, sheet_name="Reporte"):
        bio = BytesIO()
        with pd.ExcelWriter(bio, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name=sheet_name)
        bio.seek(0); return bio.getvalue()

    xlsx_simple = to_excel_simple(out)
    st.download_button("⬇️ Descargar Excel (tabla simple)", data=xlsx_simple,
                       file_name=f"reporte_dosimetria_tabla_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ---------- Helpers (logo, medidas) ----------
    def col_pixels(ws, col_letter: str) -> int:
        w = ws.column_dimensions[col_letter].width
        if w is None: w = 8.43
        return int(w * 7 + 5)

    def row_pixels(ws, row_idx: int) -> int:
        h = ws.row_dimensions[row_idx].height
        if h is None: h = 15
        return int(h * 96 / 72)

    def fit_logo(ws, logo_bytes: bytes, top_left: str = "C1", bottom_right: str = "F4", padding: int = 6):
        if not logo_bytes: return
        img = XLImage(BytesIO(logo_bytes))
        tl_col = column_index_from_string(''.join([c for c in top_left if c.isalpha()]))
        tl_row = int(''.join([c for c in top_left if c.isdigit()]))
        br_col = column_index_from_string(''.join([c for c in bottom_right if c.isalpha()]))
        br_row = int(''.join([c for c in bottom_right if c.isdigit()]))
        box_w = sum(col_pixels(ws, get_column_letter(c)) for c in range(tl_col, br_col + 1))
        box_h = sum(row_pixels(ws, r) for r in range(tl_row, br_row + 1))
        max_w = max(10, box_w - 2*padding); max_h = max(10, box_h - 2*padding)
        scale = min(max_w / img.width, max_h / img.height, 1.0)
        img.width = int(img.width * scale); img.height = int(img.height * scale)
        img.anchor = top_left; ws.add_image(img)

    def sample_logo_bytes(text="µSv  MICROSIEVERT, S.A."):
        if PILImage is None: return None
        img = PILImage.new("RGBA", (420, 110), (255, 255, 255, 0))
        d = ImageDraw.Draw(img)
        try: font = ImageFont.truetype("arial.ttf", 36)
        except Exception: font = ImageFont.load_default()
        d.text((12, 30), text, fill=(0, 70, 140, 255), font=font)
        bio = BytesIO(); img.save(bio, format="PNG"); return bio.getvalue()

    # ---------- Excel “formato plantilla” ----------
    def build_formatted_excel(df_final: pd.DataFrame,
                              header_lines: List[str],
                              logo_bytes: Optional[bytes]) -> bytes:
        wb = Workbook(); ws = wb.active; ws.title = "Reporte"
        bold = Font(bold=True)
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin = Side(style="thin"); border = Border(top=thin, bottom=thin, left=thin, right=thin)
        gray = PatternFill("solid", fgColor="DDDDDD"); group_fill = PatternFill("solid", fgColor="EEEEEE")

        widths = {"A":24,"B":28,"C":16,"D":16,"E":16,"F":16,"G":10,
                  "H":12,"I":12,"J":12,"K":12,"L":12,"M":12,"N":12,"O":12,"P":12}
        for k,v in widths.items(): ws.column_dimensions[k].width = v
        for r in range(1,5): ws.row_dimensions[r].height = 20

        for i, line in enumerate(header_lines[:4], start=1):
            ws.merge_cells(f"A{i}:B{i}"); c = ws[f"A{i}"]; c.value = line; c.fill = gray
            c.font = Font(bold=True); c.alignment = Alignment(horizontal="left", vertical="center")
            for col in ("A","B"): ws.cell(row=i, column=ord(col)-64).border = border

        ws.merge_cells("I1:J1"); ws["I1"] = "Fecha de emisión"
        ws["I1"].font = Font(bold=True, size=10); ws["I1"].alignment = center; ws["I1"].fill = gray
        ws.merge_cells("K1:P1"); ws["K1"] = datetime.now().strftime("%d-%b-%y").lower()
        ws["K1"].font = Font(bold=True, size=10); ws["K1"].alignment = center
        for col_idx in range(ord("I")-64, ord("P")-64+1):
            ws.cell(row=1, column=col_idx).border = border

        if logo_bytes is None:
            logo_bytes = sample_logo_bytes()
        if logo_bytes:
            fit_logo(ws, logo_bytes, top_left="C1", bottom_right="F4", padding=6)

        ws.merge_cells("A6:P6"); ws["A6"] = "REPORTE DE DOSIMETRÍA"
        ws["A6"].font = Font(bold=True, size=14); ws["A6"].alignment = center

        ws.merge_cells("H7:J7"); ws["H7"] = "DOSIS ACTUAL (mSv)"
        ws.merge_cells("K7:M7"); ws["K7"] = "DOSIS ANUAL (mSv)"
        ws.merge_cells("N7:P7"); ws["N7"] = "DOSIS DE POR VIDA (mSv)"
        for rng in ("H7","K7","N7"):
            ws[rng].font = bold; ws[rng].alignment = center
            # cerrar bordes de toda la franja
            start_col = column_index_from_string(rng[0]); end_col = start_col + 2
            for col in range(start_col, end_col+1):
                ws.cell(row=7, column=col).border = border
                ws.cell(row=7, column=col).fill = group_fill

        headers = [
            "PERIODO DE LECTURA","COMPAÑÍA","CÓDIGO DE DOSÍMETRO","NOMBRE","CÉDULA",
            "FECHA Y HORA DE LECTURA","TIPO DE DOSÍMETRO",
            "Hp (10) ACTUAL","Hp (0.07) ACTUAL","Hp (3) ACTUAL",
            "Hp (10) ANUAL","Hp (0.07) ANUAL","Hp (3) ANUAL",
            "Hp (10) VIDA","Hp (0.07) VIDA","Hp (3) VIDA",
        ]
        header_row = 8
        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=header_row, column=col_idx, value=h)
            cell.font = bold; cell.alignment = center; cell.border = border; cell.fill = gray

        start_row = header_row + 1
        for _, r in df_final[headers].iterrows(): ws.append(list(r.values))
        last_row = ws.max_row

        for row in ws.iter_rows(min_row=header_row, max_row=last_row, min_col=1, max_col=len(headers)):
            for c in row:
                c.border = border
                if c.row >= start_row:
                    c.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)
        for rr in range(start_row, last_row + 1): ws.row_dimensions[rr].height = 20

        ws.freeze_panes = f"A{start_row}"

        for col_cells in ws.iter_cols(min_col=1, max_col=16, min_row=header_row, max_row=last_row):
            col_letter = get_column_letter(col_cells[0].column)
            max_len = max(len("" if c.value is None else str(c.value)) for c in col_cells)
            ws.column_dimensions[col_letter].width = max(ws.column_dimensions[col_letter].width, min(max_len + 2, 42))

        # --- Sección informativa (resumida) ---
        row = last_row + 2
        ws.merge_cells(f"A{row}:P{row}"); ws[f"A{row}"] = "INFORMACIÓN DEL REPORTE DE DOSIMETRÍA"
        ws[f"A{row}"].font = Font(bold=True); ws[f"A{row}"].alignment = Alignment(horizontal="center"); row += 1

        bullets = [
            "‒ Periodo de lectura: periodo de uso del dosímetro personal.",
            "‒ Fecha de lectura: fecha en que se realizó la lectura.",
            "‒ Tipo de dosímetro:",
        ]
        for text in bullets:
            ws.merge_cells(f"A{row}:D{row}"); c = ws[f"A{row}"]; c.value = text
            c.font = Font(size=10, bold=True); c.alignment = Alignment(horizontal="left"); row += 2

        tipos = [("CE","Cuerpo Entero"), ("A","Anillo"), ("B","Brazalete"), ("CR","Cristalino")]
        border_box = Border(top=thin, bottom=thin, left=thin, right=thin)
        for clave, desc in tipos:
            ws.merge_cells(f"C{row}:D{row}"); ws[f"C{row}"] = f"{clave} = {desc}"
            ws[f"C{row}"].font = Font(size=10, bold=True); ws[f"C{row}"].alignment = Alignment(horizontal="left")
            for col in ("C","D"): ws.cell(row=row, column=ord(col)-64).border = border_box
            row += 1
        row += 1

        ws.merge_cells(f"A{row}:P{row}"); ws[f"A{row}"] = (
            "POR DEBAJO DEL MÍNIMO DETECTADO: las dosis por debajo del mínimo se reportan como “PM”."
        )



