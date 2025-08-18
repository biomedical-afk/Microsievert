import io
import re
import requests
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# ===================== CONFIG NINOX =====================
API_TOKEN   = "0b3a1130-785a-11f0-ace0-3fb1fcb242e2"  # tu API key
TEAM_ID     = "ihp8o8AaLzfodwc4J"
DATABASE_ID = "ksqzvuts5aq0"
REPORT_TABLE_ID = "C"  # Tabla REPORTE (ID)

BASE_URL = "https://api.ninox.com/v1"
HEADERS  = {"Authorization": f"Bearer {API_TOKEN}", "Content-Type": "application/json"}

# ===================== STREAMLIT BASE =====================
st.set_page_config(page_title="Microsievert - REPORTE desde Ninox", page_icon="🧪", layout="wide")
st.title("🧪 Reporte de Dosimetría — Datos de Ninox (REPORTE)")
st.caption("El archivo de dosis se usa ÚNICAMENTE para filtrar; los valores Hp* provienen de Ninox.")

# ===================== HELPERS NINOX =====================
@st.cache_data(ttl=300, show_spinner=False)
def ninox_fetch_records_with_id(team_id: str, db_id: str, table_id: str, per_page: int = 1000) -> pd.DataFrame:
    """
    Descarga registros de Ninox: retorna DataFrame con 'id' + fields().
    """
    url = f"{BASE_URL}/teams/{team_id}/databases/{db_id}/tables/{table_id}/records"
    out, offset = [], 0
    while True:
        r = requests.get(url, headers=HEADERS, params={"perPage": per_page, "offset": offset}, timeout=60)
        if r.status_code != 200:
            raise RuntimeError(f"{r.status_code} {r.text}")
        batch = r.json()
        if not batch:
            break
        out.extend(batch)
        if len(batch) < per_page:
            break
        offset += per_page
    rows = []
    for rec in out:
        row = {"id": rec.get("id")}
        row.update(rec.get("fields", {}))
        rows.append(row)
    df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["id"])
    # Normalizar nombres a str (con acentos/espacios intactos)
    df.columns = [str(c) for c in df.columns]
    return df

def excel_bytes_from_df(df: pd.DataFrame, sheet_name="Reporte"):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'),  bottom=Side(style='thin'))

    # Encabezados
    for j, h in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=j, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill('solid', fgColor='DDDDDD')
        cell.border = border

    # Datos
    for i, (_, row) in enumerate(df.iterrows(), start=2):
        for j, val in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=val)
            c.alignment = Alignment(horizontal='center', wrap_text=True)
            c.font = Font(size=10)
            c.border = border

    # Ancho columnas
    for col in ws.columns:
        mx = max(len(str(c.value)) if c.value else 0 for c in col) + 2
        ws.column_dimensions[get_column_letter(col[0].column)].width = mx

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.read()

# ===================== LECTURA REPORTE DE NINOX =====================
st.subheader("📥 Leer tabla REPORTE desde Ninox")
try:
    df_rep = ninox_fetch_records_with_id(TEAM_ID, DATABASE_ID, REPORT_TABLE_ID)
    if df_rep.empty:
        st.warning("La tabla REPORTE (id C) está vacía.")
        st.stop()
    st.success(f"Leídos {len(df_rep)} registros de REPORTE.")
    with st.expander("Ver primeras filas de REPORTE"):
        st.dataframe(df_rep.head(20), use_container_width=True)
except Exception as e:
    st.error(f"❌ Error leyendo REPORTE de Ninox: {e}")
    st.stop()

# ===================== ARCHIVO DE DOSIS SOLO PARA FILTRAR =====================
st.subheader("📂 Archivo de dosis para FILTRAR (CSV/XLS/XLSX)")
up = st.file_uploader("Selecciona archivo (se usará solo para obtener NOMBRE+CÉDULA o dosímetros de referencia)",
                      type=["csv", "xls", "xlsx"])

def leer_archivo_filtro(upload):
    if upload is None:
        return pd.DataFrame()
    try:
        if upload.name.lower().endswith(".csv"):
            df = pd.read_csv(upload)
        else:
            df = pd.read_excel(upload)
        # Normalizar headers a texto plano
        df.columns = [str(c) for c in df.columns]
        return df
    except Exception as e:
        st.error(f"No se pudo leer el archivo de filtro: {e}")
        return pd.DataFrame()

df_filtro = leer_archivo_filtro(up)

# ===================== FILTRADO =====================
st.subheader("🎯 Criterio de filtrado")
modo = st.radio(
    "¿Cómo filtrar los registros a incluir en el reporte?",
    options=["AUTO (intentar NOMBRE+CÉDULA, si no, por dosímetro)", "Por NOMBRE+CÉDULA", "Por dosímetro"],
    index=0
)

# columnas esperadas en REPORTE
COLS_REP = {
    "per": "PERIODO DE LECTURA",
    "cia": "COMPAÑÍA",
    "cod": "CÓDIGO DE DOSÍMETRO",
    "nom": "NOMBRE",
    "ced": "CÉDULA",
    "fec": "FECHA DE LECTURA",
    "tipo":"TIPO DE DOSÍMETRO",
    "hp10":"Hp (10)",
    "hp07":"Hp (0.07)",
    "hp3":"Hp (3)",
}

# Función para construir filtro de claves
def construir_claves_filtrado(df_rep_local: pd.DataFrame, df_filtro_local: pd.DataFrame, modo_sel: str):
    """
    Devuelve set de CLAVES (NOMBRE+CEDULA) a mantener.
    Si no hay NOMBRE/CÉDULA en el archivo y el modo lo permite, intenta por dosímetro -> mapea a NOMBRE/CÉDULA desde REPORTE.
    """
    rep = df_rep_local.copy()

    # Asegurar columnas clave en REPORTE
    for need in [COLS_REP["nom"], COLS_REP["ced"], COLS_REP["cod"]]:
        if need not in rep.columns:
            rep[need] = ""

    rep["CLAVE"] = rep[COLS_REP["nom"]].astype(str).str.strip() + "_" + rep[COLS_REP["ced"]].astype(str).str.strip()

    # 1) Intentar por NOMBRE + CÉDULA
    def claves_por_nombre_cedula(df):
        if "NOMBRE" in df.columns and "CÉDULA" in df.columns:
            df2 = df.copy()
            df2["CLAVE"] = df2["NOMBRE"].astype(str).str.strip() + "_" + df2["CÉDULA"].astype(str).str.strip()
            return set(df2["CLAVE"].dropna().astype(str))
        # alternativas de encabezados
        alt_nom = None
        alt_ced = None
        for c in df.columns:
            lc = c.lower()
            if alt_nom is None and ("nombre" in lc):
                alt_nom = c
            if alt_ced is None and ("cédula" in lc or "cedula" in lc or "id" == lc):
                alt_ced = c
        if alt_nom and alt_ced:
            df2 = df.copy()
            df2["CLAVE"] = df2[alt_nom].astype(str).str.strip() + "_" + df2[alt_ced].astype(str).str.strip()
            return set(df2["CLAVE"].dropna().astype(str))
        return set()

    # 2) Por dosímetro (tomar dosimeter del archivo -> mapear a nombre/cedula desde REPORTE)
    def claves_por_dosimetro(df, rep_local):
        # buscar una columna similar a 'dosimeter'
        col_dos = None
        for c in df.columns:
            lc = c.lower().strip().replace(" ", "")
            if lc in {"dosimeter","dosimetro","codigo","codigodosimetro","codigo_dosimetro"}:
                col_dos = c
                break
        if col_dos is None:
            return set()
        # normalizar a uppercase
        lista = df[col_dos].dropna().astype(str).str.strip().str.upper().unique().tolist()
        # obtener NOMBRE+CÉDULA de REPORTE para esos código(s)
        rep_local = rep_local.copy()
        rep_local[COLS_REP["cod"]] = rep_local[COLS_REP["cod"]].astype(str).str.strip().str.upper()
        claves = rep_local.loc[rep_local[COLS_REP["cod"]].isin(lista), "CLAVE"].dropna().astype(str)
        return set(claves)

    # determinar según "modo"
    claves = set()
    if modo_sel == "Por NOMBRE+CÉDULA":
        claves = claves_por_nombre_cedula(df_filtro_local)
    elif modo_sel == "Por dosímetro":
        claves = claves_por_dosimetro(df_filtro_local, rep)
    else:  # AUTO
        claves = claves_por_nombre_cedula(df_filtro_local)
        if not claves:
            claves = claves_por_dosimetro(df_filtro_local, rep)

    return claves

# ===================== CÁLCULO DEL REPORTE =====================
def to_number_preserving_pm(series: pd.Series) -> pd.Series:
    """
    Convierte a número; 'PM' (o texto) -> 0 para sumar.
    Devuelve Serie numérica para cálculos.
    """
    def _conv(x):
        if isinstance(x, str) and x.strip().upper() == "PM":
            return 0.0
        try:
            return float(x)
        except Exception:
            return 0.0
    return series.apply(_conv)

def construir_reporte_final(df_rep_local: pd.DataFrame, claves_keep: set) -> pd.DataFrame:
    rep = df_rep_local.copy()
    # asegurar columnas
    for need in COLS_REP.values():
        if need not in rep.columns:
            rep[need] = ""

    # clave
    rep["CLAVE"] = rep[COLS_REP["nom"]].astype(str).str.strip() + "_" + rep[COLS_REP["ced"]].astype(str).str.strip()

    # filtrar por CLAVE (si no hay claves, no filtramos)
    if claves_keep:
        rep = rep[rep["CLAVE"].isin(claves_keep)].copy()

    if rep.empty:
        return pd.DataFrame(columns=[
            COLS_REP["per"], COLS_REP["cia"], COLS_REP["cod"], COLS_REP["nom"], COLS_REP["ced"],
            COLS_REP["fec"], COLS_REP["tipo"], COLS_REP["hp10"], COLS_REP["hp07"], COLS_REP["hp3"],
            "Hp (10) ANUAL","Hp (0.07) ANUAL","Hp (3) ANUAL",
            "Hp (10) DE POR VIDA","Hp (0.07) DE POR VIDA","Hp (3) DE POR VIDA"
        ])

    # columnas numéricas para sumar (PM -> 0)
    rep["_hp10_num"] = to_number_preserving_pm(rep[COLS_REP["hp10"]])
    rep["_hp07_num"] = to_number_preserving_pm(rep[COLS_REP["hp07"]])
    rep["_hp3_num"]  = to_number_preserving_pm(rep[COLS_REP["hp3"]])

    # ANUAL: por PERIODO + DOSÍMETRO + persona
    anual = rep.groupby([COLS_REP["per"], COLS_REP["cod"], "CLAVE"], as_index=False).agg({
        "_hp10_num": "sum",
        "_hp07_num": "sum",
        "_hp3_num":  "sum"
    }).rename(columns={
        "_hp10_num": "Hp (10) ANUAL",
        "_hp07_num": "Hp (0.07) ANUAL",
        "_hp3_num":  "Hp (3) ANUAL",
    })

    # DE POR VIDA: por DOSÍMETRO + persona (todos los períodos)
    vida = rep.groupby([COLS_REP["cod"], "CLAVE"], as_index=False).agg({
        "_hp10_num": "sum",
        "_hp07_num": "sum",
        "_hp3_num":  "sum"
    }).rename(columns={
        "_hp10_num": "Hp (10) DE POR VIDA",
        "_hp07_num": "Hp (0.07) DE POR VIDA",
        "_hp3_num":  "Hp (3) DE POR VIDA",
    })

    # Merge a nivel de fila
    out = rep.merge(anual, on=[COLS_REP["per"], COLS_REP["cod"], "CLAVE"], how="left") \
             .merge(vida,  on=[COLS_REP["cod"], "CLAVE"],                 how="left")

    # Armar columnas finales (mantener valores Hp originales de la fila)
    columnas_finales = [
        COLS_REP["per"], COLS_REP["cia"], COLS_REP["cod"], COLS_REP["nom"], COLS_REP["ced"],
        COLS_REP["fec"], COLS_REP["tipo"], COLS_REP["hp10"], COLS_REP["hp07"], COLS_REP["hp3"],
        "Hp (10) ANUAL","Hp (0.07) ANUAL","Hp (3) ANUAL",
        "Hp (10) DE POR VIDA","Hp (0.07) DE POR VIDA","Hp (3) DE POR VIDA"
    ]
    out = out[columnas_finales].copy()

    # Orden y tipos: ANUAL / VIDA como float con 2 decimales
    for c in ["Hp (10) ANUAL","Hp (0.07) ANUAL","Hp (3) ANUAL",
              "Hp (10) DE POR VIDA","Hp (0.07) DE POR VIDA","Hp (3) DE POR VIDA"]:
        out[c] = out[c].astype(float).round(2)

    return out

# ===================== UI: PROCESAR Y DESCARGAR =====================
st.markdown("---")
col1, col2 = st.columns([1,1])
with col1:
    nombre_archivo = st.text_input(
        "Nombre del archivo de salida (sin extensión)",
        value=f"Reporte_Dosimetria_{datetime.now().strftime('%Y-%m-%d')}"
    )
with col2:
    boton = st.button("✅ Generar Reporte", type="primary", use_container_width=True)

if boton:
    try:
        claves = construir_claves_filtrado(df_rep, df_filtro, modo)
        if not claves:
            st.info("No se detectaron claves de filtrado. Se generará el reporte con **todos** los registros de REPORTE.")
        df_final = construir_reporte_final(df_rep, claves)

        if df_final.empty:
            st.warning("No hay filas resultantes tras aplicar el filtro.")
        else:
            st.success(f"✅ Reporte generado con {len(df_final)} filas.")
            st.dataframe(df_final, use_container_width=True)

            # Descargas
            csv_bytes  = df_final.to_csv(index=False).encode("utf-8-sig")
            xlsx_bytes = excel_bytes_from_df(df_final, sheet_name="REPORTE")

            st.download_button(
                "⬇️ Descargar CSV",
                data=csv_bytes,
                file_name=f"{(nombre_archivo.strip() or 'Reporte_Dosimetria')}.csv",
                mime="text/csv",
                use_container_width=True
            )
            st.download_button(
                "⬇️ Descargar Excel",
                data=xlsx_bytes,
                file_name=f"{(nombre_archivo.strip() or 'Reporte_Dosimetria')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    except Exception as e:
        st.error(f"❌ Error al generar el reporte: {e}")



