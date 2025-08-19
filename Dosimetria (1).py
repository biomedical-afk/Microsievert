# app.py — Ninox + filtro por archivos de dosis → Reporte Actual / Anual / Vida
# pip install streamlit pandas requests openpyxl python-dateutil

import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime
from dateutil.parser import parse as dtparse
from typing import List, Dict, Any, Optional, Set

# ======= CREDENCIALES NINOX =======
API_TOKEN   = "0b3a1130-785a-11f0-ace0-3fb1fcb242e2"
TEAM_ID     = "ihp8o8AaLzfodwc4J"
DATABASE_ID = "ksqzvuts5aq0"
BASE_URL    = "https://api.ninox.com/v1"
TABLE_ID    = "C"   # ID de la tabla REPORTE
# ==================================

st.set_page_config(page_title="Reporte de Dosimetría (Ninox)", layout="wide")
st.title("Reporte de Dosimetría — Actual, Anual y de por Vida")

# ----------------- Utils -----------------
def headers() -> Dict[str, str]:
    return {"Authorization": f"Bearer {API_TOKEN}", "Content-Type": "application/json"}

def as_float(v: Any) -> float:
    if v is None:
        return 0.0
    s = str(v).strip().replace(",", ".")
    if s == "" or s.upper() == "PM":
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0

def round2(x: float) -> float:
    return float(f"{x:.2f}")

def fetch_all_records(table_id: str, page_size: int = 1000) -> List[Dict[str, Any]]:
    url = f"{BASE_URL}/teams/{TEAM_ID}/databases/{DATABASE_ID}/tables/{table_id}/records"
    skip, out = 0, []
    while True:
        r = requests.get(url, headers=headers(), params={"limit": page_size, "skip": skip}, timeout=60)
        r.raise_for_status()
        chunk = r.json()
        if not chunk:
            break
        out.extend(chunk)
        if len(chunk) < page_size:
            break
        skip += page_size
    return out

def normalize_df(records: List[Dict[str, Any]]) -> pd.DataFrame:
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
            "Hp (10)": as_float(f.get("Hp (10)")),
            "Hp (0.07)": as_float(f.get("Hp (0.07)")),
            "Hp (3)": as_float(f.get("Hp (3)")),
        })
    df = pd.DataFrame(rows)
    # fecha a datetime (acepta 08/08/2025 11:26, etc.)
    if "FECHA DE LECTURA" in df.columns:
        df["FECHA_DE_LECTURA_DT"] = pd.to_datetime(
            df["FECHA DE LECTURA"].apply(
                lambda x: dtparse(str(x), dayfirst=True) if pd.notna(x) and str(x).strip() != "" else pd.NaT
            ),
            errors="coerce"
        )
    else:
        df["FECHA_DE_LECTURA_DT"] = pd.NaT
    return df

def to_excel_bytes(df: pd.DataFrame, sheet_name="Reporte"):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet_name)
    out.seek(0)
    return out

# Detecta la(s) columna(s) con códigos en un archivo de dosis
CAND_COLS = ["CÓDIGO DE DOSÍMETRO","CODIGO DE DOSIMETRO","CODIGO","CÓDIGO","DOSIMETRO","DOSÍMETRO","WB","COD. DOSÍMETRO"]

def read_codes_from_file(file) -> Set[str]:
    name = file.name.lower()
    if name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(file)
    else:
        # CSV robusto (utf-8 / latin-1 / ; / ,)
        content = file.read()
        for enc in ("utf-8-sig", "latin-1"):
            try:
                df = pd.read_csv(BytesIO(content), encoding=enc, sep=None, engine="python")
                break
            except Exception:
                continue
        else:
            df = pd.read_csv(BytesIO(content))  # último intento estándar
    cols = [c for c in df.columns]
    # intenta por nombre
    target = None
    for c in cols:
        if any(k.lower() in str(c).lower() for k in CAND_COLS):
            target = c
            break
    # si no encontró, prueba patrones de tipo WB000123 en todo el DF
    if target is None:
        for c in cols:
            if df[c].astype(str).str.contains(r"^WB\d{5,}$", case=False, na=False).any():
                target = c
                break
    if target is None:
        # cae a primera columna
        target = cols[0]
    codes = (
        df[target]
        .astype(str)
        .str.strip()
        .str.replace("\u200b", "", regex=False)  # zero-width si viniera pegado
        .replace({"nan": ""})
    )
    return set([c for c in codes if c])

# ----------------- Carga Ninox -----------------
with st.spinner("Cargando datos desde Ninox…"):
    recs = fetch_all_records(TABLE_ID)
    base = normalize_df(recs)

if base.empty:
    st.warning("No hay registros en la tabla REPORTE.")
    st.stop()

# ----------------- Sidebar -----------------
with st.sidebar:
    st.header("Filtros")
    st.caption("Sube uno o varios **archivos de dosis** (CSV/Excel) para filtrar por códigos.")
    files = st.file_uploader("Archivos de dosis", type=["csv","xlsx","xls"], accept_multiple_files=True)

    # Periodo actual: por defecto el más reciente (excluye CONTROL)
    periodos_validos = base.loc[base["PERIODO DE LECTURA"].notna(), "PERIODO DE LECTURA"].astype(str)
    periodos_validos = [p for p in periodos_validos.unique().tolist() if p.strip().upper() != "CONTROL"]

    # Ordena periodos por la última lectura dentro de cada periodo
    per_orden = (
        base.groupby("PERIODO DE LECTURA")["FECHA_DE_LECTURA_DT"].max()
        .sort_values(ascending=False)
        .index.astype(str).tolist()
    )
    per_orden = [p for p in per_orden if p in periodos_validos]

    periodo_actual = st.selectbox("Periodo actual", per_orden, index=0 if per_orden else None)
    periodos_anteriores = st.multiselect(
        "Periodos anteriores (a sumar para 'Anual')",
        [p for p in per_orden if p != periodo_actual],
        default=[p for p in per_orden[1:2]]  # por defecto, el inmediatamente anterior
    )

    st.caption("Opcional: filtra también por **Compañía** o **Tipo de dosímetro**")
    companias = ["(todas)"] + sorted([c for c in base["COMPAÑÍA"].dropna().astype(str).unique()])
    compania_sel = st.selectbox("Compañía", companias, index=0)
    tipos = ["(todos)"] + sorted([c for c in base["TIPO DE DOSÍMETRO"].dropna().astype(str).unique()])
    tipo_sel = st.selectbox("Tipo de dosímetro", tipos, index=0)

# ----------------- Aplicar filtros por archivos -----------------
codes_filter: Optional[Set[str]] = None
if files:
    codes: Set[str] = set()
    for f in files:
        try:
            codes |= read_codes_from_file(f)
        except Exception as e:
            st.warning(f"No pude leer '{f.name}': {e}")
    if codes:
        codes_filter = set([c.strip() for c in codes if c.strip()])
        st.success(f"Códigos detectados para filtrar: {len(codes_filter)}")

df = base.copy()
if codes_filter:
    df = df[df["CÓDIGO DE DOSÍMETRO"].isin(codes_filter)]
if compania_sel != "(todas)":
    df = df[df["COMPAÑÍA"].astype(str) == compania_sel]
if tipo_sel != "(todos)":
    df = df[df["TIPO DE DOSÍMETRO"].astype(str) == tipo_sel]

if df.empty:
    st.warning("No hay registros que cumplan los filtros.")
    st.stop()

# ----------------- Cálculos -----------------
# Actual: para cada código, toma el registro más reciente DENTRO del periodo_actual
df_actual = (
    df[df["PERIODO DE LECTURA"].astype(str) == str(periodo_actual)]
    .sort_values(["CÓDIGO DE DOSÍMETRO","FECHA_DE_LECTURA_DT"], ascending=[True, False])
    .groupby("CÓDIGO DE DOSÍMETRO", as_index=False)
    .first()
    .rename(columns={
        "Hp (10)": "Hp10_ACTUAL",
        "Hp (0.07)": "Hp007_ACTUAL",
        "Hp (3)": "Hp3_ACTUAL",
        "FECHA DE LECTURA": "FECHA_LECTURA_ACTUAL"
    })
)

# Anteriores: suma por código en los periodos seleccionados
df_prev_sum = (
    df[df["PERIODO DE LECTURA"].astype(str).isin([str(p) for p in periodos_anteriores])]
    .groupby("CÓDIGO DE DOSÍMETRO", as_index=False)[["Hp (10)","Hp (0.07)","Hp (3)"]]
    .sum()
    .rename(columns={
        "Hp (10)": "Hp10_ANTERIOR_SUM",
        "Hp (0.07)": "Hp007_ANTERIOR_SUM",
        "Hp (3)": "Hp3_ANTERIOR_SUM"
    })
)

# Vida: suma histórica sobre todos los periodos (para el filtro aplicado)
df_vida = (
    df.groupby("CÓDIGO DE DOSÍMETRO", as_index=False)[["Hp (10)","Hp (0.07)","Hp (3)"]]
    .sum()
    .rename(columns={
        "Hp (10)": "Hp10_VIDA",
        "Hp (0.07)": "Hp007_VIDA",
        "Hp (3)": "Hp3_VIDA"
    })
)

# Unir
out = df_actual.merge(df_prev_sum, on="CÓDIGO DE DOSÍMETRO", how="left").merge(df_vida, on="CÓDIGO DE DOSÍMETRO", how="left")
for col in ["Hp10_ANTERIOR_SUM","Hp007_ANTERIOR_SUM","Hp3_ANTERIOR_SUM","Hp10_VIDA","Hp007_VIDA","Hp3_VIDA"]:
    if col not in out:
        out[col] = 0.0
    out[col] = out[col].fillna(0.0)

# Anual = Actual + suma anteriores
out["Hp10_ANUAL"]  = out["Hp10_ACTUAL"].fillna(0.0)  + out["Hp10_ANTERIOR_SUM"]
out["Hp007_ANUAL"] = out["Hp007_ACTUAL"].fillna(0.0) + out["Hp007_ANTERIOR_SUM"]
out["Hp3_ANUAL"]   = out["Hp3_ACTUAL"].fillna(0.0)   + out["Hp3_ANTERIOR_SUM"]

# Redondeo 2 decimales
num_cols = ["Hp10_ACTUAL","Hp007_ACTUAL","Hp3_ACTUAL","Hp10_ANUAL","Hp007_ANUAL","Hp3_ANUAL","Hp10_VIDA","Hp007_VIDA","Hp3_VIDA"]
for c in num_cols:
    if c in out.columns:
        out[c] = out[c].apply(round2)

# Columnas de contexto (nombre, cédula, compañía, tipo)
context_cols = ["CÓDIGO DE DOSÍMETRO","NOMBRE","CÉDULA","COMPAÑÍA","TIPO DE DOSÍMETRO","PERIODO DE LECTURA","FECHA_LECTURA_ACTUAL"]
# PERIODO DE LECTURA en df_actual es el actual; renombra para claridad
out = out.rename(columns={"PERIODO DE LECTURA": "PERIODO_ACTUAL"})

# Reordenar
final_cols = [
    "CÓDIGO DE DOSÍMETRO","NOMBRE","CÉDULA","COMPAÑÍA","TIPO DE DOSÍMETRO",
    "PERIODO_ACTUAL","FECHA_LECTURA_ACTUAL",
    "Hp10_ACTUAL","Hp007_ACTUAL","Hp3_ACTUAL",
    "Hp10_ANUAL","Hp007_ANUAL","Hp3_ANUAL",
    "Hp10_VIDA","Hp007_VIDA","Hp3_VIDA"
]
out = out[[c for c in final_cols if c in out.columns]].sort_values("CÓDIGO DE DOSÍMETRO")

st.subheader("Reporte final")
st.dataframe(out, use_container_width=True, hide_index=True)

# Descargas
csv_bytes = out.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "⬇️ Descargar CSV (UTF-8 con BOM)",
    data=csv_bytes,
    file_name=f"reporte_dosimetria_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
    mime="text/csv"
)
xlsx_bytes = to_excel_bytes(out)
st.download_button(
    "⬇️ Descargar Excel",
    data=xlsx_bytes,
    file_name=f"reporte_dosimetria_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

with st.expander("Notas"):
    st.markdown("""
- **PM** se trata como **0.00**.
- *Periodo actual* toma **el registro más reciente por dosímetro** dentro del periodo elegido.
- *Anual* = Actual + **suma** de todos los registros en los *Periodos anteriores seleccionados*.
- *De por vida* = suma de **todas** las lecturas históricas (según los filtros activos).
- Puedes subir **varios archivos**; se unifican todos los códigos detectados para filtrar.
""")


