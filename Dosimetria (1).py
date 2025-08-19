# app.py — Reporte de Dosimetría (Ninox) + filtro por archivos + Excel maquetado
# Requisitos:
#   pip install streamlit pandas requests openpyxl python-dateutil

import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime
from dateutil.parser import parse as dtparse
from typing import List, Dict, Any, Optional, Set
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ================== CREDENCIALES NINOX ==================
API_TOKEN   = "0b3a1130-785a-11f0-ace0-3fb1fcb242e2"  # <-- tu token
TEAM_ID     = "ihp8o8AaLzfodwc4J"
DATABASE_ID = "ksqzvuts5aq0"
BASE_URL    = "https://api.ninox.com/v1"
TABLE_ID    = "C"   # ID interno de la tabla REPORTE
# ========================================================

st.set_page_config(page_title="Reporte de Dosimetría — Ninox", layout="wide")
st.title("Reporte de Dosimetría — Actual, Anual y de por Vida")

# ---------------------- Utilidades ----------------------
def headers() -> Dict[str, str]:
    return {"Authorization": f"Bearer {API_TOKEN}", "Content-Type": "application/json"}

def as_value(v: Any):
    """Devuelve número si lo es; mantiene 'PM' para mostrar."""
    if v is None:
        return ""
    s = str(v).strip().replace(",", ".")
    if s.upper() == "PM":
        return "PM"
    try:
        return float(s)
    except Exception:
        return s

def as_num(v: Any) -> float:
    """Para cálculos: convierte a número; PM o vacío -> 0.0."""
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
            "FECHA DE NACIMIENTO": f.get("FECHA DE NACIMIENTO"),
            "FECHA DE LECTURA": f.get("FECHA DE LECTURA"),
            "TIPO DE DOSÍMETRO": f.get("TIPO DE DOSÍMETRO"),
            # RAW (mostrar, conserva PM)
            "Hp10_RAW": as_value(f.get("Hp (10)")),
            "Hp007_RAW": as_value(f.get("Hp (0.07)")),
            "Hp3_RAW": as_value(f.get("Hp (3)")),
            # NUM (cálculo)
            "Hp10_NUM": as_num(f.get("Hp (10)")),
            "Hp007_NUM": as_num(f.get("Hp (0.07)")),
            "Hp3_NUM": as_num(f.get("Hp (3)")),
        })
    df = pd.DataFrame(rows)
    # Parseo de fecha/hora
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

def read_codes_from_files(files) -> Set[str]:
    """Lee CSV/Excel y extrae códigos de dosímetro (columna candidata o patrón WB\d+)."""
    codes: Set[str] = set()
    from io import BytesIO
    for f in files:
        raw = f.read(); f.seek(0)
        name = f.name.lower()
        df = None
        try:
            if name.endswith((".xlsx", ".xls")):
                df = pd.read_excel(BytesIO(raw))
            else:
                for enc in ("utf-8-sig", "latin-1"):
                    try:
                        df = pd.read_csv(BytesIO(raw), sep=None, engine="python", encoding=enc)
                        break
                    except Exception:
                        continue
                if df is None:
                    df = pd.read_csv(BytesIO(raw))
        except Exception:
            continue
        if df is None or df.empty:
            continue

        cand = None
        for c in df.columns:
            cl = str(c).lower()
            if any(k in cl for k in ["dosim", "código", "codigo", "wb", "dosímetro", "dosimetro"]):
                cand = c; break
        if cand is None:
            for c in df.columns:
                if df[c].astype(str).str.contains(r"^WB\d{5,}$", case=False, na=False).any():
                    cand = c; break
        if cand is None:
            cand = df.columns[0]

        col = df[cand].astype(str).str.strip()
        codes |= set([c for c in col if c and c.lower() != "nan"])
    return codes

# --------- Excel maquetado (estilo plantilla) ----------
def build_formatted_excel(df: pd.DataFrame) -> bytes:
    """
    Genera un Excel con encabezados:
    [PERIODO..., COMPAÑÍA, CÓDIGO..., NOMBRE, CÉDULA, FECHA NAC., FECHA Y HORA DE LECTURA, TIPO]
    + bloques: DOSIS ACTUAL / DOSIS ANUAL / DOSIS DE POR VIDA (Hp10, Hp0.07, Hp3)
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte"

    # Estilos
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Título
    ws.merge_cells("A1:Q1")
    ws["A1"] = "REPORTE DE DOSIMETRÍA"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = center

    # Fila 2: cabeceras de bloques
    ws.merge_cells("A2:H3"); ws["A2"] = ""    # área datos fijos
    ws.merge_cells("I2:K2"); ws["I2"] = "DOSIS ACTUAL (mSv)"
    ws.merge_cells("L2:N2"); ws["L2"] = "DOSIS ANUAL (mSv)"
    ws.merge_cells("O2:Q2"); ws["O2"] = "DOSIS DE POR VIDA (mSv)"
    for c in ["I2","L2","O2"]:
        ws[c].font = bold; ws[c].alignment = center

    # Fila 3: subcolumnas Hp
    ws["I3"], ws["J3"], ws["K3"] = "Hp(10)", "Hp(0.07)", "Hp(3)"
    ws["L3"], ws["M3"], ws["N3"] = "Hp(10)", "Hp(0.07)", "Hp(3)"
    ws["O3"], ws["P3"], ws["Q3"] = "Hp(10)", "Hp(0.07)", "Hp(3)"
    for c in ["I3","J3","K3","L3","M3","N3","O3","P3","Q3"]:
        ws[c].font = bold; ws[c].alignment = center; ws[c].border = border

    # Fila 4: cabeceras de datos fijos
    headers = [
        "PERIODO DE LECTURA","COMPAÑÍA","CÓDIGO DE DOSÍMETRO","NOMBRE","CÉDULA",
        "FECHA DE NACIMIENTO","FECHA Y HORA DE LECTURA","TIPO DE DOSÍMETRO",
        "Hp (10) ACTUAL","Hp (0.07) ACTUAL","Hp (3) ACTUAL",
        "Hp (10) ANUAL","Hp (0.07) ANUAL","Hp (3) ANUAL",
        "Hp (10) VIDA","Hp (0.07) VIDA","Hp (3) VIDA",
    ]
    ws.append(headers)
    for col in range(1, 18):
        cell = ws.cell(row=4, column=col)
        cell.font = bold; cell.alignment = center; cell.border = border

    # Ancho de columnas
    widths = [16,16,18,24,16,20,22,16,12,12,12,12,14,12,12,14,12]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    # Datos
    start_row = 5
    for _, r in df.iterrows():
        ws.append(list(r.values))
    # Bordes para todo el rango de datos
    last_row = ws.max_row
    for row in ws.iter_rows(min_row=4, max_row=last_row, min_col=1, max_col=17):
        for cell in row:
            cell.border = border
            if cell.row >= 5:
                cell.alignment = Alignment(vertical="center", wrap_text=True)

    # Salida
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

# ------------------- Carga Ninox -------------------
with st.spinner("Cargando datos desde Ninox…"):
    recs = fetch_all_records(TABLE_ID)
    base = normalize_df(recs)

if base.empty:
    st.warning("No hay registros en la tabla REPORTE.")
    st.stop()

# ------------------- Sidebar -------------------
with st.sidebar:
    st.header("Filtros")
    files = st.file_uploader("Archivos de dosis (para filtrar)", type=["csv", "xlsx", "xls"], accept_multiple_files=True)

    # periodos válidos (excluye CONTROL para selección)
    per_order = (base.groupby("PERIODO DE LECTURA")["FECHA_DE_LECTURA_DT"].max()
                 .sort_values(ascending=False).index.astype(str).tolist())
    per_valid = [p for p in per_order if p.strip().upper() != "CONTROL"]

    periodo_actual = st.selectbox("Periodo actual", per_valid, index=0 if per_valid else None)
    periodos_anteriores = st.multiselect(
        "Periodos anteriores (para ANUAL)",
        [p for p in per_valid if p != periodo_actual],
        default=[per_valid[1]] if len(per_valid) > 1 else []
    )

    comp_opts = ["(todas)"] + sorted(base["COMPAÑÍA"].dropna().astype(str).unique().tolist())
    compania = st.selectbox("Compañía", comp_opts, index=0)
    tipo_opts = ["(todos)"] + sorted(base["TIPO DE DOSÍMETRO"].dropna().astype(str).unique().tolist())
    tipo = st.selectbox("Tipo de dosímetro", tipo_opts, index=0)

# Filtro por archivos
codes_filter: Optional[Set[str]] = None
if files:
    codes_filter = read_codes_from_files(files)
    if codes_filter:
        st.success(f"Códigos detectados en archivos: {len(codes_filter)}")

df = base.copy()
if codes_filter:
    df = df[df["CÓDIGO DE DOSÍMETRO"].isin(codes_filter)]
if compania != "(todas)":
    df = df[df["COMPAÑÍA"].astype(str) == compania]
if tipo != "(todos)":
    df = df[df["TIPO DE DOSÍMETRO"].astype(str) == tipo]

if df.empty:
    st.warning("No hay registros que cumplan el filtro.")
    st.stop()

# ---- Identificar CONTROL (por nombre) y opción manual
control_codes = set(df.loc[df["NOMBRE"].astype(str).str.strip().str.upper() == "CONTROL",
                           "CÓDIGO DE DOSÍMETRO"].unique())
all_codes = sorted(df["CÓDIGO DE DOSÍMETRO"].unique().tolist())
manual_control = st.sidebar.selectbox("Código CONTROL (manual, opcional)", ["(auto)"] + all_codes, index=0)
if manual_control != "(auto)":
    control_codes.add(manual_control)

# ------------------- Cálculos -------------------
def ultimo_en_periodo(g: pd.DataFrame, periodo: str) -> pd.Series:
    x = g[g["PERIODO DE LECTURA"].astype(str) == str(periodo)].sort_values("FECHA_DE_LECTURA_DT", ascending=False)
    return x.iloc[0] if not x.empty else pd.Series(dtype="object")

# Actual: último por código dentro del periodo actual
rows = []
for code, sub in df.groupby("CÓDIGO DE DOSÍMETRO", as_index=False):
    ult = ultimo_en_periodo(sub, periodo_actual)
    if ult.empty:
        continue
    rows.append({
        "CÓDIGO DE DOSÍMETRO": code,
        "PERIODO DE LECTURA": periodo_actual,
        "COMPAÑÍA": ult.get("COMPAÑÍA"),
        "NOMBRE": ult.get("NOMBRE"),
        "CÉDULA": ult.get("CÉDULA"),
        "FECHA DE NACIMIENTO": ult.get("FECHA DE NACIMIENTO"),
        "FECHA Y HORA DE LECTURA": ult.get("FECHA DE LECTURA"),
        "TIPO DE DOSÍMETRO": ult.get("TIPO DE DOSÍMETRO"),
        # Mostrar (conserva PM)
        "Hp10_ACTUAL":  ult.get("Hp10_RAW"),
        "Hp007_ACTUAL": ult.get("Hp007_RAW"),
        "Hp3_ACTUAL":   ult.get("Hp3_RAW"),
        # Num para cálculos
        "Hp10_ACTUAL_NUM":  ult.get("Hp10_NUM", 0.0),
        "Hp007_ACTUAL_NUM": ult.get("Hp007_NUM", 0.0),
        "Hp3_ACTUAL_NUM":   ult.get("Hp3_NUM", 0.0),
    })
df_actual = pd.DataFrame(rows)

# Suma de periodos anteriores (para ANUAL)
df_prev = df[df["PERIODO DE LECTURA"].astype(str).isin(periodos_anteriores)]
prev_sum = (df_prev.groupby("CÓDIGO DE DOSÍMETRO")[["Hp10_NUM", "Hp007_NUM", "Hp3_NUM"]]
            .sum()
            .rename(columns={"Hp10_NUM": "Hp10_ANT_SUM",
                             "Hp007_NUM": "Hp007_ANT_SUM",
                             "Hp3_NUM": "Hp3_ANT_SUM"}))

# Suma de por vida (histórico, con filtros aplicados)
vida_sum = (df.groupby("CÓDIGO DE DOSÍMETRO")[["Hp10_NUM", "Hp007_NUM", "Hp3_NUM"]]
            .sum()
            .rename(columns={"Hp10_NUM": "Hp10_VIDA",
                             "Hp007_NUM": "Hp007_VIDA",
                             "Hp3_NUM": "Hp3_VIDA"}))

# Unir
out = (df_actual.set_index("CÓDIGO DE DOSÍMETRO")
       .join(prev_sum, how="left")
       .join(vida_sum, how="left")).reset_index()

for c in ["Hp10_ANT_SUM", "Hp007_ANT_SUM", "Hp3_ANT_SUM", "Hp10_VIDA", "Hp007_VIDA", "Hp3_VIDA"]:
    if c not in out:
        out[c] = 0.0
    out[c] = out[c].fillna(0.0)

# ANUAL = ACTUAL + ANTERIORES (numéricos)
out["Hp10_ANUAL"]  = out["Hp10_ACTUAL_NUM"]  + out["Hp10_ANT_SUM"]
out["Hp007_ANUAL"] = out["Hp007_ACTUAL_NUM"] + out["Hp007_ANT_SUM"]
out["Hp3_ANUAL"]   = out["Hp3_ACTUAL_NUM"]   + out["Hp3_ANT_SUM"]

# Redondeos de numéricos (ANUAL y VIDA)
for c in ["Hp10_ANUAL", "Hp007_ANUAL", "Hp3_ANUAL", "Hp10_VIDA", "Hp007_VIDA", "Hp3_VIDA"]:
    out[c] = out[c].apply(round2)

# Mantener PM en ACTUAL (si era PM); si no, redondear número
def show_raw_or_num(raw):
    return raw if str(raw).upper() == "PM" else round2(float(raw))

out["Hp10_ACTUAL"]  = out["Hp10_ACTUAL"].apply(show_raw_or_num)
out["Hp007_ACTUAL"] = out["Hp007_ACTUAL"].apply(show_raw_or_num)
out["Hp3_ACTUAL"]   = out["Hp3_ACTUAL"].apply(show_raw_or_num)

# CONTROL primero
out["__is_control"] = out["CÓDIGO DE DOSÍMETRO"].isin(control_codes)
out = out.sort_values(["__is_control", "CÓDIGO DE DOSÍMETRO"], ascending=[False, True])

# Renombrar a encabezados finales y ordenar columnas
rename_map = {
    "Hp10_ACTUAL":  "Hp (10) ACTUAL",
    "Hp007_ACTUAL": "Hp (0.07) ACTUAL",
    "Hp3_ACTUAL":   "Hp (3) ACTUAL",
    "Hp10_ANUAL":   "Hp (10) ANUAL",
    "Hp007_ANUAL":  "Hp (0.07) ANUAL",
    "Hp3_ANUAL":    "Hp (3) ANUAL",
    "Hp10_VIDA":    "Hp (10) VIDA",
    "Hp007_VIDA":   "Hp (0.07) VIDA",
    "Hp3_VIDA":     "Hp (3) VIDA",
}
out = out.rename(columns=rename_map)

final_cols = [
    "PERIODO DE LECTURA", "COMPAÑÍA", "CÓDIGO DE DOSÍMETRO", "NOMBRE", "CÉDULA",
    "FECHA DE NACIMIENTO", "FECHA Y HORA DE LECTURA", "TIPO DE DOSÍMETRO",
    "Hp (10) ACTUAL", "Hp (0.07) ACTUAL", "Hp (3) ACTUAL",
    "Hp (10) ANUAL", "Hp (0.07) ANUAL", "Hp (3) ANUAL",
    "Hp (10) VIDA", "Hp (0.07) VIDA", "Hp (3) VIDA",
]
for c in final_cols:
    if c not in out.columns:
        out[c] = ""
out = out[final_cols]

# ------------------- Mostrar / Descargar -------------------
st.subheader("Reporte final")
st.dataframe(out, use_container_width=True, hide_index=True)

# CSV
csv_bytes = out.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "⬇️ Descargar CSV (UTF-8 con BOM)",
    data=csv_bytes,
    file_name=f"reporte_dosimetria_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
    mime="text/csv"
)

# Excel simple
def to_excel_simple(df: pd.DataFrame, sheet_name="Reporte"):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet_name)
    bio.seek(0)
    return bio.getvalue()

xlsx_simple = to_excel_simple(out)
st.download_button(
    "⬇️ Descargar Excel (tabla simple)",
    data=xlsx_simple,
    file_name=f"reporte_dosimetria_tabla_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Excel maquetado tipo plantilla
xlsx_fmt = build_formatted_excel(out.copy())
st.download_button(
    "⬇️ Descargar Excel (formato plantilla)",
    data=xlsx_fmt,
    file_name=f"reporte_dosimetria_plantilla_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

with st.expander("Notas"):
    st.markdown("""
- **PM** se mantiene en **ACTUAL**; en **ANUAL** y **VIDA** se considera como **0.00** para las sumas.
- **ANUAL** = último valor del periodo seleccionado (**ACTUAL**) + **suma** de los periodos anteriores seleccionados.
- **VIDA** = suma histórica de todas las lecturas (con los filtros activos).
- La fila **CONTROL** (si existe) se muestra primero.
""")



