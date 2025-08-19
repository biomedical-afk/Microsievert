# app.py — Reporte de Dosimetría (Ninox) + filtros + Excel tipo plantilla
import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime
from dateutil.parser import parse as dtparse
from typing import List, Dict, Any, Optional, Set

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

# ================== CREDENCIALES NINOX ==================
API_TOKEN   = "0b3a1130-785a-11f0-ace0-3fb1fcb242e2"  # <-- tu API key
TEAM_ID     = "ihp8o8AaLzfodwc4J"
DATABASE_ID = "ksqzvuts5aq0"
BASE_URL    = "https://api.ninox.com/v1"
TABLE_ID    = "C"   # Tabla REPORTE
# ========================================================

st.set_page_config(page_title="Reporte de Dosimetría — Ninox", layout="wide")
st.title("Reporte de Dosimetría — Actual, Anual y de por Vida")

# ---------------------- Utilidades ----------------------
def headers() -> Dict[str, str]:
    return {"Authorization": f"Bearer {API_TOKEN}", "Content-Type": "application/json"}

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

def fetch_all_records(table_id: str, page_size: int = 1000):
    url = f"{BASE_URL}/teams/{TEAM_ID}/databases/{DATABASE_ID}/tables/{table_id}/records"
    skip, out = 0, []
    while True:
        r = requests.get(url, headers=headers(), params={"limit": page_size, "skip": skip}, timeout=60)
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

def read_codes_from_files(files) -> Set[str]:
    codes: Set[str] = set()
    from io import BytesIO
    for f in files:
        raw = f.read(); f.seek(0)
        name = f.name.lower()
        try:
            if name.endswith((".xlsx", ".xls")):
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

def pm_or_sum(raws: List[Any], numeric_sum: float) -> Any:
    vals = [str(x).upper() for x in raws if str(x).strip()!=""]
    if vals and all(v == "PM" for v in vals): return "PM"
    return round2(numeric_sum)

# ------------------- Sidebar -------------------
with st.sidebar:
    st.header("Filtros")
    files = st.file_uploader("Archivos de dosis (para filtrar)", type=["csv","xlsx","xls"], accept_multiple_files=True)

    st.markdown("---")
    st.subheader("Encabezado del reporte")
    header_line1 = st.text_input("Línea 1", "MICROSIEVERT, S.A.")
    header_line2 = st.text_input("Línea 2", "PH Conardo")
    header_line3 = st.text_input("Línea 3", "Calle 41 Este, Panamá")
    header_line4 = st.text_input("Línea 4", "PANAMÁ")
    logo_file = st.file_uploader("Logo (PNG/JPG) opcional", type=["png","jpg","jpeg"])

# ------------------- Carga Ninox -------------------
with st.spinner("Cargando datos desde Ninox…"):
    base = normalize_df(fetch_all_records(TABLE_ID))

if base.empty:
    st.warning("No hay registros en la tabla REPORTE.")
    st.stop()

# ------------------- Filtros adicionales -------------------
with st.sidebar:
    st.markdown("---")
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

# --------- Filtrado de registros ---------
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

# ---- Identificar CONTROL (por nombre) ----
control_codes = set(df.loc[df["NOMBRE"].astype(str).str.strip().str.upper()=="CONTROL",
                           "CÓDIGO DE DOSÍMETRO"].unique())

# ------------------- Cálculos -------------------
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
       .join(prev_sum, how="left")
       .join(vida_sum, how="left")
       .join(vida_raw, how="left")).reset_index()

for c in ["Hp10_ANT_SUM","Hp007_ANT_SUM","Hp3_ANT_SUM","Hp10_VIDA_NUM","Hp007_VIDA_NUM","Hp3_VIDA_NUM"]:
    if c not in out: out[c] = 0.0
    out[c] = out[c].fillna(0.0)

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

# ------------------- Vista previa -------------------
st.subheader("Reporte final (vista previa)")
st.dataframe(out, use_container_width=True, hide_index=True)

# ------------------- Descargas -------------------
csv_bytes = out.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "⬇️ Descargar CSV (UTF-8 con BOM)",
    data=csv_bytes,
    file_name=f"reporte_dosimetria_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
    mime="text/csv"
)

def to_excel_simple(df: pd.DataFrame, sheet_name="Reporte"):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet_name)
    bio.seek(0); return bio.getvalue()

xlsx_simple = to_excel_simple(out)
st.download_button(
    "⬇️ Descargar Excel (tabla simple)",
    data=xlsx_simple,
    file_name=f"reporte_dosimetria_tabla_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ------------------- Excel formato plantilla -------------------
def build_formatted_excel(df_final: pd.DataFrame, header_lines: List[str], logo_bytes: Optional[bytes]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte"

    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Encabezado (A1:B4) + Logo (C1) + Fecha de emisión derecha
    gray = PatternFill("solid", fgColor="DDDDDD")
    for i, line in enumerate(header_lines[:4], start=1):
        ws.merge_cells(f"A{i}:B{i}")
        c = ws[f"A{i}"]; c.value = line; c.fill = gray
        c.font = Font(bold=True); c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[i].height = 18
        for col in ("A","B"):
            ws.cell(row=i, column=ord(col)-64).border = border

    # Fecha de emisión (colocada en el extremo derecho de la fila 1)
    label_cell = "I1"; value_cell = "K1"
    ws.merge_cells(f"{label_cell}:J1")
    ws[label_cell] = "Fecha de emisión"
    ws[label_cell].font = Font(bold=True, size=10)
    ws[label_cell].alignment = center
    ws[label_cell].fill = gray
    ws.merge_cells(f"{value_cell}:P1")
    fecha_emision = datetime.now().strftime("%d-%b-%y").lower()
    ws[value_cell] = fecha_emision
    ws[value_cell].alignment = center
    ws[value_cell].font = Font(bold=True, size=10)
    for col_idx in range(ord("I")-64, ord("P")-64+1):
        ws.cell(row=1, column=col_idx).border = border

    # Logo (tamaño controlado)
    if logo_bytes:
        try:
            img = XLImage(BytesIO(logo_bytes))
            img.width = 300   # ← ajusta si lo quieres más grande/pequeño
            img.height = 90
            img.anchor = "C1"  # lo colocamos desde C1
            ws.add_image(img)
        except Exception:
            pass

    # Título
    TITLE_ROW = 6
    ws.merge_cells(f"A{TITLE_ROW}:P{TITLE_ROW}")
    ws[f"A{TITLE_ROW}"] = "REPORTE DE DOSIMETRÍA"
    ws[f"A{TITLE_ROW}"].font = Font(bold=True, size=14)
    ws[f"A{TITLE_ROW}"].alignment = center

    # Bloques
    BLOCK_ROW = 7
    ws.merge_cells(f"H{BLOCK_ROW}:J{BLOCK_ROW}"); ws[f"H{BLOCK_ROW}"] = "DOSIS ACTUAL (mSv)"
    ws.merge_cells(f"K{BLOCK_ROW}:M{BLOCK_ROW}"); ws[f"K{BLOCK_ROW}"] = "DOSIS ANUAL (mSv)"
    ws.merge_cells(f"N{BLOCK_ROW}:P{BLOCK_ROW}"); ws[f"N{BLOCK_ROW}"] = "DOSIS DE POR VIDA (mSv)"
    for c in (f"H{BLOCK_ROW}", f"K{BLOCK_ROW}", f"N{BLOCK_ROW}"):
        ws[c].font = bold; ws[c].alignment = center

    # Encabezados
    HEADER_ROW = 8
    headers = [
        "PERIODO DE LECTURA","COMPAÑÍA","CÓDIGO DE DOSÍMETRO","NOMBRE","CÉDULA",
        "FECHA Y HORA DE LECTURA","TIPO DE DOSÍMETRO",
        "Hp (10) ACTUAL","Hp (0.07) ACTUAL","Hp (3) ACTUAL",
        "Hp (10) ANUAL","Hp (0.07) ANUAL","Hp (3) ANUAL",
        "Hp (10) VIDA","Hp (0.07) VIDA","Hp (3) VIDA",
    ]
    for idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=HEADER_ROW, column=idx, value=h)
        cell.font = bold; cell.alignment = center; cell.border = border

    # Datos
    START_ROW = HEADER_ROW + 1
    for _, r in df_final[headers].iterrows():
        ws.append(list(r.values))
    last_row = ws.max_row

    # Bordes y wrap
    for row in ws.iter_rows(min_row=HEADER_ROW, max_row=last_row, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = border
            if cell.row >= START_ROW:
                cell.alignment = Alignment(vertical="center", wrap_text=True)

    ws.freeze_panes = f"A{START_ROW}"

    # Auto-ancho
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for c in col_cells:
            txt = "" if c.value is None else str(c.value)
            if len(txt) > max_len: max_len = len(txt)
        ws.column_dimensions[col_letter].width = max(12, min(max_len + 2, 42))

    # ====== Sección informativa ======
    row = last_row + 2
    ws.merge_cells(f"A{row}:P{row}")
    c = ws[f"A{row}"]; c.value = "INFORMACIÓN DEL REPORTE DE DOSIMETRÍA"
    c.font = Font(bold=True); c.alignment = Alignment(horizontal="center")
    row += 1

    bullets = [
        "‒ Periodo de lectura: periodo de uso del dosímetro personal.",
        "‒ Fecha de lectura: fecha en que se realizó la lectura.",
        "‒ Tipo de dosímetro:",
    ]
    for text in bullets:
        ws.merge_cells(f"A{row}:D{row}")
        c = ws[f"A{row}"]; c.value = text
        c.font = Font(size=10, bold=True); c.alignment = Alignment(horizontal="left")
        row += 2

    tipos = [("CE","Cuerpo Entero"), ("A","Anillo"), ("B","Brazalete"), ("CR","Cristalino")]
    for clave, desc in tipos:
        ws.merge_cells(f"C{row}:D{row}")
        c = ws[f"C{row}"]; c.value = f"{clave} = {desc}"
        c.font = Font(size=10, bold=True); c.alignment = Alignment(horizontal="left")
        for col in ("C","D"):
            ws.cell(row=row, column=ord(col)-64).border = border
        row += 1
    row += 1

    ws.merge_cells(f"F{row}:I{row}")
    c = ws[f"F{row}"]; c.value = "LÍMITES ANUALES DE EXPOSICIÓN A RADIACIONES"
    c.font = Font(bold=True, size=10); c.alignment = Alignment(horizontal="center")
    row += 1

    limites = [
        ("Cuerpo Entero", "20 mSv/año"),
        ("Cristalino", "150 mSv/año"),
        ("Extremidades y piel", "500 mSv/año"),
        ("Fetal", "1 mSv/periodo de gestación"),
        ("Público", "1 mSv/año"),
    ]
    for cat, val in limites:
        ws.merge_cells(f"F{row}:G{row}"); ws[f"F{row}"].value = cat
        ws[f"F{row}"].font = Font(size=10); ws[f"F{row}"].alignment = Alignment(horizontal="left")
        ws.merge_cells(f"H{row}:I{row}"); ws[f"H{row}"].value = val
        ws[f"H{row}"].font = Font(size=10); ws[f"H{row}"].alignment = Alignment(horizontal="right")
        for col in ("F","G","H","I"):
            ws.cell(row=row, column=ord(col)-64).border = border
        row += 1
    row += 2

    ws.merge_cells(f"A{row}:P{row}")
    c = ws[f"A{row}"]; c.value = "‒ DATOS DEL PARTICIPANTE:"
    c.font = Font(bold=True, size=10); c.alignment = Alignment(horizontal="left")
    row += 1

    datos = [
        "‒ Código de usuario: Número único asignado al usuario por Microsievert, S.A.",
        "‒ Nombre: Persona a la cual se le asigna el dosímetro personal.",
        "‒ Cédula: Número del documento de identidad personal del usuario.",
    ]
    for txt in datos:
        ws.merge_cells(f"A{row}:P{row}")
        c = ws[f"A{row}"]; c.value = txt
        c.font = Font(size=10); c.alignment = Alignment(horizontal="left")
        row += 1
    row += 2

    ws.merge_cells(f"A{row}:P{row}")
    c = ws[f"A{row}"]; c.value = "‒ DOSIS EN MILISIEVERT:"
    c.font = Font(bold=True, size=10); c.alignment = Alignment(horizontal="left")
    row += 1

    shade = PatternFill("solid", fgColor="DDDDDD")
    ws.merge_cells(f"B{row}:C{row}")
    hb = ws[f"B{row}"]; hb.value = "Nombre"; hb.font = Font(bold=True, size=10)
    hb.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); hb.fill = shade
    ws.merge_cells(f"D{row}:I{row}")
    hd = ws[f"D{row}"]; hd.value = "Definición"; hd.font = Font(bold=True, size=10)
    hd.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); hd.fill = shade
    ws.merge_cells(f"J{row}:J{row}")
    hu = ws[f"J{row}"]; hu.value = "Unidad"; hu.font = Font(bold=True, size=10)
    hu.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); hu.fill = shade
    for col in ("B","C","D","E","F","G","H","I","J"):
        ws.cell(row=row, column=ord(col)-64).border = border
    ws.row_dimensions[row].height = 30
    row += 1

    definitions = [
        ("Dosis efectiva Hp(10)",  "Es la dosis equivalente en tejido blando, J·kg⁻¹ o Sv a una profundidad de 10 mm, bajo determinado punto del cuerpo.", "mSv"),
        ("Dosis superficial Hp(0,07)", "Es la dosis equivalente en tejido blando, J·kg⁻¹ o Sv a una profundidad de 0,07 mm, bajo determinado punto del cuerpo.", "mSv"),
        ("Dosis cristalino Hp(3)", "Es la dosis equivalente en tejido blando, J·kg⁻¹ o Sv a una profundidad de 3 mm, bajo determinado punto del cuerpo.", "mSv"),
    ]
    for nom, desc, uni in definitions:
        ws.merge_cells(f"B{row}:C{row}")
        c = ws[f"B{row}"]; c.value = nom
        c.font = Font(size=10, bold=True); c.alignment = Alignment(horizontal="left", wrap_text=True)
        ws.merge_cells(f"D{row}:I{row}")
        c = ws[f"D{row}"]; c.value = desc
        c.font = Font(size=10); c.alignment = Alignment(horizontal="left", wrap_text=True)
        ws.merge_cells(f"J{row}:J{row}")
        c = ws[f"J{row}"]; c.value = uni
        c.font = Font(size=10); c.alignment = Alignment(horizontal="center", wrap_text=True)
        for col in ("B","C","D","E","F","G","H","I","J"):
            cc = ws.cell(row=row, column=ord(col)-64)
            cc.border = border; cc.alignment = Alignment(wrap_text=True)
        ws.row_dimensions[row].height = 30
        row += 1

    row += 1
    ws.merge_cells(f"A{row}:P{row}")
    c = ws[f"A{row}"]
    c.value = "LECTURAS DE ANILLO: las lecturas del dosímetro de anillo son registradas como una dosis equivalente superficial Hp(0,07)."
    c.font = Font(size=10, bold=True); c.alignment = Alignment(horizontal="left", wrap_text=True)
    row += 1

    ws.merge_cells(f"A{row}:P{row}")
    c = ws[f"A{row}"]
    c.value = "Los resultados de las dosis individuales de radiación son reportados para diferentes periodos de tiempo:"
    c.font = Font(size=10); c.alignment = Alignment(horizontal="left", wrap_text=True)
    row += 1

    periods = [
        ("DOSIS ACTUAL",      "Es el correspondiente de dosis acumulada durante el período de lectura definido."),
        ("DOSIS ANUAL",       "Es el correspondiente de dosis acumulada desde el inicio del año hasta la fecha."),
        ("DOSIS DE POR VIDA", "Es el correspondiente de dosis acumulada desde el inicio del servicio dosimétrico hasta la fecha."),
    ]
    for clave, texto in periods:
        ws.merge_cells(f"B{row}:C{row}")
        c = ws[f"B{row}"]; c.value = clave
        c.font = Font(bold=True, size=10); c.alignment = Alignment(horizontal="center")
        ws.merge_cells(f"D{row}:P{row}")
        c = ws[f"D{row}"]; c.value = texto
        c.font = Font(size=10); c.alignment = Alignment(horizontal="left", wrap_text=True)
        for col_idx in range(ord("B")-64, ord("P")-64+1):
            ws.cell(row=row, column=col_idx).border = border
        row += 1

    row += 2
    ws.merge_cells(f"A{row}:P{row}")
    c = ws[f"A{row}"]
    c.value = ("DOSÍMETRO DE CONTROL: incluido en cada paquete entregado para monitorear la exposición a la radiación "
               "recibida durante el tránsito y almacenamiento. Este dosímetro debe ser guardado por el cliente en un "
               "área libre de radiación durante el período de uso.")
    c.font = Font(size=10, bold=True); c.alignment = Alignment(horizontal="left", wrap_text=True)
    row += 2

    ws.merge_cells(f"A{row}:P{row}")
    c = ws[f"A{row}"]
    c.value = ("POR DEBAJO DEL MÍNIMO DETECTADO: es la dosis por debajo de la cantidad mínima reportada para el período "
               "de uso y son registradas como \"PM\".")
    c.font = Font(size=10, bold=True); c.alignment = Alignment(horizontal="left", wrap_text=True)

    bio = BytesIO(); wb.save(bio); bio.seek(0)
    return bio.getvalue()

header_lines = [
    st.session_state.get("header_line1", header_line1),
    st.session_state.get("header_line2", header_line2),
    st.session_state.get("header_line3", header_line3),
    st.session_state.get("header_line4", header_line4),
]
logo_bytes = logo_file.read() if logo_file is not None else None

xlsx_fmt = build_formatted_excel(out.copy(), header_lines, logo_bytes)
st.download_button(
    "⬇️ Descargar Excel (formato plantilla)",
    data=xlsx_fmt,
    file_name=f"reporte_dosimetria_plantilla_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

with st.expander("Notas"):
    st.markdown("""
- Logo escalado a **300×90 px** (ajústalo en `img.width` / `img.height` si lo deseas).
- “**Fecha de emisión**” se coloca automáticamente en la fila 1 (derecha) con el momento de descarga.
- **PM** se respeta en ACTUAL y también en ANUAL/VIDA cuando todas las lecturas que aportan son PM.
- **CONTROL** se ordena primero.
""")






