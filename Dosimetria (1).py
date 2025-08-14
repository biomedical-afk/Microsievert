# app.py — Sistema de Gestión de Dosimetría (versión corregida)

import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from datetime import datetime
import re

# =========================
# CONFIGURACIÓN DE PÁGINA
# =========================
st.set_page_config(page_title="Sistema de Gestión de Dosimetría", layout="centered")

# --- COLORES DE LA APP ---
COLOR_PRIMARIO   = "#1c6758"
COLOR_SECUNDARIO = "#3d8361"
COLOR_FONDO      = "#f8f9fa"
COLOR_TEXTO      = "#222"
COLOR_FOOTER     = "#6c757d"

# --- LOGIN SIMPLE ---
USUARIOS = {
    "Mispanama": "Maxilo2000",  # Cambia estos valores a lo que prefieras
    "usuario1": "password123"
}

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.markdown(
        "<h2 style='text-align:center; color:#1c6758'>Acceso al Sistema de Gestión de Dosimetría</h2>",
        unsafe_allow_html=True
    )
    usuario = st.text_input("Usuario", key="login_user")
    password = st.text_input("Contraseña", type="password", key="login_pass")
    if st.button("Ingresar"):
        if usuario in USUARIOS and password == USUARIOS[usuario]:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos.")
    st.stop()  # Detiene aquí si no ha iniciado sesión

# --- OPCIÓN CERRAR SESIÓN ---
if st.sidebar.button("Cerrar sesión"):
    st.session_state["autenticado"] = False
    st.rerun()

# =========================
# INTERFAZ PRINCIPAL
# =========================
st.markdown(
    f"""
    <h1 style="color:{COLOR_PRIMARIO};text-align:center">Sistema de Gestión de Dosimetría</h1>
    <hr style="border-top: 2px solid {COLOR_SECUNDARIO};">
    """,
    unsafe_allow_html=True
)
st.write("Sube los archivos necesarios, selecciona el período, pon el nombre del reporte y genera el Excel profesional.")

col1, col2 = st.columns(2)
with col1:
    participantes_file = st.file_uploader("Archivo de Participantes (.xlsx, .xls)", type=["xlsx", "xls"], key="participantes")
with col2:
    dosis_file = st.file_uploader("Archivo de Dosis (.xlsx, .xls, .csv)", type=["xlsx", "xls", "csv"], key="dosis")

# =========================
# UTILIDADES
# =========================
def _try_read_csv(file):
    """Intenta leer CSV con ; y ,"""
    try:
        return pd.read_csv(file, sep=';', engine='python')
    except Exception:
        file.seek(0)
        return pd.read_csv(file)

def normalizar_cols_participantes(df):
    # Pasar a mayúsculas, quitar espacios del inicio/fin
    df.columns = [c.strip().upper() for c in df.columns]
    return df

def leer_participantes(f):
    df = pd.read_excel(f)
    df = normalizar_cols_participantes(df)

    # Campos que solemos usar; si no están, se crean vacíos para evitar KeyError
    for col in ["DOSIMETRO 1", "DOSIMETRO 2", "NOMBRE", "APELLIDO", "CÉDULA", "COMPAÑÍA", "PERIODO 1", "PERIODO 2"]:
        if col not in df.columns:
            df[col] = ""

    # Limpieza básica
    for c in ["DOSIMETRO 1", "DOSIMETRO 2"]:
        df[c] = df[c].astype(str).str.strip().str.upper().replace({"NAN": ""})

    return df

def leer_dosis(f):
    # Lee Excel o CSV y normaliza columnas
    if f.name.lower().endswith(".csv"):
        df = _try_read_csv(f)
    else:
        df = pd.read_excel(f)

    # Estandarizar columnas a minúsculas sin espacios y sin paréntesis/puntos molestos
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(' ', '', regex=False)
        .str.replace('(', '', regex=False)
        .str.replace(')', '', regex=False)
    )

    # Mapear varias variantes a un nombre canónico
    rename_map = {
        # dosímetro y timestamp
        "dosimeter": "dosimeter",
        "serial": "dosimeter",
        "codigo": "dosimeter",
        "timestamp": "timestamp",
        "fecha": "timestamp",
        "fechalectura": "timestamp",

        # dosis (soportar *_corr. y sin _corr, y hp0.07)
        "hp10dosecorr.": "hp10",
        "hp10dosecorr": "hp10",
        "hp10dose": "hp10",
        "hp007dosecorr.": "hp007",
        "hp0.07dosecorr.": "hp007",
        "hp0.07dosecorr": "hp007",
        "hp007dose": "hp007",
        "hp0.07dose": "hp007",
        "hp3dosecorr.": "hp3",
        "hp3dosecorr": "hp3",
        "hp3dose": "hp3",
    }

    # Aplicar renombrado solo a las columnas presentes
    present_map = {k: v for k, v in rename_map.items() if k in df.columns}
    df = df.rename(columns=present_map)

    # Validaciones mínimas
    if "dosimeter" not in df.columns:
        raise ValueError("El archivo de Dosis no contiene la columna del dosímetro (ej. 'dosimeter', 'serial' o 'codigo').")

    # Asegurar columnas de dosis
    for dose_col in ["hp10", "hp007", "hp3"]:
        if dose_col not in df.columns:
            df[dose_col] = 0.0

    # Normalizar tipos/cadenas clave
    df["dosimeter"] = df["dosimeter"].astype(str).str.strip().str.upper()
    if "timestamp" in df.columns:
        df["timestamp"] = pd.to_datetime(df["timestamp"], errors="coerce")

    # Dejar solo columnas útiles
    keep = ["dosimeter", "timestamp", "hp10", "hp007", "hp3"]
    df = df[[c for c in keep if c in df.columns]]

    # Eliminar duplicados por dosimeter quedándonos con el más reciente
    if "timestamp" in df.columns:
        df = df.sort_values("timestamp").drop_duplicates(subset=["dosimeter"], keep="last")
    else:
        df = df.drop_duplicates(subset=["dosimeter"], keep="last")

    return df

def generar_reporte(df_final: pd.DataFrame, logo_bytes: bytes | None = None) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "REPORTE DE DOSIS"

    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # Encabezado: fecha
    ws['I1'] = f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y')}"
    ws['I1'].font = Font(size=10, italic=True)
    ws['I1'].alignment = Alignment(horizontal='right', vertical='top')

    # Logo (opcional)
    ws['A1'] = ""
    ws['A1'].font = Font(size=10)
    ws['A1'].alignment = Alignment(horizontal='left', vertical='top')

    if logo_bytes:
        try:
            logo_img = XLImage(BytesIO(logo_bytes))
            logo_img.width, logo_img.height = 240, 100
            ws.add_image(logo_img, "A1")
        except Exception:
            pass

    # Título
    ws.merge_cells('A5:J5')
    c = ws['A5']
    c.value = 'REPORTE DE DOSIMETRÍA'
    c.font = Font(bold=True, size=14)
    c.alignment = Alignment(horizontal='center')

    # Encabezados de tabla principal
    headers = [
        'PERIODO DE LECTURA', 'COMPAÑÍA', 'CÓDIGO DE DOSÍMETRO', 'NOMBRE',
        'CÉDULA', 'FECHA DE LECTURA', 'TIPO DE DOSÍMETRO', 'Hp(10)', 'Hp(0.07)', 'Hp(3)'
    ]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=7, column=i, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill('solid', fgColor='DDDDDD')
        cell.border = border

    # Filas
    start_row = 8
    for idx, row in df_final.iterrows():
        for col_idx, val in enumerate(row, 1):
            c = ws.cell(row=start_row + idx, column=col_idx, value=val)
            c.alignment = Alignment(horizontal='center', wrap_text=True)
            c.font = Font(size=10)
            c.border = border

    # Auto-ancho de columnas
    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value else 0 for c in col) + 2
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len

    # Secciones informativas
    info_start = start_row + len(df_final) + 2
    row = info_start

    ws.merge_cells(f'A{row}:P{row}')
    c = ws[f'A{row}']
    c.value = 'INFORMACIÓN DEL REPORTE DE DOSIMETRÍA'
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal='center')
    row += 1

    bullets = [
        '‒ Periodo de lectura: periodo de uso del dosímetro personal.',
        '‒ Fecha de lectura: fecha en que se realizó la lectura.',
        '‒ Tipo de dosímetro:'
    ]
    for text in bullets:
        ws.merge_cells(f'A{row}:D{row}')
        c = ws[f'A{row}']
        c.value = text
        c.font = Font(size=10, bold=True)
        c.alignment = Alignment(horizontal='left')
        row += 2

    tipos = [('CE', 'Cuerpo Entero'), ('A', 'Anillo'), ('B', 'Brazalete'), ('CR', 'Cristalino')]
    for clave, desc in tipos:
        ws.merge_cells(f'C{row}:D{row}')
        c = ws[f'C{row}']
        c.value = f"{clave} = {desc}"
        c.font = Font(size=10, bold=True)
        c.alignment = Alignment(horizontal='left')
        for col in ('C', 'D'):
            ws.cell(row=row, column=ord(col)-64).border = border
        row += 1

    row += 1
    ws.merge_cells(f'F{row}:I{row}')
    c = ws[f'F{row}']
    c.value = 'LÍMITES ANUALES DE EXPOSICIÓN A RADIACIONES'
    c.font = Font(bold=True, size=10)
    c.alignment = Alignment(horizontal='center')
    row += 1

    limites = [
        ("Cuerpo Entero", "20 mSv/año"),
        ("Cristalino", "150 mSv/año"),
        ("Extremidades y piel", "500 mSv/año"),
        ("Fetal", "1 mSv/periodo de gestación"),
        ("Público", "1 mSv/año")
    ]
    for cat, val in limites:
        ws.merge_cells(f'F{row}:G{row}')
        ws[f'F{row}'].value = cat
        ws[f'F{row}'].font = Font(size=10)
        ws[f'F{row}'].alignment = Alignment(horizontal='left')
        ws.merge_cells(f'H{row}:I{row}')
        ws[f'H{row}'].value = val
        ws[f'H{row}'].font = Font(size=10)
        ws[f'H{row}'].alignment = Alignment(horizontal='right')
        for col in ('F','G','H','I'):
            ws.cell(row=row, column=ord(col)-64).border = border
        row += 1

    row += 2
    ws.merge_cells(f'A{row}:P{row}')
    c = ws[f'A{row}']
    c.value = '‒ DATOS DEL PARTICIPANTE:'
    c.font = Font(bold=True, size=10)
    c.alignment = Alignment(horizontal='left')
    row += 1

    datos = [
        '‒ Código de usuario: Número único asignado al usuario por Microsievert, S.A.',
        '‒ Nombre: Persona a la cual se le asigna el dosímetro personal.',
        '‒ Cédula: Número del documento de identidad personal del usuario.',
        '‒ Fecha de nacimiento: Registro de la fecha de nacimiento del usuario.'
    ]
    for txt in datos:
        ws.merge_cells(f'A{row}:P{row}')
        c = ws[f'A{row}']
        c.value = txt
        c.font = Font(size=10)
        c.alignment = Alignment(horizontal='left', wrap_text=True)
        row += 1

    row += 2
    ws.merge_cells(f'A{row}:P{row}')
    c = ws[f'A{row}']
    c.value = '‒ DOSIS EN MILISIEVERT:'
    c.font = Font(bold=True, size=10)
    c.alignment = Alignment(horizontal='left')
    row += 1

    # Encabezados de definiciones
    ws.merge_cells(f'B{row}:C{row}')
    hb = ws[f'B{row}']; hb.value = 'Nombre'
    hb.font = Font(bold=True, size=10)
    hb.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    hb.fill = PatternFill('solid', fgColor='DDDDDD')

    ws.merge_cells(f'D{row}:I{row}')
    hd = ws[f'D{row}']; hd.value = 'Definición'
    hd.font = Font(bold=True, size=10)
    hd.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    hd.fill = PatternFill('solid', fgColor='DDDDDD')

    ws.merge_cells(f'J{row}:J{row}')
    hu = ws[f'J{row}']; hu.value = 'Unidad'
    hu.font = Font(bold=True, size=10)
    hu.alignment = Alignment(horizontal='center', wrap_text=True)
    hu.fill = PatternFill('solid', fgColor='DDDDDD')

    for col in ('B','C','D','E','F','G','H','I','J'):
        ws.cell(row=row, column=ord(col)-64).border = border
    ws.row_dimensions[row].height = 30
    row += 1

    definitions = [
        ("Dosis efectiva Hp(10)",  "Es la dosis equivalente en tejido blando, J·kg⁻¹ o Sv a una profundidad de 10 mm, bajo determinado punto del cuerpo.", "mSv"),
        ("Dosis superficial Hp(0,07)", "Es la dosis equivalente en tejido blando, J·kg⁻¹ o Sv a una profundidad de 0,07 mm, bajo determinado punto del cuerpo.", "mSv"),
        ("Dosis cristalino Hp(3)", "Es la dosis equivalente en tejido blando, J·kg⁻¹ o Sv a una profundidad de 3 mm, bajo determinado punto del cuerpo.", "mSv")
    ]
    for nom, desc, uni in definitions:
        ws.merge_cells(f'B{row}:C{row}')
        c = ws[f'B{row}']; c.value = nom; c.font = Font(size=10, bold=True); c.alignment = Alignment(horizontal='left', wrap_text=True)
        ws.merge_cells(f'D{row}:I{row}')
        c = ws[f'D{row}']; c.value = desc; c.font = Font(size=10); c.alignment = Alignment(horizontal='left', wrap_text=True)
        ws.merge_cells(f'J{row}:J{row}')
        c = ws[f'J{row}']; c.value = uni; c.font = Font(size=10); c.alignment = Alignment(horizontal='center', wrap_text=True)
        for col in ('B','C','D','E','F','G','H','I','J'):
            cell = ws.cell(row=row, column=ord(col)-64)
            cell.border = border
            cell.alignment = Alignment(wrap_text=True)
        ws.row_dimensions[row].height = 30
        row += 1

    row += 1
    ws.merge_cells(f'A{row}:P{row}')
    c = ws[f'A{row}']
    c.value = 'LECTURAS DE ANILLO: las lecturas del dosímetro de anillo son registradas como una dosis equivalente superficial Hp(0,07).'
    c.font = Font(size=10, bold=True)
    c.alignment = Alignment(horizontal='left', wrap_text=True)
    row += 1

    ws.merge_cells(f'A{row}:P{row}')
    c = ws[f'A{row}']
    c.value = 'Los resultados de las dosis individuales de radiación son reportados para diferentes periodos de tiempo:'
    c.font = Font(size=10)
    c.alignment = Alignment(horizontal='left', wrap_text=True)
    row += 1

    periods = [
        ('DOSIS ACTUAL', 'Es el correspondiente de dosis acumulada durante el período de lectura definido.'),
        ('DOSIS ANUAL', 'Es el correspondiente de dosis acumulada desde el inicio del año hasta la fecha.'),
        ('DOSIS DE POR VIDA', 'Es el correspondiente de dosis acumulada desde el inicio del servicio dosimétrico hasta la fecha.')
    ]
    for clave, texto in periods:
        ws.merge_cells(f'B{row}:C{row}')
        c = ws[f'B{row}']; c.value = clave; c.font = Font(bold=True, size=10); c.alignment = Alignment(horizontal='center')
        ws.merge_cells(f'D{row}:P{row}')
        c = ws[f'D{row}']; c.value = texto; c.font = Font(size=10); c.alignment = Alignment(horizontal='left', wrap_text=True)
        for col in ('B','C') + tuple(chr(x) for x in range(68, 81)):
            ws.cell(row=row, column=ord(col)-64).border = border
        row += 1

    row += 2
    ws.merge_cells(f'A{row}:P{row}')
    c = ws[f'A{row}']
    c.value = ('DOSÍMETRO DE CONTROL: incluido en cada paquete entregado para monitorear la exposición a la radiación recibida durante el tránsito y almacenamiento. '
               'Este dosímetro debe ser guardado por el cliente en un área libre de radiación durante el período de uso.')
    c.font = Font(size=10, bold=True)
    c.alignment = Alignment(horizontal='left', wrap_text=True)
    row += 2

    ws.merge_cells(f'A{row}:P{row}')
    c = ws[f'A{row}']
    c.value = ('POR DEBAJO DEL MÍNIMO DETECTADO: es la dosis por debajo de la cantidad mínima reportada para el período de uso y son registradas como "PM".')
    c.font = Font(size=10, bold=True)
    c.alignment = Alignment(horizontal='left', wrap_text=True)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# =========================
# LECTURA DE ARCHIVOS / PERÍODO
# =========================
periodo_opciones = []
dfp = None

if participantes_file:
    try:
        dfp = leer_participantes(participantes_file)
        periodos = set()
        for col in ['PERIODO 1', 'PERIODO 2']:
            if col in dfp.columns:
                periodos.update(
                    dfp[col].dropna().astype(str).str.strip().str.upper()
                )
        periodo_opciones = sorted([p for p in periodos if p and p != ''])
        if not periodo_opciones:
            st.info("No se detectaron períodos en el archivo de Participantes. Puedes continuar, pero el reporte mostrará el campo de período con lo que selecciones manualmente.")
    except Exception as e:
        st.error(f"Error leyendo participantes: {e}")

periodo_seleccionado = st.selectbox("Selecciona el período a mostrar", options=periodo_opciones if periodo_opciones else ["(sin período)"])
if periodo_seleccionado == "(sin período)":
    periodo_seleccionado = ""

nombre_reporte = st.text_input(
    "Nombre del archivo de reporte (sin extensión):",
    value=f"ReporteDosimetria_{datetime.now().strftime('%Y-%m-%d')}"
)

logo_file = st.file_uploader("Sube el logo de Microsievert (.png, opcional)", type=["png"], key="logo")
logo_bytes = logo_file.read() if logo_file else None

# =========================
# BOTÓN GENERAR
# =========================
btn_disabled = not (participantes_file and dosis_file)

if st.button("✅ Generar Reporte", disabled=btn_disabled):
    try:
        if not participantes_file or not dosis_file:
            st.error("Debes subir el archivo de Participantes y el de Dosis.")
        elif not nombre_reporte.strip():
            st.error("El nombre del archivo es obligatorio.")
        elif re.search(r'[\\/:*?"<>|]', nombre_reporte):
            st.error("El nombre del archivo contiene caracteres no permitidos.")
        else:
            dfp = leer_participantes(participantes_file)
            dfd = leer_dosis(dosis_file)

            # Índice por dosimeter para acceso rápido
            dfd_index = dfd.set_index("dosimeter")

            registros = []
            for _, fila in dfp.iterrows():
                for cod in [fila.get('DOSIMETRO 1', ''), fila.get('DOSIMETRO 2', '')]:
                    cod = (cod or "").strip().upper()
                    if not cod:
                        continue
                    if cod in dfd_index.index:
                        dosis = dfd_index.loc[cod]

                        nombre_raw = f"{fila.get('NOMBRE','')} {fila.get('APELLIDO','')}".strip()
                        nombre = "CONTROL" if "CONTROL" in nombre_raw.upper() else nombre_raw

                        fecha = dosis.get('timestamp', None)
                        if pd.notna(fecha):
                            fecha_str = pd.to_datetime(fecha, errors='coerce').strftime('%d/%m/%Y %H:%M')
                        else:
                            fecha_str = ''

                        registros.append({
                            'PERIODO DE LECTURA': periodo_seleccionado,
                            'COMPAÑÍA': fila.get('COMPAÑÍA', ''),
                            'CÓDIGO DE DOSÍMETRO': cod,
                            'NOMBRE': nombre,
                            'CÉDULA': fila.get('CÉDULA', ''),
                            'FECHA DE LECTURA': fecha_str,
                            'TIPO DE DOSÍMETRO': 'CE',
                            'Hp(10)': float(dosis.get('hp10', 0.0)),
                            'Hp(0.07)': float(dosis.get('hp007', 0.0)),
                            'Hp(3)': float(dosis.get('hp3', 0.0))
                        })

            if not registros:
                st.warning("No se encontraron coincidencias entre Participantes y Dosis (verifica que los códigos coincidan).")
            else:
                # Primer registro = CONTROL como base
                base10 = float(registros[0]['Hp(10)'])
                base07 = float(registros[0]['Hp(0.07)'])
                base3  = float(registros[0]['Hp(3)'])

                for i, r in enumerate(registros):
                    if i == 0:
                        # Mostrar el control como valor base
                        r['Hp(10)']  = f"{base10:.2f}"
                        r['Hp(0.07)'] = f"{base07:.2f}"
                        r['Hp(3)']   = f"{base3:.2f}"
                    else:
                        for key, base in [('Hp(10)', base10), ('Hp(0.07)', base07), ('Hp(3)', base3)]:
                            # Diferencia respecto al control (PM si < 0.005)
                            diff = float(r[key]) - base
                            r[key] = "PM" if diff < 0.005 else f"{diff:.2f}"

                df_final = pd.DataFrame(registros)

                # Generar Excel
                excel_bytes = generar_reporte(df_final, logo_bytes)

                st.success(f"Reporte generado con {len(df_final)} registros.")
                st.download_button(
                    label="Descargar Reporte Excel",
                    data=excel_bytes,
                    file_name=f"{nombre_reporte.strip()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"Error generando reporte: {e}")

# Footer
st.markdown(
    f"""<div style="color:{COLOR_FOOTER};text-align:center;font-size:12px;">
    Sistema de Gestión de Dosimetría - Microsievert
    </div>""",
    unsafe_allow_html=True
)
