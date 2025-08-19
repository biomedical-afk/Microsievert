from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image as XLImage
from datetime import datetime
from io import BytesIO

def build_formatted_excel(df_final: pd.DataFrame,
                          header_lines: List[str],
                          logo_bytes: Optional[bytes]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte"

    # Estilos básicos
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    gray = PatternFill("solid", fgColor="DDDDDD")
    group_fill = PatternFill("solid", fgColor="EEEEEE")

    # Anchos/altos cabecera (para el cuadro del logo)
    widths = {
        "A": 24, "B": 28, "C": 16, "D": 16, "E": 16, "F": 16,
        "G": 10, "H": 12, "I": 12, "J": 12, "K": 12, "L": 12, "M": 12,
        "N": 12, "O": 12, "P": 12
    }
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    for r in range(1, 4 + 1):
        ws.row_dimensions[r].height = 20  # alto de filas del área del logo

    # Texto (A1:B4)
    for i, line in enumerate(header_lines[:4], start=1):
        ws.merge_cells(f"A{i}:B{i}")
        c = ws[f"A{i}"]; c.value = line; c.fill = gray
        c.font = Font(bold=True); c.alignment = Alignment(horizontal="left", vertical="center")
        for col in ("A", "B"):
            ws.cell(row=i, column=ord(col) - 64).border = border

    # Fecha de emisión (I1:P1)
    ws.merge_cells("I1:J1"); ws["I1"] = "Fecha de emisión"
    ws["I1"].font = Font(bold=True, size=10); ws["I1"].alignment = center; ws["I1"].fill = gray
    ws.merge_cells("K1:P1"); ws["K1"] = datetime.now().strftime("%d-%b-%y").lower()
    ws["K1"].font = Font(bold=True, size=10); ws["K1"].alignment = center
    for col_idx in range(ord("I")-64, ord("P")-64+1):
        ws.cell(row=1, column=col_idx).border = border

    # ---- Logo dentro de C1:F4 con ajuste de tamaño ----
    def col_pixels(letter: str) -> int:
        w = ws.column_dimensions[letter].width or 8.43
        return int(w * 7 + 5)
    def row_pixels(idx: int) -> int:
        h = ws.row_dimensions[idx].height or 15
        return int(h * 96 / 72)

    if logo_bytes:
        try:
            img = XLImage(BytesIO(logo_bytes))
            # Caja C1:F4
            box_w = sum(col_pixels(c) for c in ("C","D","E","F"))
            box_h = sum(row_pixels(r) for r in (1,2,3,4))
            max_w, max_h = box_w - 8, box_h - 8  # padding
            scale = min(max_w / img.width, max_h / img.height, 1.0)
            img.width = int(img.width * scale)
            img.height = int(img.height * scale)
            img.anchor = "C1"
            ws.add_image(img)
        except Exception:
            pass

    # ---- Título ----
    ws.merge_cells("A6:P6")
    ws["A6"] = "REPORTE DE DOSIMETRÍA"
    ws["A6"].font = Font(bold=True, size=14)
    ws["A6"].alignment = center

    # ---- Fila de grupos de dosis (cerrada con borde) ----
    ws.merge_cells("H7:J7"); ws["H7"] = "DOSIS ACTUAL (mSv)"
    ws.merge_cells("K7:M7"); ws["K7"] = "DOSIS ANUAL (mSv)"
    ws.merge_cells("N7:P7"); ws["N7"] = "DOSIS DE POR VIDA (mSv)"
    for rng in (("H7","J7"), ("K7","M7"), ("N7","P7")):
        start_col = column_index_from_string(rng[0][0]); end_col = column_index_from_string(rng[1][0])
        row = 7
        # Estilo base del título
        c = ws[rng[0]]; c.font = bold; c.alignment = center; c.fill = group_fill
        # Borde en toda la franja del rango fusionado
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = border
            cell.fill = group_fill

    # ---- Encabezados de tabla ----
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
        cell.font = bold; cell.alignment = center; cell.border = border
        cell.fill = gray if col_idx >= 8 else gray  # todos gris claro

    # ---- Datos ----
    start_row = header_row + 1
    for _, r in df_final[headers].iterrows():
        ws.append(list(r.values))
    last_row = ws.max_row

    # Bordes, alineación y ALTO UNIFORME de filas de datos
    for row in ws.iter_rows(min_row=header_row, max_row=last_row, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = border
            if cell.row >= start_row:
                cell.alignment = Alignment(vertical="center", wrap_text=True, horizontal="center")
    # Altura fija (espaciado consistente)
    for rr in range(start_row, last_row + 1):
        ws.row_dimensions[rr].height = 20  # <- cambia a 22 si los quieres un poco más altos

    ws.freeze_panes = f"A{start_row}"

    # Autoancho moderado
    for col_cells in ws.iter_cols(min_col=1, max_col=16, min_row=header_row, max_row=last_row):
        col_letter = get_column_letter(col_cells[0].column)
        max_len = max(len("" if c.value is None else str(c.value)) for c in col_cells)
        ws.column_dimensions[col_letter].width = max(ws.column_dimensions[col_letter].width, min(max_len + 2, 42))

    # ---- Sección informativa (igual que antes) ----
    row = last_row + 2
    ws.merge_cells(f"A{row}:P{row}")
    ws[f"A{row}"] = "INFORMACIÓN DEL REPORTE DE DOSIMETRÍA"
    ws[f"A{row}"].font = Font(bold=True); ws[f"A{row}"].alignment = Alignment(horizontal="center")
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
    ws[f"F{row}"] = "LÍMITES ANUALES DE EXPOSICIÓN A RADIACIONES"
    ws[f"F{row}"].font = Font(bold=True, size=10); ws[f"F{row}"].alignment = Alignment(horizontal="center")
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
    ws[f"A{row}"] = "‒ DATOS DEL PARTICIPANTE:"
    ws[f"A{row}"].font = Font(bold=True, size=10); ws[f"A{row}"].alignment = Alignment(horizontal="left")
    row += 1

    datos = [
        "‒ Código de usuario: Número único asignado al usuario por Microsievert, S.A.",
        "‒ Nombre: Persona a la cual se le asigna el dosímetro personal.",
        "‒ Cédula: Número del documento de identidad personal del usuario.",
    ]
    for txt in datos:
        ws.merge_cells(f"A{row}:P{row}")
        ws[f"A{row}"].value = txt
        ws[f"A{row}"].font = Font(size=10); ws[f"A{row}"].alignment = Alignment(horizontal="left")
        row += 1
    row += 2

    ws.merge_cells(f"A{row}:P{row}")
    ws[f"A{row}"] = "‒ DOSIS EN MILISIEVERT:"
    ws[f"A{row}"].font = Font(bold=True, size=10); ws[f"A{row}"].alignment = Alignment(horizontal="left")
    row += 1

    shade = PatternFill("solid", fgColor="DDDDDD")
    ws.merge_cells(f"B{row}:C{row}"); ws[f"B{row}"] = "Nombre"
    ws[f"B{row}"].font = Font(bold=True, size=10)
    ws[f"B{row}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); ws[f"B{row}"].fill = shade
    ws.merge_cells(f"D{row}:I{row}"); ws[f"D{row}"] = "Definición"
    ws[f"D{row}"].font = Font(bold=True, size=10)
    ws[f"D{row}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); ws[f"D{row}"].fill = shade
    ws.merge_cells(f"J{row}:J{row}"); ws[f"J{row}"] = "Unidad"
    ws[f"J{row}"].font = Font(bold=True, size=10)
    ws[f"J{row}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); ws[f"J{row}"].fill = shade
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
        ws.merge_cells(f"B{row}:C{row}"); ws[f"B{row}"] = nom
        ws[f"B{row}"].font = Font(size=10, bold=True); ws[f"B{row}"].alignment = Alignment(horizontal="left", wrap_text=True)
        ws.merge_cells(f"D{row}:I{row}"); ws[f"D{row}"] = desc
        ws[f"D{row}"].font = Font(size=10); ws[f"D{row}"].alignment = Alignment(horizontal="left", wrap_text=True)
        ws.merge_cells(f"J{row}:J{row}"); ws[f"J{row}"] = uni
        ws[f"J{row}"].font = Font(size=10); ws[f"J{row}"].alignment = Alignment(horizontal="center", wrap_text=True)
        for col in ("B","C","D","E","F","G","H","I","J"):
            ws.cell(row=row, column=ord(col)-64).border = border
        ws.row_dimensions[row].height = 30
        row += 1

    row += 1
    ws.merge_cells(f"A{row}:P{row}")
    ws[f"A{row}"] = "LECTURAS DE ANILLO: las lecturas del dosímetro de anillo son registradas como una dosis equivalente superficial Hp(0,07)."
    ws[f"A{row}"].font = Font(size=10, bold=True); ws[f"A{row}"].alignment = Alignment(horizontal="left", wrap_text=True)
    row += 1

    ws.merge_cells(f"A{row}:P{row}")
    ws[f"A{row}"] = "Los resultados de las dosis individuales de radiación son reportados para diferentes periodos de tiempo:"
    ws[f"A{row}"].font = Font(size=10); ws[f"A{row}"].alignment = Alignment(horizontal="left", wrap_text=True)
    row += 1

    blocks = [
        ("DOSIS ACTUAL",      "Es el correspondiente de dosis acumulada durante el período de lectura definido."),
        ("DOSIS ANUAL",       "Es el correspondiente de dosis acumulada desde el inicio del año hasta la fecha."),
        ("DOSIS DE POR VIDA", "Es el correspondiente de dosis acumulada desde el inicio del servicio dosimétrico hasta la fecha."),
    ]
    for clave, texto in blocks:
        ws.merge_cells(f"B{row}:C{row}"); ws[f"B{row}"] = clave
        ws[f"B{row}"].font = Font(bold=True, size=10); ws[f"B{row}"].alignment = Alignment(horizontal="center")
        ws.merge_cells(f"D{row}:P{row}"); ws[f"D{row}"] = texto
        ws[f"D{row}"].font = Font(size=10); ws[f"D{row}"].alignment = Alignment(horizontal="left", wrap_text=True)
        for col_idx in range(ord("B")-64, ord("P")-64+1):
            ws.cell(row=row, column=col_idx).border = border
        row += 1

    row += 2
    ws.merge_cells(f"A{row}:P{row}")
    ws[f"A{row}"] = ("DOSÍMETRO DE CONTROL: incluido en cada paquete entregado para monitorear la exposición a la radiación "
                     "recibida durante el tránsito y almacenamiento. Este dosímetro debe ser guardado por el cliente en un "
                     "área libre de radiación durante el período de uso.")
    ws[f"A{row}"].font = Font(size=10, bold=True); ws[f"A{row}"].alignment = Alignment(horizontal="left", wrap_text=True)
    row += 2

    ws.merge_cells(f"A{row}:P{row}")
    ws[f"A{row}"] = ("POR DEBAJO DEL MÍNIMO DETECTADO: es la dosis por debajo de la cantidad mínima reportada para el período "
                     "de uso y son registradas como \"PM\".")
    ws[f"A{row}"].font = Font(size=10, bold=True); ws[f"A{row}"].alignment = Alignment(horizontal="left", wrap_text=True)

    bio = BytesIO(); wb.save(bio); bio.seek(0)
    return bio.getvalue()


