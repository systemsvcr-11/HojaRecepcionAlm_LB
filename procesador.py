import io
from copy import copy
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, PatternFill, Side
from openpyxl.utils import get_column_letter

# ReportLab para PDF
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle

# Mapeos de áreas por comprador según el local seleccionado
MAPEO_AREAS = {
    "La Benita": {
        "COCINA BENITA": "COCINA",
        "Gisela Benitez": "ATENCION",
        "Ariana Maquera": "BARRA",
        "Angie Apaza": "CAJA",
    },
    "La Guardiana": {
        "COCINA GUARDIANA": "COCINA",
        "BARRA GUARDIANA": "BARRA",
        "CAJA GUARDIANA": "CAJA",
    },
    "Centro de Producción": {
        "PRODUCCION": "PRODUCCION",
        "ALMACEN": "ALMACEN",
    },
}


def procesar_archivos_excel(archivos_subidos, local_seleccionado):
    """Procesa los archivos Excel aplicando el mapeo según el local elegido."""
    lista_df = [pd.read_excel(f) for f in archivos_subidos]
    df = pd.concat(lista_df, ignore_index=True)

    columnas_ffill = [
        "ID",
        "Referencia de la orden",
        "Cliente",
        "Comprador",
        "Fecha esperada",
        "Líneas de la orden/Producto",
        "Líneas de la orden/Cantidad",
        "Líneas de la orden/Unidad",
        "Líneas de la orden/Descripción",
        "Líneas de la orden/Producto/ID",
    ]

    cols_existentes = [c for c in columnas_ffill if c in df.columns]
    df[cols_existentes] = df[cols_existentes].ffill()

    mapeo_actual = MAPEO_AREAS.get(local_seleccionado, {})

    if "Comprador" in df.columns:
        df["AREA"] = df["Comprador"].map(mapeo_actual).fillna(df["Comprador"])
        df = df.rename(columns={"Comprador": "Local"})

    df["RQ"] = (
        df["Referencia de la orden"]
        if "Referencia de la orden" in df.columns
        else ""
    )

    if "Fecha esperada" in df.columns:
        df["Fecha esperada"] = pd.to_datetime(
            df["Fecha esperada"], errors="coerce"
        ).dt.strftime("%d/%m/%Y")

    df = df.rename(
        columns={
            "Líneas de la orden/Producto": "Producto",
            "Líneas de la orden/Unidad": "Unidad",
            "Líneas de la orden/Cantidad": "Cantidad",
            "Líneas de la orden/Descripción": "Observaciones",
        }
    )

    df["Ubicacion"] = ""
    df["Cant 1"] = ""
    df["Falta"] = ""

    cols_deseadas = [
        "RQ",
        "AREA",
        "Fecha esperada",
        "Producto",
        "Unidad",
        "Ubicacion",
        "Cantidad",
        "Cant 1",
        "Falta",
        "Observaciones",
        "ID",
        "Líneas de la orden/Producto/ID",
    ]

    cols_finales = [c for c in cols_deseadas if c in df.columns]
    df = df[cols_finales]

    if "Cantidad" in df.columns:
        df = df[df["Cantidad"] != 0.0]

    if "Producto" in df.columns and "Observaciones" in df.columns:
        df.loc[df["Producto"] == df["Observaciones"], "Observaciones"] = ""
        df["Observaciones"] = df.apply(
            lambda row: ""
            if pd.isna(row["Observaciones"]) or pd.isna(row["Producto"])
            else str(row["Observaciones"])
            .replace(str(row["Producto"]), "")
            .replace("\n", " "),
            axis=1,
        )

    return df


def generar_excel_formateado(df):
    """Exporta y aplica los estilos en el archivo Excel."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Datos")

    output.seek(0)
    wb = load_workbook(output)
    ws = wb["Datos"]

    fill_g = PatternFill(
        start_color="D8E4BC", end_color="D8E4BC", fill_type="solid"
    )
    fill_j = PatternFill(
        start_color="FCD5B4", end_color="FCD5B4", fill_type="solid"
    )
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    center_align = Alignment(horizontal="center", vertical="center")

    for row in range(2, ws.max_row + 1):
        if ws.max_column >= 6:
            ws.cell(row=row, column=6).fill = fill_g  # Ubicacion
        if ws.max_column >= 9:
            ws.cell(row=row, column=9).fill = fill_j  # Falta

    for row in ws.iter_rows(
        min_row=1, max_row=ws.max_row, max_col=ws.max_column
    ):
        for cell in row:
            cell.border = thin_border

    headers = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}

    for col_name in ["Fecha esperada", "Unidad", "Cantidad"]:
        if col_name in headers:
            col_idx = headers[col_name]
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col_idx).alignment = center_align

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max_len + 3

    colores = [
        "D9EAD3",
        "D0E0E3",
        "FCE5CD",
        "EAD1DC",
        "FFF2CC",
        "D9D2E9",
        "CFE2F3",
    ]
    mapa_colores = {}
    color_idx = 0

    for row in range(2, ws.max_row + 1):
        valor = ws.cell(row=row, column=1).value
        if valor not in mapa_colores:
            mapa_colores[valor] = colores[color_idx % len(colores)]
            color_idx += 1

        fill = PatternFill(
            start_color=mapa_colores[valor],
            end_color=mapa_colores[valor],
            fill_type="solid",
        )
        ws.cell(row=row, column=1).fill = fill
        if ws.max_column >= 2:
            ws.cell(row=row, column=2).fill = fill

    for col_hid in ["ID", "Líneas de la orden/Producto/ID"]:
        if col_hid in headers:
            ws.column_dimensions[
                get_column_letter(headers[col_hid])
            ].hidden = True

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output


def generar_pdf_reportlab(df, local_seleccionado):
    """Genera el PDF con réplica de colores, centrados y título en todas las hojas."""
    pdf_buffer = io.BytesIO()
    margin = 0.5 * cm

    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(A4),
        leftMargin=margin,
        rightMargin=margin,
        topMargin=1.8 * cm,  # Espacio superior reservado para el título recurrente
        bottomMargin=0.8 * cm,
    )

    fecha_str = datetime.now().strftime("%d.%m")
    titulo_texto = f"RQ {local_seleccionado.upper()} - {fecha_str}"

    # Función ejecutada en cada página para estampar el título del reporte
    def agregar_encabezado_pagina(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica-Bold", 16)
        canvas.drawCentredString(landscape(A4)[0] / 2.0, landscape(A4)[1] - 1.3 * cm, titulo_texto)
        canvas.restoreState()

    elements = []

    cols_pdf = [
        c
        for c in df.columns
        if c not in ["ID", "Líneas de la orden/Producto/ID"]
    ]
    df_pdf = df[cols_pdf]

    # Paleta de colores idéntica a la de Excel
    HEX_COLORES_RQ = [
        "#D9EAD3",
        "#D0E0E3",
        "#FCE5CD",
        "#EAD1DC",
        "#FFF2CC",
        "#D9D2E9",
        "#CFE2F3",
    ]
    mapa_colores_rq = {}
    color_idx = 0

    for valor in df_pdf["RQ"].unique():
        mapa_colores_rq[valor] = HEX_COLORES_RQ[color_idx % len(HEX_COLORES_RQ)]
        color_idx += 1

    # Estilos tipográficos
    style_header = ParagraphStyle(
        "PDFHeader",
        fontName="Helvetica-Bold",
        fontSize=8,
        alignment=1,  # Centrado
        textColor=colors.whitesmoke,
    )

    style_cell_left = ParagraphStyle(
        "PDFCellLeft", fontName="Helvetica", fontSize=7, alignment=0
    )
    style_cell_center = ParagraphStyle(
        "PDFCellCenter", fontName="Helvetica", fontSize=7, alignment=1
    )

    cols_centradas = ["Fecha esperada", "Unidad", "Cantidad"]

    formatted_data = []

    # Encabezado
    formatted_data.append([Paragraph(str(c), style_header) for c in cols_pdf])

    # Construcción de celdas con sus alineaciones correspondientes
    for _, row in df_pdf.iterrows():
        row_cells = []
        for col_name in cols_pdf:
            text = "" if pd.isna(row[col_name]) else str(row[col_name])
            style = (
                style_cell_center
                if col_name in cols_centradas
                else style_cell_left
            )
            row_cells.append(Paragraph(text, style))
        formatted_data.append(row_cells)

    # Definir formato estético de la tabla (Bordes y Rellenos de celdas)
    table_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#343A40")),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#A0A0A0")),
    ]

    # Aplicar rellenos por celda replicando el Excel
    idx_ubicacion = cols_pdf.index("Ubicacion") if "Ubicacion" in cols_pdf else -1
    idx_falta = cols_pdf.index("Falta") if "Falta" in cols_pdf else -1

    for row_idx, row in enumerate(df_pdf.iterrows(), start=1):
        rq_val = row[1]["RQ"]
        hex_color = mapa_colores_rq.get(rq_val, "#FFFFFF")

        # Color por grupo de RQ en columnas RQ (0) y AREA (1)
        table_styles.append(
            ("BACKGROUND", (0, row_idx), (1, row_idx), colors.HexColor(hex_color))
        )

        # Color verde para columna Ubicacion
        if idx_ubicacion != -1:
            table_styles.append(
                (
                    "BACKGROUND",
                    (idx_ubicacion, row_idx),
                    (idx_ubicacion, row_idx),
                    colors.HexColor("#D8E4BC"),
                )
            )

        # Color naranja/salmón para columna Falta
        if idx_falta != -1:
            table_styles.append(
                (
                    "BACKGROUND",
                    (idx_falta, row_idx),
                    (idx_falta, row_idx),
                    colors.HexColor("#FCD5B4"),
                )
            )

    # Distribución proporcional de anchos
    page_width = landscape(A4)[0] - (margin * 2)
    widths_ratio = {
        "RQ": 0.08,
        "AREA": 0.08,
        "Fecha esperada": 0.08,
        "Producto": 0.28,
        "Unidad": 0.06,
        "Ubicacion": 0.07,
        "Cantidad": 0.07,
        "Cant 1": 0.06,
        "Falta": 0.06,
        "Observaciones": 0.16,
    }

    col_widths = [
        page_width * widths_ratio.get(col, 1 / len(cols_pdf)) for col in cols_pdf
    ]

    t = Table(formatted_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle(table_styles))
    elements.append(t)

    # Construir PDF vinculando la plantilla del título superior en cada página
    doc.build(
        elements,
        onFirstPage=agregar_encabezado_pagina,
        onLaterPages=agregar_encabezado_pagina,
    )

    pdf_buffer.seek(0)
    return pdf_buffer
