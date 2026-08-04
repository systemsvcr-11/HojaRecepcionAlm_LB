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
        "Maritza Alcasihuincha": "BARRA",
        "Coordinador Guardiana": "ATENCION",
        "Maria Villalobos": "CAJA",
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

    # Mapeo según el local seleccionado
    mapeo_actual = MAPEO_AREAS.get(local_seleccionado, {})

    if "Comprador" in df.columns:
        df["AREA"] = (
            df["Comprador"].map(mapeo_actual).fillna(df["Comprador"])
        )  # Mantiene el nombre si no está en el mapa
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
    """Exporta y aplica los estilos condicionales en memoria."""
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
        if ws.max_column >= 7:
            ws.cell(row=row, column=7).fill = fill_g
        if ws.max_column >= 10:
            ws.cell(row=row, column=10).fill = fill_j

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
    """Genera el PDF landscape aprovechando el 100% del ancho de la hoja."""
    pdf_buffer = io.BytesIO()

    # Configuración de página con márgenes mínimos para maximizar ancho (0.5 cm por lado)
    margin = 0.5 * cm
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(A4),
        leftMargin=margin,
        rightMargin=margin,
        topMargin=1.2 * cm,
        bottomMargin=0.8 * cm,
    )

    elements = []
    styles = getSampleStyleSheet()

    fecha_str = datetime.now().strftime("%d.%m")
    titulo_texto = (
        f"RQ {local_seleccionado.upper()} - {fecha_str}"  # Título dinámico
    )

    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=20,
        alignment=1,
        spaceAfter=12,
    )
    elements.append(Paragraph(titulo_texto, title_style))

    cols_pdf = [
        c
        for c in df.columns
        if c not in ["ID", "Líneas de la orden/Producto/ID"]
    ]
    df_pdf = df[cols_pdf]

    # Ajustar estilos de celda envolviendo texto largo en Paragraph para autofit de altura
    cell_style = ParagraphStyle("CellStyle", fontName="Helvetica", fontSize=8)
    header_style = ParagraphStyle(
        "HeaderStyle",
        fontName="Helvetica-Bold",
        fontSize=8,
        textColor=colors.whitesmoke,
        alignment=1,
    )

    formatted_data = []
    # Fila del Encabezado
    formatted_data.append(
        [Paragraph(str(col), header_style) for col in df_pdf.columns]
    )

    # Filas de Datos
    for _, row in df_pdf.iterrows():
        row_cells = []
        for val in row:
            text = "" if pd.isna(val) else str(val)
            row_cells.append(Paragraph(text, cell_style))
        formatted_data.append(row_cells)

    # Ancho total disponible en la hoja A4 apapaisada
    page_width = landscape(A4)[0] - (margin * 2)  # ~28.7 cm disponibles
    num_cols = len(cols_pdf)

    # Repartir anchos proporcionalmente entre las columnas para ocupar todo el ancho
    col_widths = [page_width / num_cols] * num_cols

    table_style = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#343A40")),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B0BEC5")),
        ]
    )

    t = Table(formatted_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(table_style)
    elements.append(t)

    doc.build(elements)
    pdf_buffer.seek(0)
    return pdf_buffer
