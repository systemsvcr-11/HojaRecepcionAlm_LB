from datetime import datetime
import io
from copy import copy
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, PatternFill, Side
from openpyxl.utils import get_column_letter

# ReportLab para la generación de PDF nativa en Linux
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


def procesar_archivos_excel(archivos_subidos):
    """Procesa los archivos Excel recibidos vía Flask y retorna un DataFrame procesado."""
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

    if "Comprador" in df.columns:
        df["AREA"] = (
            df["Comprador"]
            .map(
                {
                    "COCINA BENITA": "COCINA",
                    "Gisela Benitez": "ATENCION",
                    "Ariana Maquera": "BARRA",
                    "Angie Apaza": "CAJA",
                }
            )
            .fillna("")
        )
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
    """Aplica formatos openpyxl en memoria y retorna el archivo BytesIO."""
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

    # Colorear columna RQ y AREA por valor único
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

    # Ocultar columnas ID si existen
    for col_hid in ["ID", "Líneas de la orden/Producto/ID"]:
        if col_hid in headers:
            ws.column_dimensions[
                get_column_letter(headers[col_hid])
            ].hidden = True

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output


def generar_pdf_reportlab(df):
    """Genera un archivo PDF landscape usando ReportLab."""
    pdf_buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(A4),
        leftMargin=0.5 * cm,
        rightMargin=0.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.0 * cm,
    )

    elements = []
    styles = getSampleStyleSheet()

    fecha_str = datetime.now().strftime("%d.%m")
    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=22,
        alignment=1,
        spaceAfter=15,
    )

    elements.append(Paragraph(f"RQ BENITA {fecha_str}", title_style))

    # Ocultar columnas ID para la tabla del PDF
    cols_pdf = [
        c
        for c in df.columns
        if c not in ["ID", "Líneas de la orden/Producto/ID"]
    ]
    df_pdf = df[cols_pdf]

    data = [df_pdf.columns.tolist()] + df_pdf.astype(str).values.tolist()

    # Estilos de la tabla
    table_style = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#343A40")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
            ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 1), (-1, -1), 8),
        ]
    )

    t = Table(data, repeatRows=1)
    t.setStyle(table_style)
    elements.append(t)

    doc.build(elements)
    pdf_buffer.seek(0)
    return pdf_buffer