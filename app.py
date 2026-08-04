from datetime import datetime
from flask import Flask, render_template, request, send_file
from procesador import (
    generar_excel_formateado,
    generar_pdf_reportlab,
    procesar_archivos_excel,
)

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/procesar", methods=["POST"])
def procesar():
    archivos = request.files.getlist("archivos")
    formato = request.form.get("formato", "excel")
    local = request.form.get("local", "La Benita")  # Captura local seleccionado

    if not archivos or archivos[0].filename == "":
        return "No se subieron archivos válidos.", 400

    df_procesado = procesar_archivos_excel(archivos, local)
    fecha_str = datetime.now().strftime("%d.%m")

    # Formatear el nombre del archivo de salida
    nombre_base = f"RQ_{local.replace(' ', '_').upper()}_{fecha_str}"

    if formato == "pdf":
        pdf_buffer = generar_pdf_reportlab(df_procesado, local)
        return send_file(
            pdf_buffer,
            download_name=f"{nombre_base}.pdf",
            mimetype="application/pdf",
            as_attachment=True,
        )
    else:
        excel_buffer = generar_excel_formateado(df_procesado)
        return send_file(
            excel_buffer,
            download_name=f"{nombre_base}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
        )


if __name__ == "__main__":
    app.run(debug=True)
