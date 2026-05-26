import os
import subprocess
import tempfile
from flask import Blueprint, request, jsonify, send_file
from openpyxl import load_workbook
import datetime

certbot_bp = Blueprint('certbot', __name__)

HOJA_CERT = {
    "TORQUIMETRO": "CERTIFICADO",
}

def detectar_tipo(nombre):
    nombre = nombre.upper()
    for tipo in HOJA_CERT:
        if tipo in nombre:
            return tipo
    return None

def extraer_datos(ruta_excel):
    wb = load_workbook(ruta_excel, read_only=True, data_only=True)
    ws = wb["REGISTRO"]
    def cel(coord):
        val = ws[coord].value
        return str(val).strip() if val is not None else ""
    datos = {
        "n_certificado": cel("J7"),
        "orden_trabajo": cel("J5"),
        "solicitante": cel("J9"),
        "instrumento": cel("J12"),
    }
    wb.close()
    return datos

def construir_nombre(datos, tipo):
    cert  = datos.get("n_certificado", "CERT").replace("/", "-")
    ot    = datos.get("orden_trabajo", "")
    inst  = datos.get("instrumento", tipo)[:15]
    fecha = datetime.date.today().strftime("%Y%m%d")
    nombre = f"{cert}_{inst}_{ot}_{fecha}.pdf"
    for c in ['\\','/',':','*','?','"','<','>','|',' ']:
        nombre = nombre.replace(c, '_')
    return nombre

@certbot_bp.route('/generar-certificado', methods=['POST'])
def generar_certificado():
    if 'file' not in request.files:
        return jsonify({"error": "No se envió archivo"}), 400

    archivo = request.files['file']
    nombre  = archivo.filename

    tipo = detectar_tipo(nombre)
    if tipo is None:
        return jsonify({"error": f"Tipo no reconocido: {nombre}"}), 400

    with tempfile.TemporaryDirectory() as tmpdir:
        ruta_excel = os.path.join(tmpdir, nombre)
        archivo.save(ruta_excel)

        datos = extraer_datos(ruta_excel)

        cmd = [
            "libreoffice", "--headless",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            ruta_excel
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if result.returncode != 0:
            return jsonify({"error": "Error LibreOffice", "detalle": result.stderr}), 500

        pdf_generado = os.path.join(tmpdir, os.path.splitext(nombre)[0] + ".pdf")
        if not os.path.exists(pdf_generado):
            return jsonify({"error": "PDF no generado"}), 500

        try:
            from pypdf import PdfReader, PdfWriter
            reader = PdfReader(pdf_generado)
            total  = len(reader.pages)
            inicio = max(0, total - 3)
            writer = PdfWriter()
            for i in range(inicio, total):
                writer.add_page(reader.pages[i])
            pdf_final = os.path.join(tmpdir, construir_nombre(datos, tipo))
            with open(pdf_final, "wb") as f:
                writer.write(f)
        except Exception:
            pdf_final = pdf_generado

        return send_file(
            pdf_final,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=os.path.basename(pdf_final)
        )
