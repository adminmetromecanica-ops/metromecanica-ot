import os
import subprocess
import tempfile
from flask import Blueprint, request, jsonify, send_file
from openpyxl import load_workbook
import datetime

certbot_bp = Blueprint('certbot', __name__)

HOJA_CERT = {
    "TORQUIMETRO": "CERTIFICADO",
    "BALANZA":     "CERTIFICADO",
    "PESAS":       "CERTIFICADO",
    "PRESION":     "CERTIFICADO",
    "TEMPERATURA": "CERTIFICADO",
    "FUERZA":      "CERTIFICADO",
    "LONGITUD":    "CERTIFICADO",
    "ENERGIA":     "CERTIFICADO",
    "MEDICO":      "CERTIFICADO",
    "QUIMICA":     "CERTIFICADO",
    "ENSAYO":      "CERTIFICADO",
    "OTROS":       "CERTIFICADO",
}

HOJAS_DATOS = ["REGISTRO", "DATOS_EQUIPO", "DATOS", "EQUIPO", "INFO", "INFORMACION"]

def detectar_tipo(nombre):
    nombre = nombre.upper()
    for tipo in HOJA_CERT:
        if tipo in nombre:
            return tipo
    return None

def leer_celda(ws, coord):
    val = ws[coord].value
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ["none", "nan", ""]:
        return ""
    return s

def extraer_datos(ruta_excel):
    try:
        # Intentar primero con data_only=True
        wb = load_workbook(ruta_excel, read_only=True, data_only=True)
        ws = None
        for nombre_hoja in HOJAS_DATOS:
            if nombre_hoja in wb.sheetnames:
                ws = wb[nombre_hoja]
                break
        if ws is None:
            ws = wb.worksheets[0]

        datos = {
            "n_certificado": leer_celda(ws, "J7"),
            "orden_trabajo": leer_celda(ws, "J5"),
            "solicitante":   leer_celda(ws, "J9"),
            "instrumento":   leer_celda(ws, "J12"),
        }
        wb.close()

        # Si J7 vacío buscar en B7
        if not datos["n_certificado"]:
            wb2 = load_workbook(ruta_excel, read_only=True, data_only=True)
            ws2 = None
            for nombre_hoja in HOJAS_DATOS:
                if nombre_hoja in wb2.sheetnames:
                    ws2 = wb2[nombre_hoja]
                    break
            if ws2 is None:
                ws2 = wb2.worksheets[0]
            datos["n_certificado"] = leer_celda(ws2, "B7")
            wb2.close()

        # Si aun vacío buscar en todo el libro
        if not datos["n_certificado"]:
            wb3 = load_workbook(ruta_excel, read_only=True, data_only=True)
            for sheet in wb3.worksheets:
                for row in sheet.iter_rows(min_row=1, max_row=20, values_only=True):
                    for cell in row:
                        if cell and isinstance(cell, str):
                            c = cell.strip()
                            if len(c) > 5 and ('-' in c) and any(p in c.upper() for p in ['MLL','MLF','MLE','MLP','MLT','MLM','MLQ','MLC']):
                                datos["n_certificado"] = c
                                break
                    if datos["n_certificado"]:
                        break
                if datos["n_certificado"]:
                    break
            wb3.close()

        return datos
    except Exception:
        return {
            "n_certificado": "CERT",
            "orden_trabajo": "",
            "solicitante":   "",
            "instrumento":   "",
        }

def construir_nombre(datos, tipo):
    cert  = datos.get("n_certificado", "CERT")
    inst  = datos.get("instrumento", "")
    solic = datos.get("solicitante", "")
    ot    = datos.get("orden_trabajo", "")
    fecha = datetime.date.today().strftime("%Y%m%d")

    if not cert or cert.lower() in ["none", "nan", ""]:
        cert = "CERT"
    if not inst or inst.lower() in ["none", "nan"]:
        inst = ""
    if not solic or solic.lower() in ["none", "nan"]:
        solic = ""
    if not ot or ot.lower() in ["none", "nan"]:
        ot = ""

    nombre = f"{cert}_{inst}_{solic}_{ot}_{fecha}.pdf"
    for c in ['\\', '/', ':', '*', '?', '"', '<', '>', '|', ' ', '\n', '\r']:
        nombre = nombre.replace(c, '_')
    while '__' in nombre:
        nombre = nombre.replace('__', '_')
    if len(nombre) > 150:
        nombre = nombre[:146] + ".pdf"
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

        env = os.environ.copy()
        env["LANG"]       = "es_PE.UTF-8"
        env["LC_ALL"]     = "es_PE.UTF-8"
        env["LC_NUMERIC"] = "es_PE.UTF-8"

        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            ruta_excel
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60, env=env)
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
