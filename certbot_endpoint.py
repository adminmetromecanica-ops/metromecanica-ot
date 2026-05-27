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
CELDAS_CERT = ["B8", "B5", "J7", "B7", "J5"]

def detectar_tipo(nombre):
    nombre = nombre.upper()
    for tipo in HOJA_CERT:
        if tipo in nombre:
            return tipo
    return None

def leer_celda(ws, coord):
    try:
        val = ws[coord].value
        if val is None:
            return ""
        s = str(val).strip()
        if s.lower() in ["none", "nan", ""]:
            return ""
        return s
    except Exception:
        return ""

def extraer_datos(ruta_excel):
    try:
        wb = load_workbook(ruta_excel, read_only=True, data_only=True)
        ws = None
        for nombre_hoja in HOJAS_DATOS:
            if nombre_hoja in wb.sheetnames:
                ws = wb[nombre_hoja]
                break
        if ws is None:
            ws = wb.worksheets[0]

        n_cert = ""
        for celda in CELDAS_CERT:
            val = leer_celda(ws, celda)
            if val and len(val) > 3 and '-' in val:
                n_cert = val
                break

        if not n_cert:
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(min_row=1, max_row=25, values_only=True):
                    for cell in row:
                        if cell and isinstance(cell, str):
                            c = cell.strip()
                            if len(c) > 5 and '-' in c and any(
                                c.upper().startswith(p) for p in
                                ['MLL','MLF','MLE','MLP','MLT','MLM','MLQ','MLC','MLB']
                            ):
                                n_cert = c
                                break
                    if n_cert:
                        break
                if n_cert:
                    break

        ot      = leer_celda(ws, "D5") or leer_celda(ws, "J5") or leer_celda(ws, "B8")
        solic   = leer_celda(ws, "B9") or leer_celda(ws, "J9") or leer_celda(ws, "B11")
        instrum = leer_celda(ws, "B15") or leer_celda(ws, "J12") or leer_celda(ws, "B16") or leer_celda(ws, "B17")

        wb.close()
        return {
            "n_certificado": n_cert,
            "orden_trabajo": ot,
            "solicitante":   solic,
            "instrumento":   instrum,
        }
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

        # Convertir SOLO la hoja CERTIFICADO a PDF
        # Usando LibreOffice con filtro para seleccionar hoja específica
        try:
            from openpyxl import load_workbook as lw
            wb_temp = lw(ruta_excel)
            sheetnames = wb_temp.sheetnames
            wb_temp.close()

            # Crear Excel temporal con SOLO la hoja CERTIFICADO
            wb_solo = lw(ruta_excel)
            hojas_borrar = [s for s in wb_solo.sheetnames if s != "CERTIFICADO"]
            for h in hojas_borrar:
                del wb_solo[h]
            
            ruta_cert_only = os.path.join(tmpdir, "cert_only.xlsx")
            wb_solo.save(ruta_cert_only)
            wb_solo.close()

            cmd = [
                "libreoffice", "--headless",
                "--convert-to", "pdf",
                "--outdir", tmpdir,
                ruta_cert_only
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=90, env=env)

            if result.returncode != 0:
                raise Exception(result.stderr)

            pdf_generado = os.path.join(tmpdir, "cert_only.pdf")

        except Exception:
            # Fallback: convertir libro completo y tomar última página
            cmd = [
                "libreoffice", "--headless",
                "--convert-to", "pdf",
                "--outdir", tmpdir,
                ruta_excel
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=90, env=env)
            pdf_generado = os.path.join(tmpdir, os.path.splitext(nombre)[0] + ".pdf")

        if not os.path.exists(pdf_generado):
            return jsonify({"error": "PDF no generado"}), 500

        pdf_final = os.path.join(tmpdir, construir_nombre(datos, tipo))

        try:
            from pypdf import PdfReader, PdfWriter
            reader = PdfReader(pdf_generado)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
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
