import os
import subprocess
import tempfile
import datetime
from flask import Blueprint, request, jsonify, send_file
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

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

def detectar_tipo(nombre):
    nombre = nombre.upper()
    for tipo in HOJA_CERT:
        if tipo in nombre:
            return tipo
    return None

def leer_celda_val(ws, coord):
    try:
        val = ws[coord].value
        if val is None:
            return ""
        if isinstance(val, datetime.datetime):
            return val.strftime("%Y-%m-%d")
        s = str(val).strip()
        return "" if s.lower() in ["none", "nan"] else s
    except Exception:
        return ""

def extraer_datos_calibracion(ruta_excel):
    """
    Lee datos desde pestaña CALIBRACION celdas B150-B154.
    Estructura estándar Metromecanica:
    B150 = N° Certificado
    B151 = Magnitud
    B152 = Equipo
    B153 = Cliente
    B154 = OT
    """
    try:
        wb = load_workbook(ruta_excel, read_only=True, data_only=True)
        
        # Buscar hoja CALIBRACION
        ws = None
        for nombre_hoja in wb.sheetnames:
            if nombre_hoja.upper() in ["CALIBRACION", "CALIBRACIÓN"]:
                ws = wb[nombre_hoja]
                break
        
        if ws is None:
            # Fallback: buscar en REGISTRO
            for nombre_hoja in ["REGISTRO", "DATOS_EQUIPO", "DATOS"]:
                if nombre_hoja in wb.sheetnames:
                    ws = wb[nombre_hoja]
                    break
        
        if ws is None:
            ws = wb.worksheets[0]

        # Leer desde CALIBRACION B150-B154
        n_cert  = leer_celda_val(ws, "B150")
        magnitud = leer_celda_val(ws, "B151")
        equipo  = leer_celda_val(ws, "B152")
        cliente = leer_celda_val(ws, "B153")
        ot      = leer_celda_val(ws, "B154")

        # Si no encontró en B150, buscar en celdas anteriores
        if not n_cert:
            for celda in ["B8", "B5", "J7"]:
                val = leer_celda_val(ws, celda)
                if val and '-' in val and len(val) > 5:
                    n_cert = val
                    break

        wb.close()
        return {
            "n_certificado": n_cert,
            "magnitud":      magnitud,
            "instrumento":   equipo,
            "solicitante":   cliente,
            "orden_trabajo": ot,
        }
    except Exception as e:
        return {
            "n_certificado": "CERT",
            "magnitud":      "",
            "instrumento":   "",
            "solicitante":   "",
            "orden_trabajo": "",
        }

def detectar_tipo_desde_datos(datos, nombre_archivo):
    """Detecta el tipo primero desde CALIBRACION!B151, luego desde nombre."""
    magnitud = datos.get("magnitud", "").upper()
    
    # Mapa de magnitudes a tipos CertBot
    mapa = {
        "LONGITUD":    "LONGITUD",
        "MASA":        "BALANZA",
        "BALANZA":     "BALANZA",
        "PESAS":       "PESAS",
        "PRESION":     "PRESION",
        "PRESIÓN":     "PRESION",
        "TEMPERATURA": "TEMPERATURA",
        "FUERZA":      "FUERZA",
        "TORQUE":      "FUERZA",
        "ENERGIA":     "ENERGIA",
        "ENERGÍA":     "ENERGIA",
        "ELECTRICA":   "ENERGIA",
        "ELÉCTRICA":   "ENERGIA",
        "MEDICO":      "MEDICO",
        "MÉDICO":      "MEDICO",
        "QUIMICA":     "QUIMICA",
        "QUÍMICA":     "QUIMICA",
        "ENSAYO":      "ENSAYO",
    }
    
    for key, tipo in mapa.items():
        if key in magnitud:
            return tipo
    
    # Fallback: buscar en nombre del archivo
    nombre_up = nombre_archivo.upper()
    for tipo in HOJA_CERT:
        if tipo in nombre_up:
            return tipo
    
    return "OTROS"

def construir_nombre(datos, nombre_archivo=""):
    """Construye nombre del PDF desde datos de CALIBRACION."""
    cert    = datos.get("n_certificado", "CERT")
    magnitud = datos.get("magnitud", "")
    inst    = datos.get("instrumento", "")
    solic   = datos.get("solicitante", "")
    ot      = datos.get("orden_trabajo", "")
    fecha   = datetime.date.today().strftime("%Y%m%d")

    # Limpiar valores
    if not cert or cert.lower() in ["none","nan",""]: cert = "CERT"
    if not magnitud or magnitud.lower() in ["none","nan"]: magnitud = ""
    if not inst or inst.lower() in ["none","nan"]: inst = ""
    if not solic or solic.lower() in ["none","nan"]: solic = ""
    if not ot or ot.lower() in ["none","nan"]: ot = ""

    # Si no hay magnitud del Excel, sacarla del nombre del archivo
    if not magnitud:
        for tipo in HOJA_CERT:
            if tipo in nombre_archivo.upper():
                magnitud = tipo
                break

    nombre = f"{cert}_{magnitud}_{inst}_{solic}_{ot}_{fecha}.pdf"
    for c in ['\\','/',':', '*','?','"','<','>','|',' ','\n','\r']:
        nombre = nombre.replace(c, '_')
    while '__' in nombre:
        nombre = nombre.replace('__', '_')
    if len(nombre) > 180:
        nombre = nombre[:176] + ".pdf"
    return nombre

def escribir_celda(ws, coord, valor):
    try:
        cell = ws[coord]
        if isinstance(cell, MergedCell):
            for rng in ws.merged_cells.ranges:
                if coord in rng:
                    master = ws.cell(row=rng.min_row, column=rng.min_col)
                    master.value = valor
                    return
        else:
            cell.value = valor
    except Exception:
        pass

def formatear_decimal(val, decimales=2):
    try:
        if val is None or val == "":
            return ""
        f = float(val)
        r = round(f, decimales)
        if decimales == 0 or r == int(r):
            return str(int(r))
        return f"{r:.{decimales}f}".replace(".", ",")
    except Exception:
        return str(val) if val else ""

def convertir_excel_a_pdf(ruta_excel, tmpdir, env):
    """
    Convierte la hoja CERTIFICADO del Excel a PDF.
    Estrategia: extraer solo la hoja CERTIFICADO y convertir.
    """
    try:
        wb = load_workbook(ruta_excel)
        
        # Verificar que existe hoja CERTIFICADO
        cert_sheet = None
        for nombre in wb.sheetnames:
            if nombre.upper() == "CERTIFICADO":
                cert_sheet = nombre
                break
        
        if cert_sheet is None:
            # Usar última hoja si no hay CERTIFICADO
            cert_sheet = wb.sheetnames[-1]
        
        # Hacer visible la hoja CERTIFICADO
        wb[cert_sheet].sheet_state = "visible"
        
        # Eliminar otras hojas
        hojas_borrar = [s for s in wb.sheetnames if s != cert_sheet]
        for h in hojas_borrar:
            try:
                wb[h].sheet_state = "visible"
                del wb[h]
            except Exception:
                pass
        
        ruta_cert = os.path.join(tmpdir, "certificado_final.xlsx")
        wb.save(ruta_cert)
        wb.close()

        # Convertir a PDF con LibreOffice
        cmd = [
            "libreoffice", "--headless",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            ruta_cert
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=90, env=env)
        
        if result.returncode != 0:
            return None, result.stderr
        
        pdf_path = os.path.join(tmpdir, "certificado_final.pdf")
        return pdf_path if os.path.exists(pdf_path) else None, None
        
    except Exception as e:
        return None, str(e)

@certbot_bp.route('/generar-certificado', methods=['POST'])
def generar_certificado():
    if 'file' not in request.files:
        return jsonify({"error": "No se envió archivo"}), 400
    
    archivo = request.files['file']
    nombre  = archivo.filename

    with tempfile.TemporaryDirectory() as tmpdir:
        ruta_excel = os.path.join(tmpdir, nombre)
        archivo.save(ruta_excel)

        # Extraer datos desde pestaña CALIBRACION
        datos = extraer_datos_calibracion(ruta_excel)
        
        # Detectar tipo de magnitud
        tipo = detectar_tipo_desde_datos(datos, nombre)

        env = os.environ.copy()
        env["LANG"]       = "es_PE.UTF-8"
        env["LC_ALL"]     = "es_PE.UTF-8"
        env["LC_NUMERIC"] = "es_PE.UTF-8"

        # Convertir hoja CERTIFICADO a PDF
        pdf_generado, error = convertir_excel_a_pdf(ruta_excel, tmpdir, env)
        
        if pdf_generado is None:
            return jsonify({"error": "Error generando PDF", "detalle": error}), 500

        # Nombrar el PDF final
        nombre_pdf = construir_nombre(datos, nombre)
        pdf_final  = os.path.join(tmpdir, nombre_pdf)

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
