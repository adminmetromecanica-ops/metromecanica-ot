import os
import subprocess
import tempfile
import datetime
from flask import Blueprint, request, jsonify, send_file
from openpyxl import load_workbook

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
        if isinstance(val, datetime.datetime):
            return val.strftime("%Y-%m-%d")
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
        return {"n_certificado": n_cert, "orden_trabajo": ot, "solicitante": solic, "instrumento": instrum}
    except Exception:
        return {"n_certificado": "CERT", "orden_trabajo": "", "solicitante": "", "instrumento": ""}

def construir_nombre(datos, tipo, nombre_archivo=""):
    # Leer cert del nombre del archivo primero
    cert = ""
    if nombre_archivo:
        partes = nombre_archivo.upper().replace(".XLSM","").replace(".XLSX","").split("_")
        for parte in partes:
            if len(parte) > 5 and '-' in parte and any(
                parte.startswith(p) for p in ['MLL','MLF','MLE','MLP','MLT','MLM','MLQ','MLC','MLB']
            ):
                cert = parte
                break

    if not cert:
        cert = datos.get("n_certificado", "CERT")

    inst  = datos.get("instrumento", "")
    solic = datos.get("solicitante", "")
    ot    = datos.get("orden_trabajo", "")
    fecha = datetime.date.today().strftime("%Y%m%d")

    if not cert or cert.lower() in ["none","nan",""]: cert = "CERT"
    if not inst or inst.lower() in ["none","nan"]: inst = ""
    if not solic or solic.lower() in ["none","nan"]: solic = ""
    if not ot or ot.lower() in ["none","nan"]: ot = ""

    nombre = f"{cert}_{inst}_{solic}_{ot}_{fecha}.pdf"
    for c in ['\\','/',':', '*','?','"','<','>','|',' ','\n','\r']:
        nombre = nombre.replace(c, '_')
    while '__' in nombre:
        nombre = nombre.replace('__', '_')
    if len(nombre) > 150:
        nombre = nombre[:146] + ".pdf"
    return nombre

def resolver_formulas_e_inyectar(ruta_excel, tmpdir):
    wb_vals = load_workbook(ruta_excel, data_only=True)

    reg = None
    for h in HOJAS_DATOS:
        if h in wb_vals.sheetnames:
            reg = wb_vals[h]
            break
    if reg is None:
        reg = wb_vals.worksheets[0]

    med = wb_vals["MEDICION"] if "MEDICION" in wb_vals.sheetnames else None

    def v(ws, coord):
        if ws is None: return ""
        val = ws[coord].value
        if val is None: return ""
        if isinstance(val, datetime.datetime):
            return val.strftime("%Y-%m-%d")
        s = str(val).strip()
        return "" if s.lower() in ["none","nan"] else s

    r = {
        "cert":    v(reg, "B8") or v(reg, "B5"),
        "ot":      v(reg, "D5"),
        "fecha_e": v(reg, "D11"),
        "cliente": v(reg, "B9"),
        "dir":     v(reg, "B10"),
        "fecha_c": v(reg, "B11"),
        "instrum": v(reg, "B15"),
        "marca":   v(reg, "D15"),
        "modelo":  v(reg, "B16"),
        "serie":   v(reg, "D16"),
        "ident":   v(reg, "B17"),
        "rmin":    v(reg, "B18"),
        "rmax":    v(reg, "D18"),
        "resol":   v(reg, "B19"),
        "unidad":  v(reg, "D19"),
        "ti":      v(reg, "B23"),
        "tf":      v(reg, "D23"),
        "hi":      v(reg, "B24"),
        "hf":      v(reg, "D24"),
        "obs":     v(reg, "B27"),
    }

    med_ext = []
    if med:
        for row_idx in range(5, 15):
            fila = [v(med, f"{col}{row_idx}") for col in ["B","C","D","E","F"]]
            med_ext.append(fila)

    wb_vals.close()

    wb_edit = load_workbook(ruta_excel)
    wc = wb_edit["CERTIFICADO"] if "CERTIFICADO" in wb_edit.sheetnames else wb_edit.worksheets[-1]

    wc["A2"] = r["cert"]
    wc["B4"] = r["ot"]
    wc["D4"] = r["fecha_e"]
    wc["B7"] = r["cliente"]
    wc["B8"] = r["dir"]
    wc["B11"] = r["instrum"]
    wc["B12"] = r["marca"]
    wc["B13"] = r["modelo"]
    wc["B14"] = r["serie"]
    wc["B15"] = r["ident"]
    wc["B16"] = f"{r['rmin']} {r['unidad']} a {r['rmax']} {r['unidad']}"
    wc["B17"] = f"{r['resol']} {r['unidad']}"
    wc["B19"] = r["fecha_c"]
    wc["B23"] = f"{r['ti']} °C A {r['tf']} °C"
    wc["D23"] = f"{r['hi']} %HR A {r['hf']} %HR"

    for i, fila in enumerate(med_ext):
        row = 28 + i
        wc[f"A{row}"] = fila[0]
        wc[f"B{row}"] = fila[1]
        wc[f"C{row}"] = fila[2]
        wc[f"D{row}"] = fila[3]

    hojas_borrar = [s for s in wb_edit.sheetnames if s != "CERTIFICADO"]
    for h in hojas_borrar:
        del wb_edit[h]

    ruta_cert = os.path.join(tmpdir, "certificado_final.xlsx")
    wb_edit.save(ruta_cert)
    wb_edit.close()

    return ruta_cert

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

        try:
            ruta_cert_xlsx = resolver_formulas_e_inyectar(ruta_excel, tmpdir)
        except Exception as e:
            return jsonify({"error": f"Error procesando Excel: {str(e)}"}), 500

        cmd = [
            "libreoffice", "--headless",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            ruta_cert_xlsx
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=90, env=env)
        if result.returncode != 0:
            return jsonify({"error": "Error LibreOffice", "detalle": result.stderr}), 500

        pdf_generado = os.path.join(tmpdir, "certificado_final.pdf")
        if not os.path.exists(pdf_generado):
            return jsonify({"error": "PDF no generado"}), 500

        pdf_final = os.path.join(tmpdir, construir_nombre(datos, tipo, nombre))

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
