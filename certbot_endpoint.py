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

        ot      = leer_celda(ws, "D5") or leer_celda(ws, "J5")
        solic   = leer_celda(ws, "B9") or leer_celda(ws, "J9")
        instrum = leer_celda(ws, "B15") or leer_celda(ws, "J12")

        wb.close()
        return {"n_certificado": n_cert, "orden_trabajo": ot, "solicitante": solic, "instrumento": instrum}
    except Exception:
        return {"n_certificado": "CERT", "orden_trabajo": "", "solicitante": "", "instrumento": ""}

def construir_nombre(datos, tipo, nombre_archivo=""):
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
    """Formatea número con coma decimal (estilo peruano/europeo)"""
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

def resolver_formulas_e_inyectar(ruta_excel, tmpdir):
    wb_vals = load_workbook(ruta_excel, data_only=True)

    reg = None
    for h in HOJAS_DATOS:
        if h in wb_vals.sheetnames:
            reg = wb_vals[h]
            break
    if reg is None:
        reg = wb_vals.worksheets[0]

    med       = wb_vals["MEDICION"]     if "MEDICION"     in wb_vals.sheetnames else None
    cert_vals = wb_vals["CERTIFICADO"]  if "CERTIFICADO"  in wb_vals.sheetnames else None

    def v(ws, coord):
        if ws is None: return ""
        val = ws[coord].value
        if val is None: return ""
        if isinstance(val, datetime.datetime):
            return val.strftime("%Y-%m-%d")
        s = str(val).strip()
        return "" if s.lower() in ["none","nan"] else s

    r = {
        "cert":      v(reg, "B8"),
        "ot":        v(reg, "D5"),
        "fecha_e":   v(reg, "D11"),
        "fecha_c":   v(reg, "D12"),
        "cliente":   v(reg, "B9"),
        "dir":       v(reg, "B10"),
        "instrum":   v(reg, "B15"),
        "marca":     v(reg, "D15"),
        "modelo":    v(reg, "B16"),
        "serie":     v(reg, "D16") or v(reg, "B5"),
        "ident":     v(reg, "B17"),
        "rmin":      v(reg, "B18"),
        "rmax":      v(reg, "D18"),
        "resol":     v(reg, "B19"),
        "unidad":    v(reg, "D19"),
        "tipo_eq":   v(reg, "B20"),
        "ubicacion": v(reg, "D20"),
        "ti":        v(reg, "B23"),
        "tf":        v(reg, "D23"),
        "hi":        v(reg, "B24"),
        "hf":        v(reg, "D24"),
        "obs":       v(reg, "B27"),
    }

    # Mediciones filas 6-15 con coma decimal
    med_data = []
    if med:
        for row_idx in range(6, 16):
            fila = [v(med, f"{col}{row_idx}") for col in ["B","C","D","E","F"]]
            med_data.append(fila)

    # Trazabilidad filas 103-105
    traz_data = []
    if med:
        for row_idx in range(103, 106):
            fila = [v(med, f"{col}{row_idx}") for col in ["D","E","F"]]
            traz_data.append(fila)

    # Valores calculados internos del CERTIFICADO
    cert_internos = {}
    if cert_vals:
        celdas_int = ["A94","B94","A97","B97","A100","B100","A103","B103",
                      "A121","B121","A124","B124","A127","B127"]
        for c in celdas_int:
            cert_internos[c] = v(cert_vals, c)

    wb_vals.close()

    # Abrir para editar
    wb_edit = load_workbook(ruta_excel)
    wc = wb_edit["CERTIFICADO"] if "CERTIFICADO" in wb_edit.sheetnames else wb_edit.worksheets[-1]
    wc.sheet_state = "visible"

    # Datos administrativos e instrumento
    escribir_celda(wc, "A7",  r["cert"])
    escribir_celda(wc, "B9",  r["ot"])
    escribir_celda(wc, "D9",  r["fecha_e"])
    escribir_celda(wc, "B12", r["cliente"])
    escribir_celda(wc, "B13", r["dir"])
    escribir_celda(wc, "B16", r["instrum"])
    escribir_celda(wc, "B17", r["marca"])
    escribir_celda(wc, "B18", r["modelo"])
    escribir_celda(wc, "B19", r["serie"])
    escribir_celda(wc, "B20", r["ident"])
    escribir_celda(wc, "B21", f"{r['rmin']} {r['unidad']} a {r['rmax']} {r['unidad']}")
    escribir_celda(wc, "B22", f"{r['resol']} {r['unidad']}")
    escribir_celda(wc, "B23", r["tipo_eq"])
    escribir_celda(wc, "A24", r["ubicacion"])
    escribir_celda(wc, "B26", r["fecha_c"])
    escribir_celda(wc, "B31", f"{r['ti']} °C A {r['tf']} °C")
    escribir_celda(wc, "B32", f"{r['hi']} %HR A {r['hf']} %HR")
    escribir_celda(wc, "A145", r["obs"])

    # Mediciones con coma decimal
    for i, fila in enumerate(med_data):
        row = 81 + i
        escribir_celda(wc, f"A{row}", formatear_decimal(fila[0], 3))
        escribir_celda(wc, f"B{row}", formatear_decimal(fila[1], 3))
        escribir_celda(wc, f"C{row}", formatear_decimal(fila[2], 1))
        escribir_celda(wc, f"D{row}", formatear_decimal(fila[3], 4))

    # Trazabilidad
    for i, fila in enumerate(traz_data):
        row = 71 + i
        escribir_celda(wc, f"A{row}", fila[0])
        escribir_celda(wc, f"B{row}", fila[1])
        escribir_celda(wc, f"C{row}", fila[2])

    # Valores calculados internos con coma decimal
    for coord, val in cert_internos.items():
        escribir_celda(wc, coord, formatear_decimal(val, 2))

    # Eliminar otras hojas
    hojas_borrar = [s for s in wb_edit.sheetnames if s != "CERTIFICADO"]
    for h in hojas_borrar:
        wb_edit[h].sheet_state = "visible"
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
