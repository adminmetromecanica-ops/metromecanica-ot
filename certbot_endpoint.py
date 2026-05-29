import os
import re
import subprocess
import tempfile
import datetime
from flask import Blueprint, request, jsonify, send_file
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

certbot_bp = Blueprint('certbot', __name__)


def decimales_desde_formato(number_format):
    """
    Extrae la cantidad de decimales del number_format de Excel.
    Ej: '0.000' → 3,  '0.00' → 2,  '0.0' → 1,  '0' → 0
    """
    if not number_format or number_format in ("General", "@"):
        return None
    # Buscar la parte decimal del formato (después del punto)
    m = re.search(r'\.([0#]+)', number_format)
    if m:
        return len(m.group(1))
    return 0


def formatear_valor(val, number_format=None):
    """
    Formatea un valor numérico con coma decimal respetando el
    number_format de la celda Excel.
    - Si tiene formato → usa exactamente esos decimales
    - Si no tiene formato → devuelve el valor tal cual (int o float limpio)
    """
    if val is None:
        return val
    if isinstance(val, bool):
        return val
    if isinstance(val, str):
        return val
    if isinstance(val, datetime.datetime):
        return val.strftime("%Y-%m-%d")

    if isinstance(val, (int, float)):
        decimales = decimales_desde_formato(number_format)

        if decimales is None:
            # Sin formato conocido: devolver limpio
            if isinstance(val, float) and val == int(val):
                return int(val)
            return val

        redondeado = round(float(val), decimales)

        if decimales == 0:
            return int(redondeado)

        s = f"{redondeado:.{decimales}f}"
        return s.replace(".", ",")

    return val


def leer_todos_valores(ruta_excel):
    """
    Lee valores calculados Y number_format de TODAS las hojas del Excel.
    Retorna: { sheet_name: { coord: (value, number_format) } }
    """
    wb = load_workbook(ruta_excel, read_only=False, data_only=True)
    hojas_data = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        hojas_data[sheet_name] = {}
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                try:
                    if cell.value is not None:
                        val = cell.value
                        fmt = cell.number_format
                        hojas_data[sheet_name][cell.coordinate] = (val, fmt)
                except Exception:
                    pass
    wb.close()
    return hojas_data


def extraer_datos_nombre(hojas_data, nombre_archivo):
    """Extrae datos para nombrar el PDF desde pestaña CALIBRACION."""
    cal = hojas_data.get("CALIBRACION", {})

    def get(coord):
        entry = cal.get(coord)
        return str(entry[0] if entry else "").strip()

    n_cert   = get("B150")
    magnitud = get("B151")
    equipo   = get("B152")
    cliente  = get("B153")
    ot       = get("B154")

    if not n_cert:
        partes = nombre_archivo.upper().replace(".XLSM","").replace(".XLSX","").split("_")
        for parte in partes:
            if len(parte) > 5 and '-' in parte and any(
                parte.startswith(p) for p in ['MLL','MLF','MLE','MLP','MLT','MLM','MLQ','MLC','MLB']
            ):
                n_cert = parte
                break

    return {
        "n_certificado": n_cert,
        "magnitud":      magnitud,
        "instrumento":   equipo,
        "solicitante":   cliente,
        "orden_trabajo": ot,
    }


def construir_nombre(datos, nombre_archivo=""):
    cert     = datos.get("n_certificado", "CERT") or "CERT"
    magnitud = datos.get("magnitud", "") or ""
    inst     = datos.get("instrumento", "") or ""
    solic    = datos.get("solicitante", "") or ""
    ot       = datos.get("orden_trabajo", "") or ""
    fecha    = datetime.date.today().strftime("%Y%m%d")

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


def preparar_certificado(ruta_excel, tmpdir):
    """
    1. Lee valores calculados + number_format de TODAS las hojas
    2. Inyecta en CERTIFICADO los valores formateados con coma decimal
       respetando exactamente el formato numérico de cada celda
    3. Elimina otras hojas y guarda
    """
    hojas_data = leer_todos_valores(ruta_excel)

    wb_edit = load_workbook(ruta_excel, data_only=False)

    # Encontrar hoja CERTIFICADO
    cert_sheet = None
    for nombre in wb_edit.sheetnames:
        if nombre.upper() == "CERTIFICADO":
            cert_sheet = nombre
            break
    if cert_sheet is None:
        cert_sheet = wb_edit.sheetnames[-1]

    wc = wb_edit[cert_sheet]
    wc.sheet_state = "visible"

    # Inyectar valores con formato correcto
    cert_vals = hojas_data.get(cert_sheet, {})
    for coord, (val, fmt) in cert_vals.items():
        valor_formateado = formatear_valor(val, fmt)
        escribir_celda(wc, coord, valor_formateado)

    # Limpiar fórmulas restantes
    for row in wc.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            try:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = None
            except Exception:
                pass

    # Eliminar otras hojas
    for h in [s for s in wb_edit.sheetnames if s != cert_sheet]:
        try:
            wb_edit[h].sheet_state = "visible"
            del wb_edit[h]
        except Exception:
            pass

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

    with tempfile.TemporaryDirectory() as tmpdir:
        ruta_excel = os.path.join(tmpdir, nombre)
        archivo.save(ruta_excel)

        hojas_data = leer_todos_valores(ruta_excel)
        datos = extraer_datos_nombre(hojas_data, nombre)

        env = os.environ.copy()
        env["LANG"]       = "es_PE.UTF-8"
        env["LC_ALL"]     = "es_PE.UTF-8"
        env["LC_NUMERIC"] = "es_PE.UTF-8"

        try:
            ruta_cert_xlsx = preparar_certificado(ruta_excel, tmpdir)
        except Exception as e:
            return jsonify({"error": f"Error preparando certificado: {str(e)}"}), 500

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

        pdf_final = os.path.join(tmpdir, construir_nombre(datos, nombre))

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
