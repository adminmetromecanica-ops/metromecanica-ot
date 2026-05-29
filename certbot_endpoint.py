import os
import re
import subprocess
import tempfile
import datetime
from flask import Blueprint, request, jsonify, send_file
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

certbot_bp = Blueprint('certbot', __name__)


def fmt_val(val, number_format=None):
    """
    Convierte un valor numérico al formato peruano (coma decimal)
    respetando el number_format de la celda.
    No toca strings, fechas, ni booleanos.
    """
    if val is None or isinstance(val, bool) or isinstance(val, str):
        return val

    if isinstance(val, datetime.datetime):
        return val.strftime("%Y-%m-%d")

    if isinstance(val, (int, float)):
        if not number_format or number_format in ("General", "@"):
            # Sin formato: limpiar ruido flotante pero no forzar decimales
            if isinstance(val, float) and val == int(val):
                return int(val)
            return val

        # Contar decimales del formato Excel (ej: "0.000" → 3)
        m = re.search(r'\.([0#]+)', number_format)
        decimales = len(m.group(1)) if m else 0

        redondeado = round(float(val), decimales)

        if decimales == 0:
            return int(redondeado)

        s = f"{redondeado:.{decimales}f}"
        return s.replace(".", ",")

    return val


def leer_certificado(ruta_excel):
    """
    Lee la hoja CERTIFICADO con data_only=True (valores calculados)
    Y con data_only=False (para leer number_format).
    Retorna dict: { coord: valor_formateado }
    """
    # Paso 1: leer valores calculados
    wb_vals = load_workbook(ruta_excel, read_only=False, data_only=True)
    cert_name = next(
        (s for s in wb_vals.sheetnames if s.upper() == "CERTIFICADO"),
        wb_vals.sheetnames[-1]
    )
    ws_vals = wb_vals[cert_name]
    valores = {}
    for row in ws_vals.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.value is not None:
                valores[cell.coordinate] = (cell.value, cell.number_format)
    wb_vals.close()

    # Paso 2: formatear cada valor
    resultado = {}
    for coord, (val, fmt) in valores.items():
        resultado[coord] = fmt_val(val, fmt)

    return resultado, cert_name


def preparar_para_pdf(ruta_excel, tmpdir):
    """
    1. Lee todos los valores calculados de CERTIFICADO (con su formato)
    2. Abre el workbook en modo edición
    3. Inyecta los valores estáticos (sin fórmulas) en CERTIFICADO
    4. Elimina todas las demás hojas
    5. Guarda
    """
    valores_cert, cert_name = leer_certificado(ruta_excel)

    wb = load_workbook(ruta_excel, data_only=False)
    ws = wb[cert_name]
    ws.sheet_state = "visible"

    # Inyectar valores estáticos
    for coord, val in valores_cert.items():
        try:
            cell = ws[coord]
            if isinstance(cell, MergedCell):
                # Escribir en la celda master del rango fusionado
                for rng in ws.merged_cells.ranges:
                    if coord in rng:
                        master = ws.cell(row=rng.min_row, column=rng.min_col)
                        master.value = val
                        break
            else:
                cell.value = val
        except Exception:
            pass

    # Limpiar cualquier fórmula residual
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            try:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = None
            except Exception:
                pass

    # Eliminar todas las demás hojas
    for nombre in [s for s in wb.sheetnames if s != cert_name]:
        try:
            wb[nombre].sheet_state = "visible"
            del wb[nombre]
        except Exception:
            pass

    ruta_out = os.path.join(tmpdir, "certificado_final.xlsx")
    wb.save(ruta_out)
    wb.close()
    return ruta_out


def construir_nombre(ruta_excel, nombre_archivo):
    """Construye el nombre del PDF desde pestaña CALIBRACION o nombre del archivo."""
    n_cert = magnitud = equipo = cliente = ot = ""
    try:
        wb = load_workbook(ruta_excel, read_only=True, data_only=True)
        if "CALIBRACION" in wb.sheetnames:
            cal = wb["CALIBRACION"]
            def g(coord):
                v = cal[coord].value
                return str(v).strip() if v else ""
            n_cert   = g("B150")
            magnitud = g("B151")
            equipo   = g("B152")
            cliente  = g("B153")
            ot       = g("B154")
        wb.close()
    except Exception:
        pass

    if not n_cert:
        for parte in nombre_archivo.upper().replace(".XLSM","").replace(".XLSX","").split("_"):
            if len(parte) > 5 and '-' in parte and any(
                parte.startswith(p) for p in ['MLL','MLF','MLE','MLP','MLT','MLM','MLQ','MLC','MLB']
            ):
                n_cert = parte
                break

    fecha  = datetime.date.today().strftime("%Y%m%d")
    nombre = f"{n_cert}_{magnitud}_{equipo}_{cliente}_{ot}_{fecha}.pdf"
    for c in ['\\','/',':', '*','?','"','<','>','|',' ','\n','\r']:
        nombre = nombre.replace(c, '_')
    while '__' in nombre:
        nombre = nombre.replace('__', '_')
    return nombre[:180]


@certbot_bp.route('/generar-certificado', methods=['POST'])
def generar_certificado():
    if 'file' not in request.files:
        return jsonify({"error": "No se envió archivo"}), 400

    archivo = request.files['file']
    nombre  = archivo.filename

    with tempfile.TemporaryDirectory() as tmpdir:
        ruta_excel = os.path.join(tmpdir, nombre)
        archivo.save(ruta_excel)

        try:
            ruta_xlsx = preparar_para_pdf(ruta_excel, tmpdir)
        except Exception as e:
            return jsonify({"error": f"Error preparando archivo: {str(e)}"}), 500

        env = os.environ.copy()
        env["LANG"]       = "es_PE.UTF-8"
        env["LC_ALL"]     = "es_PE.UTF-8"
        env["LC_NUMERIC"] = "es_PE.UTF-8"

        cmd = [
            "libreoffice", "--headless",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            ruta_xlsx
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=90, env=env)
        if result.returncode != 0:
            return jsonify({"error": "Error LibreOffice", "detalle": result.stderr}), 500

        pdf_generado = os.path.join(tmpdir, "certificado_final.pdf")
        if not os.path.exists(pdf_generado):
            return jsonify({"error": "PDF no generado"}), 500

        pdf_final = os.path.join(tmpdir, construir_nombre(ruta_excel, nombre))

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
