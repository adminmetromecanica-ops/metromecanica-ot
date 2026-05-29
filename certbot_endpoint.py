import os
import subprocess
import tempfile
import datetime
from flask import Blueprint, request, jsonify, send_file
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

certbot_bp = Blueprint('certbot', __name__)


def construir_nombre(ruta_excel, nombre_archivo):
    """Construye el nombre del PDF desde pestaña CALIBRACION o desde el nombre del archivo."""
    try:
        wb = load_workbook(ruta_excel, read_only=True, data_only=True)
        cal = wb["CALIBRACION"] if "CALIBRACION" in wb.sheetnames else None
        n_cert = magnitud = equipo = cliente = ot = ""
        if cal:
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
        n_cert = magnitud = equipo = cliente = ot = ""

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
    return nombre[:180] if len(nombre) > 180 else nombre


def preparar_para_pdf(ruta_excel, tmpdir):
    """
    Deja solo la hoja CERTIFICADO visible y guarda.
    No toca ningún valor — LibreOffice imprime tal cual.
    """
    wb = load_workbook(ruta_excel, data_only=False)

    # Identificar hoja CERTIFICADO
    cert_sheet = next(
        (s for s in wb.sheetnames if s.upper() == "CERTIFICADO"),
        wb.sheetnames[-1]
    )

    # Ocultar todas las demás hojas
    for nombre in wb.sheetnames:
        if nombre == cert_sheet:
            wb[nombre].sheet_state = "visible"
        else:
            try:
                wb[nombre].sheet_state = "hidden"
            except Exception:
                pass

    ruta_out = os.path.join(tmpdir, "certificado_final.xlsx")
    wb.save(ruta_out)
    wb.close()
    return ruta_out


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
