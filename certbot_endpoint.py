import os
import subprocess
import tempfile
import datetime
from flask import Blueprint, request, jsonify, send_file
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

certbot_bp = Blueprint('certbot', __name__)

def formatear_numero(val, decimales=None):
    """
    Convierte números al estilo peruano (coma decimal), con redondeo inteligente.
    - decimales=None  → detecta automáticamente los decimales significativos (máx 6)
    - decimales=N     → fuerza N decimales fijos
    """
    if val is None:
        return val

    if isinstance(val, bool):
        return val  # booleanos primero, antes de int

    if isinstance(val, int):
        return val

    if isinstance(val, float):
        # Redondear para eliminar ruido de punto flotante
        # Determinar cuántos decimales reales tiene el valor
        if decimales is not None:
            redondeado = round(val, decimales)
        else:
            # Detectar decimales significativos: probar desde 6 hacia abajo
            redondeado = None
            for d in range(6, -1, -1):
                r = round(val, d)
                if abs(r - val) < 1e-9:
                    redondeado = r
                    decimales_uso = d
                    break
            if redondeado is None:
                redondeado = round(val, 6)
                decimales_uso = 6
            decimales = decimales_uso

        # Si el valor redondeado es entero exacto, devolver como entero
        if decimales == 0 or redondeado == int(redondeado):
            return int(redondeado)

        # Formatear con los decimales correctos y reemplazar punto por coma
        s = f"{redondeado:.{decimales}f}"
        return s.replace(".", ",")

    if isinstance(val, str):
        try:
            f = float(val.replace(",", "."))
            return formatear_numero(f)
        except Exception:
            return val

    return val

# ---------------------------------------------------------------------------
# Mapa de precisión fija para celdas de resultados del certificado
# Se construye dinámicamente por _construir_mapa_precision()
# ---------------------------------------------------------------------------
PRECISION_CELDAS = {}

def _construir_mapa_precision(hojas_data, cert_sheet):
    """
    Detecta automáticamente las columnas de la tabla de resultados
    buscando los encabezados en CERTIFICADO y aplica precisión fija.
      Indicación (mm)  → 3 decimales
      VCV (mm)         → 3 decimales
      Error (mm)       → 3 decimales
      Error (µm)       → 2 decimales
      fmax / U         → 2 decimales
    """
    global PRECISION_CELDAS
    PRECISION_CELDAS = {}

    datos_cert = hojas_data.get(cert_sheet, {})
    col_indicacion = col_vcv = col_error_mm = col_error_um = None
    fila_inicio = None

    for coord, val in datos_cert.items():
        if not isinstance(val, str):
            continue
        vl = val.lower()
        col = ''.join(c for c in coord if c.isalpha())
        try:
            fila = int(''.join(c for c in coord if c.isdigit()))
        except ValueError:
            continue

        if "indicaci" in vl and "mm" in vl and col_indicacion is None:
            col_indicacion = col
            fila_inicio = fila + 1
        elif ("convencionalmente" in vl or "vcv" in vl or "verdadero" in vl) and "mm" in vl and col_vcv is None:
            col_vcv = col
        elif "error" in vl and "mm" in vl and col_error_mm is None:
            col_error_mm = col
        elif "error" in vl and ("\u00b5m" in vl or "um" in vl) and col_error_um is None:
            col_error_um = col

    if fila_inicio:
        for fila in range(fila_inicio, fila_inicio + 12):
            if col_indicacion:
                PRECISION_CELDAS[f"{col_indicacion}{fila}"] = 3
            if col_vcv:
                PRECISION_CELDAS[f"{col_vcv}{fila}"] = 3
            if col_error_mm:
                PRECISION_CELDAS[f"{col_error_mm}{fila}"] = 3
            if col_error_um:
                PRECISION_CELDAS[f"{col_error_um}{fila}"] = 2

    # fmax y U: celdas numéricas en filas que mencionan esas palabras clave
    for coord, val in datos_cert.items():
        if isinstance(val, str) and ("fmax" in val.lower() or "incertidumbre" in val.lower()):
            col = ''.join(c for c in coord if c.isalpha())
            try:
                fila = int(''.join(c for c in coord if c.isdigit()))
            except ValueError:
                continue
            # El valor numérico suele estar algunas columnas a la derecha
            for offset in range(1, 10):
                col_val = chr(ord(col[-1]) + offset) if len(col) == 1 else None
                if col_val:
                    cand = f"{col_val}{fila}"
                    if cand in datos_cert and isinstance(datos_cert[cand], (int, float)):
                        PRECISION_CELDAS[cand] = 2


def formatear_numero_celda(coord, val):
    """Formatea un valor usando la precisión mapeada para esa coordenada."""
    if coord in PRECISION_CELDAS:
        return formatear_numero(val, decimales=PRECISION_CELDAS[coord])
    return formatear_numero(val)


def leer_todos_valores(ruta_excel):
    """Lee valores calculados de TODAS las hojas del Excel."""
    wb = load_workbook(ruta_excel, read_only=True, data_only=True)
    hojas_data = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        hojas_data[sheet_name] = {}
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                try:
                    if cell.value is not None and hasattr(cell, 'coordinate'):
                        val = cell.value
                        if isinstance(val, datetime.datetime):
                            val = val.strftime("%Y-%m-%d")
                        hojas_data[sheet_name][cell.coordinate] = val
                except Exception:
                    pass
    wb.close()
    return hojas_data

def extraer_datos_nombre(hojas_data, nombre_archivo):
    """Extrae datos para nombrar el PDF desde pestaña CALIBRACION."""
    cal = hojas_data.get("CALIBRACION", {})

    n_cert   = str(cal.get("B150", "") or "").strip()
    magnitud = str(cal.get("B151", "") or "").strip()
    equipo   = str(cal.get("B152", "") or "").strip()
    cliente  = str(cal.get("B153", "") or "").strip()
    ot       = str(cal.get("B154", "") or "").strip()

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
    Estrategia universal:
    1. Lee valores calculados de TODAS las hojas
    2. Construye mapa de precision segun encabezados de la tabla de resultados
    3. Inyecta esos valores en CERTIFICADO con coma decimal y precision correcta
    4. Elimina otras hojas
    5. Guarda Excel listo para PDF
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

    # Construir mapa de precision ANTES de inyectar
    _construir_mapa_precision(hojas_data, cert_sheet)

    wc = wb_edit[cert_sheet]
    wc.sheet_state = "visible"

    # Inyectar valores calculados con coma decimal y precision por celda
    cert_vals = hojas_data.get(cert_sheet, {})
    for coord, val in cert_vals.items():
        if val is not None:
            escribir_celda(wc, coord, formatear_numero_celda(coord, val))

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
    hojas_borrar = [s for s in wb_edit.sheetnames if s != cert_sheet]
    for h in hojas_borrar:
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
