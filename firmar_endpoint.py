import os
import io
from flask import Blueprint, request, jsonify, send_file
from pypdf import PdfReader, PdfWriter
from copy import deepcopy

firmar_bp = Blueprint('firmar', __name__)

PAGE_W   = 595.3
PAGE_H   = 841.9
TARGET_W = 180
FIRMA_X  = (PAGE_W - TARGET_W) / 2
FIRMA_Y  = 25 + (2 * 72 / 2.54)


def aplicar_membrete_y_firma(pdf_bytes, membrete_bytes, firma_bytes):
    cert_reader     = PdfReader(io.BytesIO(pdf_bytes),      strict=False)
    membrete_reader = PdfReader(io.BytesIO(membrete_bytes), strict=False)
    firma_reader    = PdfReader(io.BytesIO(firma_bytes),    strict=False)

    firma_orig_w = float(firma_reader.pages[0].mediabox.width)
    firma_orig_h = float(firma_reader.pages[0].mediabox.height)
    TARGET_H     = TARGET_W * (firma_orig_h / firma_orig_w)
    sx = TARGET_W / firma_orig_w
    sy = TARGET_H / firma_orig_h

    membrete_page = membrete_reader.pages[0]
    firma_page    = firma_reader.pages[0]
    writer        = PdfWriter()

    for i, page in enumerate(cert_reader.pages):
        nueva = deepcopy(page)
        nueva.merge_page(deepcopy(membrete_page), expand=False, over=False)
        if i == 0:
            firma = deepcopy(firma_page)
            firma.add_transformation([sx, 0, 0, sy, FIRMA_X, FIRMA_Y])
            firma.mediabox.lower_left  = (0, 0)
            firma.mediabox.upper_right = (PAGE_W, PAGE_H)
            nueva.merge_page(firma, over=True)
        writer.add_page(nueva)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


@firmar_bp.route('/firmar-pdf', methods=['POST'])
def firmar_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "No se envió archivo"}), 400

    archivo   = request.files['file']
    pdf_bytes = archivo.read()

    # Log de diagnóstico
    print(f"[firmar] Recibido: {archivo.filename}, {len(pdf_bytes)} bytes, header={pdf_bytes[:8]}", flush=True)

    if len(pdf_bytes) < 100:
        return jsonify({"error": f"PDF demasiado pequeño: {len(pdf_bytes)} bytes. Header: {pdf_bytes[:20]}"}), 400

    if not pdf_bytes.startswith(b'%PDF'):
        return jsonify({"error": f"No es un PDF válido. Header recibido: {pdf_bytes[:20]}"}), 400

    base_dir      = os.path.dirname(os.path.abspath(__file__))
    ruta_membrete = os.path.join(base_dir, "assets", "MEMBRETE_FINAL_MMC_2025.pdf")
    ruta_firma    = os.path.join(base_dir, "assets", "FIRMA_GABRIEL_2024-2025.pdf")

    if not os.path.exists(ruta_membrete):
        return jsonify({"error": "Membrete no encontrado"}), 500
    if not os.path.exists(ruta_firma):
        return jsonify({"error": "Firma no encontrada"}), 500

    # Verificar integridad de assets
    with open(ruta_membrete, "rb") as f:
        membrete_bytes = f.read()
    with open(ruta_firma, "rb") as f:
        firma_bytes = f.read()

    print(f"[firmar] Membrete: {len(membrete_bytes)} bytes, Firma: {len(firma_bytes)} bytes", flush=True)

    if not membrete_bytes.startswith(b'%PDF'):
        return jsonify({"error": f"Membrete corrupto en servidor. Header: {membrete_bytes[:20]}"}), 500
    if not firma_bytes.startswith(b'%PDF'):
        return jsonify({"error": f"Firma corrupta en servidor. Header: {firma_bytes[:20]}"}), 500

    try:
        resultado = aplicar_membrete_y_firma(pdf_bytes, membrete_bytes, firma_bytes)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    print(f"[firmar] PDF generado: {len(resultado)} bytes", flush=True)

    return send_file(
        io.BytesIO(resultado),
        mimetype='application/pdf',
        as_attachment=True,
        download_name=archivo.filename or "firmado.pdf"
    )
