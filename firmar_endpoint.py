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
    # Recibe: file, membrete, firma — todos como multipart
    if 'file' not in request.files:
        return jsonify({"error": "Falta campo 'file'"}), 400
    if 'membrete' not in request.files:
        return jsonify({"error": "Falta campo 'membrete'"}), 400
    if 'firma' not in request.files:
        return jsonify({"error": "Falta campo 'firma'"}), 400

    pdf_bytes      = request.files['file'].read()
    membrete_bytes = request.files['membrete'].read()
    firma_bytes    = request.files['firma'].read()

    if not pdf_bytes.startswith(b'%PDF'):
        return jsonify({"error": f"PDF inválido. Header: {pdf_bytes[:20]}"}), 400
    if not membrete_bytes.startswith(b'%PDF'):
        return jsonify({"error": f"Membrete inválido. Header: {membrete_bytes[:20]}"}), 400
    if not firma_bytes.startswith(b'%PDF'):
        return jsonify({"error": f"Firma inválida. Header: {firma_bytes[:20]}"}), 400

    try:
        resultado = aplicar_membrete_y_firma(pdf_bytes, membrete_bytes, firma_bytes)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    return send_file(
        io.BytesIO(resultado),
        mimetype='application/pdf',
        as_attachment=True,
        download_name=request.files['file'].filename or "firmado.pdf"
    )
