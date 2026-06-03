import os
import io
import base64
import tempfile
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

    writer = PdfWriter()

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


@firmar_bp.route('/firmar-lote', methods=['POST'])
def firmar_lote():
    data = request.get_json(force=True, silent=True)
    if not data or "documentos" not in data:
        return jsonify({"error": "Body inválido. Se esperan 'documentos'"}), 400

    base_dir      = os.path.dirname(os.path.abspath(__file__))
    ruta_membrete = os.path.join(base_dir, "assets", "MEMBRETE_FINAL_MMC_2025.pdf")
    ruta_firma    = os.path.join(base_dir, "assets", "FIRMA_GABRIEL_2024-2025.pdf")

    if not os.path.exists(ruta_membrete):
        return jsonify({"error": "Membrete no encontrado en servidor"}), 500
    if not os.path.exists(ruta_firma):
        return jsonify({"error": "Firma no encontrada en servidor"}), 500

    with open(ruta_membrete, "rb") as f:
        membrete_bytes = f.read()
    with open(ruta_firma, "rb") as f:
        firma_bytes = f.read()

    procesados = []
    errores    = []

    for doc in data["documentos"]:
        nombre = doc.get("nombre", "documento.pdf")
        try:
            # Limpiar base64 por si viene con caracteres extra
            contenido_limpio = doc["contenido"].strip().replace("\n", "").replace("\r", "")
            # Agregar padding si falta
            padding = 4 - len(contenido_limpio) % 4
            if padding != 4:
                contenido_limpio += "=" * padding
            pdf_bytes = base64.b64decode(contenido_limpio)
            resultado = aplicar_membrete_y_firma(pdf_bytes, membrete_bytes, firma_bytes)
            procesados.append({
                "nombre":    nombre,
                "contenido": base64.b64encode(resultado).decode("utf-8")
            })
        except Exception as e:
            errores.append({"nombre": nombre, "error": str(e)})

    return jsonify({
        "procesados": procesados,
        "errores":    errores
    })
