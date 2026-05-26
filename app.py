"""
app.py — Servidor web para generación de Órdenes de Trabajo
Metromecanica Ingenieria y Metrologia S.A.C.
"""

import os
import sys
import json
import subprocess
import tempfile
from flask import Flask, request, jsonify, send_file, send_from_directory

try:
    import audit_logger
    AUDIT_ENABLED = True
except ImportError:
    AUDIT_ENABLED = False

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
EXTRACTOR     = os.path.join(BASE_DIR, "extract_proforma.py")
GENERATOR     = os.path.join(BASE_DIR, "generate_ot.js")
OUTPUT_DIR    = os.path.join(BASE_DIR, "ordenes_generadas")
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024

from certbot_endpoint import certbot_bp
app.register_blueprint(certbot_bp)

HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Generador de OT · Metromecanica</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
  :root {
    --blue-dark:  #0f2d5e;
    --blue-mid:   #1a4e8c;
    --blue-light: #2e75b6;
    --accent:     #e8811a;
    --success:    #197a3e;
    --danger:     #b91c1c;
    --bg:         #f0f4f9;
    --card:       #ffffff;
    --border:     #d0daea;
    --text:       #1a2540;
    --muted:      #6b7a99;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'DM Sans', sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; }
  header { background: var(--blue-dark); padding: 0 2.5rem; display: flex; align-items: center; justify-content: space-between; height: 64px; box-shadow: 0 2px 12px rgba(0,0,0,.25); }
  .logo { display: flex; align-items: center; gap: .9rem; }
  .logo-img { height: 55px; width: auto; border-radius: 50%; }
  .logo-text { color: #fff; }
  .logo-text strong { display: block; font-family: 'DM Serif Display', serif; font-size: 1.05rem; letter-spacing: .02em; }
  .logo-text span { font-size: .72rem; color: #90afd4; font-weight: 300; }
  .iso-badge { background: rgba(255,255,255,.08); border: 1px solid rgba(255,255,255,.15); color: #90afd4; font-size: .7rem; font-weight: 500; padding: .3rem .75rem; border-radius: 20px; letter-spacing: .06em; }
  main { max-width: 680px; margin: 3rem auto; padding: 0 1.5rem 4rem; }
  h1 { font-family: 'DM Serif Display', serif; font-size: 1.9rem; color: var(--blue-dark); margin-bottom: .4rem; }
  .subtitle { color: var(--muted); font-size: .92rem; margin-bottom: 2.2rem; font-weight: 300; }
  .card { background: var(--card); border: 1px solid var(--border); border-radius: 16px; padding: 2rem; box-shadow: 0 2px 20px rgba(15,45,94,.06); margin-bottom: 1.5rem; }
  .card-title { font-size: .7rem; font-weight: 600; letter-spacing: .1em; color: var(--muted); text-transform: uppercase; margin-bottom: 1.2rem; }
  .dropzone { border: 2px dashed var(--border); border-radius: 12px; padding: 2.5rem 1.5rem; text-align: center; cursor: pointer; transition: all .2s; background: #fafbfd; position: relative; }
  .dropzone:hover, .dropzone.drag { border-color: var(--blue-light); background: #eef3fb; }
  .dropzone input[type=file] { position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%; }
  .dropzone-icon { font-size: 2.4rem; margin-bottom: .6rem; }
  .dropzone p { color: var(--muted); font-size: .88rem; }
  .dropzone p strong { color: var(--blue-light); }
  .file-selected { margin-top: 1rem; display: flex; align-items: center; gap: .6rem; background: #eef3fb; border-radius: 8px; padding: .6rem 1rem; font-size: .85rem; color: var(--blue-dark); }
  .estado-group { display: flex; gap: .8rem; margin-top: 1.4rem; }
  .estado-btn { flex: 1; padding: .75rem; border-radius: 10px; border: 2px solid var(--border); background: #fafbfd; cursor: pointer; font-family: 'DM Sans', sans-serif; font-size: .88rem; font-weight: 500; display: flex; align-items: center; justify-content: center; gap: .5rem; transition: all .15s; color: var(--muted); }
  .estado-btn:hover { border-color: var(--blue-light); color: var(--blue-dark); }
  .estado-btn.active-aprobada { border-color: var(--success); background: #f0faf5; color: var(--success); }
  .estado-btn.active-rechazada { border-color: var(--danger); background: #fef2f2; color: var(--danger); }
  .btn-generate { width: 100%; padding: .95rem; background: var(--blue-mid); color: #fff; border: none; border-radius: 12px; font-family: 'DM Sans', sans-serif; font-size: 1rem; font-weight: 600; cursor: pointer; margin-top: 1.6rem; display: flex; align-items: center; justify-content: center; gap: .6rem; transition: background .15s, transform .1s; }
  .btn-generate:hover:not(:disabled) { background: var(--blue-dark); transform: translateY(-1px); }
  .btn-generate:disabled { opacity: .5; cursor: not-allowed; }
  .status-box { border-radius: 12px; padding: 1.2rem 1.4rem; margin-top: 1.4rem; display: none; font-size: .9rem; line-height: 1.6; }
  .status-box.show { display: flex; align-items: flex-start; gap: .8rem; }
  .status-box.loading { background: #eef3fb; color: var(--blue-mid); border: 1px solid #c5d8f0; }
  .status-box.success { background: #f0faf5; color: var(--success); border: 1px solid #a7d9b8; }
  .status-box.error { background: #fef2f2; color: var(--danger); border: 1px solid #fca5a5; }
  .status-box.rejected { background: #fffbeb; color: #92400e; border: 1px solid #fcd34d; }
  .status-icon { font-size: 1.3rem; flex-shrink: 0; margin-top: .05rem; }
  .download-group { display: none; gap: .8rem; margin-top: .8rem; }
  .download-group.show { display: flex; }
  .btn-download { flex: 1; padding: .85rem; border: none; border-radius: 12px; font-family: 'DM Sans', sans-serif; font-size: .9rem; font-weight: 600; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: .5rem; text-decoration: none; transition: all .15s; }
  .btn-download.docx { background: var(--success); color: #fff; }
  .btn-download.docx:hover { background: #15653a; transform: translateY(-1px); }
  .btn-download.pdf { background: var(--danger); color: #fff; }
  .btn-download.pdf:hover { background: #991b1b; transform: translateY(-1px); }
  .preview-grid { display: none; grid-template-columns: 1fr 1fr; gap: .5rem .8rem; margin-top: .5rem; }
  .preview-grid.show { display: grid; }
  .pv-label { font-size: .72rem; color: var(--muted); font-weight: 500; text-transform: uppercase; letter-spacing: .06em; }
  .pv-value { font-size: .88rem; color: var(--text); font-weight: 500; margin-top: .1rem; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .spin { display: inline-block; animation: spin .8s linear infinite; }
  .history-item { display: flex; align-items: center; justify-content: space-between; padding: .65rem .8rem; border-radius: 8px; background: #fafbfd; border: 1px solid var(--border); margin-bottom: .5rem; font-size: .85rem; }
  .history-item .name { font-weight: 500; color: var(--blue-dark); }
  .history-item .meta { color: var(--muted); font-size: .78rem; }
  .history-item .downloads { display: flex; gap: .5rem; }
  .history-item a { color: var(--blue-light); text-decoration: none; font-weight: 500; font-size: .8rem; }
  .history-item a:hover { text-decoration: underline; }
  .history-empty { color: var(--muted); font-size: .85rem; text-align: center; padding: 1rem 0; }
</style>
</head>
<body>
<header>
  <div class="logo">
    <img src="/logo" alt="Metromecanica" class="logo-img">
    <div class="logo-text">
      <strong>Metromecanica</strong>
      <span>Sistema de Órdenes de Trabajo</span>
    </div>
  </div>
  <div class="iso-badge">ISO/IEC 17025:2017</div>
</header>
<main>
  <h1>Generar Orden de Trabajo</h1>
  <p class="subtitle">Sube la proforma aprobada desde tu sistema de facturación y genera la OT en segundos.</p>
  <div class="card">
    <div class="card-title">1 · Seleccionar proforma</div>
    <div class="dropzone" id="dropzone">
      <input type="file" id="fileInput" accept=".pdf">
      <div class="dropzone-icon">📄</div>
      <p><strong>Haz clic para subir</strong> o arrastra el PDF aquí</p>
      <p style="margin-top:.3rem">Solo archivos PDF · máx. 10 MB</p>
    </div>
    <div class="file-selected" id="fileSelected" style="display:none">
      <span class="icon">📎</span>
      <span id="fileName">—</span>
    </div>
    <div class="card-title" style="margin-top:1.6rem">2 · Estado de la proforma</div>
    <div class="estado-group">
      <button class="estado-btn active-aprobada" id="btnAprobada" onclick="setEstado('aprobada')">✅ Aprobada</button>
      <button class="estado-btn" id="btnRechazada" onclick="setEstado('rechazada')">❌ Rechazada</button>
    </div>
    <button class="btn-generate" id="btnGenerar" onclick="procesar()">⚙️ Procesar y generar OT</button>
    <div class="status-box loading" id="statusBox">
      <span class="status-icon spin" id="statusIcon">⏳</span>
      <div id="statusText">Procesando…</div>
    </div>
    <div class="download-group" id="downloadGroup">
      <a class="btn-download docx" id="btnDownloadDocx" href="#" download>📄 Descargar Word (.docx)</a>
      <a class="btn-download pdf" id="btnDownloadPdf" href="#" download>📕 Descargar PDF</a>
    </div>
  </div>
  <div class="card" id="previewCard" style="display:none">
    <div class="card-title">Datos extraídos de la proforma</div>
    <div class="preview-grid show" id="previewGrid"></div>
  </div>
  <div class="card">
    <div class="card-title">Órdenes generadas hoy</div>
    <div id="historial"><p class="history-empty">Ninguna OT generada aún.</p></div>
  </div>
</main>
<script>
let estadoSeleccionado = 'aprobada';
const historial = [];
function setEstado(e) {
  estadoSeleccionado = e;
  document.getElementById('btnAprobada').className  = 'estado-btn' + (e==='aprobada'  ? ' active-aprobada'  : '');
  document.getElementById('btnRechazada').className = 'estado-btn' + (e==='rechazada' ? ' active-rechazada' : '');
}
const dz = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
dz.addEventListener('dragover',  e => { e.preventDefault(); dz.classList.add('drag'); });
dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
dz.addEventListener('drop', e => {
  e.preventDefault(); dz.classList.remove('drag');
  const file = e.dataTransfer.files[0];
  if (file && file.type === 'application/pdf') {
    const dataTransfer = new DataTransfer();
    dataTransfer.items.add(file);
    fileInput.files = dataTransfer.files;
    setFile(file);
  }
});
fileInput.addEventListener('change', e => { if (e.target.files[0]) setFile(e.target.files[0]); });
function setFile(file) {
  document.getElementById('fileName').textContent = file.name;
  document.getElementById('fileSelected').style.display = 'flex';
  dz.style.display = 'none';
  resetStatus();
}
function resetStatus() {
  const sb = document.getElementById('statusBox');
  sb.className = 'status-box';
  document.getElementById('downloadGroup').className = 'download-group';
  document.getElementById('previewCard').style.display = 'none';
}
function showStatus(type, icon, text) {
  const sb = document.getElementById('statusBox');
  sb.className = `status-box show ${type}`;
  document.getElementById('statusIcon').textContent = icon;
  document.getElementById('statusIcon').className = 'status-icon' + (type==='loading' ? ' spin' : '');
  document.getElementById('statusText').innerHTML = text;
}
function showPreview(data) {
  const fields = [
    ['Proforma', data.numero_proforma],['Emisión', data.fecha_emision],
    ['Cliente', (data.cliente||'').split('-')[0].trim()],['Contacto', data.contacto_cliente],
    ['Equipos', (data.equipos||[]).join(', ')],['Plazo', data.plazo_entrega],
  ];
  const grid = document.getElementById('previewGrid');
  grid.innerHTML = fields.map(([l,v]) => `<div><div class="pv-label">${l}</div><div class="pv-value">${v||'—'}</div></div>`).join('');
  document.getElementById('previewCard').style.display = 'block';
}
function addHistorial(ot_num, filename) {
  const now = new Date().toLocaleTimeString('es-PE', {hour:'2-digit', minute:'2-digit'});
  historial.unshift({ ot_num, filename, time: now });
  const cont = document.getElementById('historial');
  cont.innerHTML = historial.slice(0,8).map(h =>
    `<div class="history-item"><div><div class="name">${h.ot_num}</div><div class="meta">${h.time}</div></div>
    <div class="downloads"><a href="/descargar/${h.filename}" download>Word</a>
    <a href="/descargar-pdf/${h.filename.replace('.docx','.pdf')}" download>PDF</a></div></div>`
  ).join('');
}
async function procesar() {
  const fileInput = document.getElementById('fileInput');
  if (!fileInput.files || fileInput.files.length === 0) { alert('Selecciona un archivo PDF primero.'); return; }
  document.getElementById('btnGenerar').disabled = true;
  resetStatus();
  showStatus('loading', '⏳', 'Leyendo y extrayendo datos de la proforma…');
  const fd = new FormData();
  fd.append('pdf', fileInput.files[0]);
  fd.append('estado', estadoSeleccionado);
  try {
    const res  = await fetch('/procesar', { method:'POST', body:fd });
    const data = await res.json();
    if (data.error) {
      showStatus('error', '❌', `<strong>Error:</strong> ${data.error}`);
    } else if (!data.aprobada) {
      showStatus('rejected', '⚠️', `<strong>Proforma ${data.numero_proforma} marcada como RECHAZADA.</strong><br>No se generó OT.`);
    } else {
      showStatus('success', '✅', `<strong>OT generada: ${data.ot_num}</strong><br>Cliente: ${(data.cliente||'').split('-')[0].trim()}`);
      showPreview(data);
      const dlDocx = document.getElementById('btnDownloadDocx');
      const dlPdf  = document.getElementById('btnDownloadPdf');
      dlDocx.href = `/descargar/${data.filename}`;
      dlPdf.href  = `/descargar-pdf/${data.filename.replace('.docx','.pdf')}`;
      document.getElementById('downloadGroup').className = 'download-group show';
      addHistorial(data.ot_num, data.filename);
    }
  } catch(e) {
    showStatus('error', '❌', `Error de conexión: ${e.message}`);
  } finally {
    document.getElementById('btnGenerar').disabled = false;
  }
}
</script>
</body>
</html>"""


@app.route('/')
def index():
    return HTML

@app.route('/procesar', methods=['POST'])
def procesar():
    pdf_file = request.files.get('pdf')
    estado   = request.form.get('estado', 'aprobada')
    if not pdf_file or not pdf_file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Debes subir un archivo PDF válido.'}), 400
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
        pdf_file.save(tmp.name)
        tmp_pdf = tmp.name
    try:
        result = subprocess.run([sys.executable, EXTRACTOR, tmp_pdf, '--estado', estado], capture_output=True, text=True)
        if result.returncode != 0:
            return jsonify({'error': f'Error al leer el PDF: {result.stderr}'}), 500
        data = json.loads(result.stdout)
        if not data.get('aprobada'):
            return jsonify({'aprobada': False, 'numero_proforma': data.get('numero_proforma', '')})
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as jf:
            json.dump(data, jf, ensure_ascii=False)
            tmp_json = jf.name
        try:
            gen_result = subprocess.run(['node', GENERATOR, '--file', tmp_json], capture_output=True, text=True, cwd=OUTPUT_DIR)
            if gen_result.returncode != 0:
                return jsonify({'error': f'Error al generar OT: {gen_result.stderr}'}), 500
            ot_path     = gen_result.stdout.strip().replace('OK:', '').strip()
            ot_filename = os.path.basename(ot_path)
            ot_num      = ot_filename.replace('.docx', '')
            if AUDIT_ENABLED:
                audit_logger.register_ot({'ot_number': ot_num, 'expediente': data.get('expediente',''), 'numero_proforma': data.get('numero_proforma',''), 'cliente': data.get('cliente',''), 'ruc_cliente': data.get('ruc_cliente',''), 'total_items': data.get('total_items',0), 'tipo_servicio': data.get('tipo_servicio','GENERAL'), 'fecha_emision': data.get('fecha_emision',''), 'plazo_entrega': data.get('plazo_entrega','')}, ot_path)
            return jsonify({'aprobada': True, 'ot_num': ot_num, 'filename': ot_filename, 'cliente': data.get('cliente',''), 'equipos': data.get('equipos',[]), 'numero_proforma': data.get('numero_proforma',''), 'fecha_emision': data.get('fecha_emision',''), 'contacto_cliente': data.get('contacto_cliente',''), 'plazo_entrega': data.get('plazo_entrega','')})
        finally:
            os.unlink(tmp_json)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        os.unlink(tmp_pdf)

@app.route('/logo')
def serve_logo():
    logo_path = os.path.join(BASE_DIR, 'logo_metromecanica.png')
    if os.path.exists(logo_path):
        return send_file(logo_path, mimetype='image/png')
    return '', 404

@app.route('/descargar/<filename>')
def descargar(filename):
    safe_name = os.path.basename(filename)
    if not safe_name.endswith('.docx'):
        return 'Archivo no válido', 400
    filepath = os.path.join(OUTPUT_DIR, safe_name)
    if not os.path.exists(filepath):
        return 'Archivo no encontrado', 404
    return send_file(filepath, as_attachment=True, download_name=safe_name)

@app.route('/descargar-pdf/<filename>')
def descargar_pdf(filename):
    safe_name = os.path.basename(filename)
    if not safe_name.endswith('.pdf'):
        return 'Archivo no válido', 400
    docx_name = safe_name.replace('.pdf', '.docx')
    docx_path = os.path.join(OUTPUT_DIR, docx_name)
    if not os.path.exists(docx_path):
        return 'Archivo Word no encontrado', 404
    pdf_path = os.path.join(OUTPUT_DIR, safe_name)
    if os.path.exists(pdf_path):
        return send_file(pdf_path, as_attachment=True, download_name=safe_name)
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', OUTPUT_DIR, docx_path], check=True, capture_output=True, timeout=30)
        if os.path.exists(pdf_path):
            return send_file(pdf_path, as_attachment=True, download_name=safe_name)
        return 'Error al generar PDF', 500
    except Exception as e:
        return f'Error: {str(e)}', 500

@app.route('/auditoria/exportar')
def exportar_auditoria():
    if not AUDIT_ENABLED:
        return jsonify({'error': 'Sistema de auditoría no disponible'}), 503
    import tempfile
    from datetime import datetime, timedelta
    csv_file = tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False)
    csv_path = csv_file.name
    csv_file.close()
    end_date   = datetime.now().strftime('%Y-%m-%d')
    start_date = (datetime.now() - timedelta(days=365)).strftime('%Y-%m-%d')
    if audit_logger.export_audit_csv(csv_path, start_date, end_date):
        return send_file(csv_path, mimetype='text/csv', as_attachment=True, download_name=f'auditoria_inacal_{end_date}.csv')
    return jsonify({'error': 'No hay registros'}), 404

@app.route('/auditoria/estadisticas')
def estadisticas_auditoria():
    if not AUDIT_ENABLED:
        return jsonify({'error': 'Sistema de auditoría no disponible'}), 503
    return jsonify(audit_logger.get_statistics())

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"\n{'='*55}\n  METROMECANICA · Sistema de OT\n  Puerto: {port}\n{'='*55}\n")
    app.run(host='0.0.0.0', port=port, debug=False)
