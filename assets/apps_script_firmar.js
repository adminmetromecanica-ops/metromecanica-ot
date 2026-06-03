// ============================================================
// METROMECANICA — Apps Script: Membrete + Firma de Certificados
// ============================================================

const CONFIG = {
  CARPETA_ORIGEN:    "1q32NVy5hjv5pdOykGXHz8bAE4HrP422l",
  CARPETA_DESTINO:   "1hJkm-2HhTMmU4W12sqfyh2Kdei_lrrX-",
  RAILWAY_URL:       "https://metromecanica-ot-production.up.railway.app/firmar-pdf",
  LOTE_MAX:          5,
  NOMBRE_PROCESADOS: "PROCESADOS"
};


function procesarCertificados() {
  const carpetaOrigen     = DriveApp.getFolderById(CONFIG.CARPETA_ORIGEN);
  const carpetaDestino    = DriveApp.getFolderById(CONFIG.CARPETA_DESTINO);
  const carpetaProcesados = obtenerOCrearSubcarpeta(carpetaOrigen, CONFIG.NOMBRE_PROCESADOS);

  const archivos = carpetaOrigen.getFilesByType(MimeType.PDF);
  const lote     = [];

  while (archivos.hasNext() && lote.length < CONFIG.LOTE_MAX) {
    const archivo = archivos.next();
    if (archivo.getParents().next().getId() !== CONFIG.CARPETA_ORIGEN) continue;
    lote.push(archivo);
  }

  if (lote.length === 0) {
    Logger.log("No hay PDFs pendientes.");
    return;
  }

  Logger.log("Procesando " + lote.length + " archivo(s)...");

  const procesadosOk = [];
  const errores      = [];

  lote.forEach(function(archivo) {
    const nombre = archivo.getName();
    try {
      const blob = archivo.getBlob().setContentType("application/pdf");

      const opciones = {
        method:             "post",
        payload:            { file: blob },   // multipart/form-data automático
        muteHttpExceptions: true
      };

      const respuesta = UrlFetchApp.fetch(CONFIG.RAILWAY_URL, opciones);
      const codigo    = respuesta.getResponseCode();

      if (codigo !== 200) {
        errores.push({ nombre: nombre, error: "HTTP " + codigo + ": " + respuesta.getContentText().substring(0, 300) });
        return;
      }

      // Respuesta es el PDF firmado directamente
      const pdfBlob = respuesta.getBlob().setName(nombre).setContentType("application/pdf");
      carpetaDestino.createFile(pdfBlob);
      procesadosOk.push(archivo);
      Logger.log("✅ Emitido: " + nombre);

    } catch(e) {
      errores.push({ nombre: nombre, error: e.message });
    }
  });

  errores.forEach(function(e) {
    Logger.log("❌ Error en " + e.nombre + ": " + e.error);
  });

  procesadosOk.forEach(function(archivo) {
    carpetaProcesados.addFile(archivo);
    carpetaOrigen.removeFile(archivo);
    Logger.log("📁 Movido: " + archivo.getName());
  });

  Logger.log("Completado. OK: " + procesadosOk.length + " | Errores: " + errores.length);
}


function obtenerOCrearSubcarpeta(padre, nombre) {
  const iter = padre.getFoldersByName(nombre);
  if (iter.hasNext()) return iter.next();
  return padre.createFolder(nombre);
}


function crearTriggerHorario() {
  ScriptApp.newTrigger("procesarCertificados")
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log("Trigger horario creado.");
}
