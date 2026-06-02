// ============================================================
// METROMECANICA — Apps Script: Membrete + Firma de Certificados
// ============================================================
// Carpetas Drive:
//   ORIGEN:  CERTIFICADOS_FIRMARPDF   → 1q32NVy5hjv5pdOykGXHz8bAE4HrP422l
//   DESTINO: 02_CERTIFICADOS_EMITIDOS → 1hJkm-2HhTMmU4W12sqfyh2Kdei_lrrX-
//   PROCESADOS: subcarpeta creada automáticamente dentro de ORIGEN
//
// Railway endpoint: /firmar-lote
// ============================================================

const CONFIG = {
  CARPETA_ORIGEN:    "1q32NVy5hjv5pdOykGXHz8bAE4HrP422l",
  CARPETA_DESTINO:   "1hJkm-2HhTMmU4W12sqfyh2Kdei_lrrX-",
  RAILWAY_URL:       "https://TU-SERVICIO.railway.app/firmar-lote",  // ← cambia esto
  LOTE_MAX:          5,   // PDFs por llamada (evitar timeout)
  NOMBRE_PROCESADOS: "PROCESADOS"
};


// ------------------------------------------------------------
// FUNCIÓN PRINCIPAL — ejecutar manualmente o con trigger
// ------------------------------------------------------------
function procesarCertificados() {
  const carpetaOrigen    = DriveApp.getFolderById(CONFIG.CARPETA_ORIGEN);
  const carpetaDestino   = DriveApp.getFolderById(CONFIG.CARPETA_DESTINO);
  const carpetaProcesados = obtenerOCrearSubcarpeta(carpetaOrigen, CONFIG.NOMBRE_PROCESADOS);

  const archivos = carpetaOrigen.getFilesByType(MimeType.PDF);
  const lote     = [];

  // Recolectar hasta LOTE_MAX PDFs
  while (archivos.hasNext() && lote.length < CONFIG.LOTE_MAX) {
    const archivo = archivos.next();
    // Ignorar archivos dentro de subcarpetas
    if (archivo.getParents().next().getId() !== CONFIG.CARPETA_ORIGEN) continue;
    lote.push(archivo);
  }

  if (lote.length === 0) {
    Logger.log("No hay PDFs pendientes en la carpeta.");
    return;
  }

  Logger.log(`Procesando ${lote.length} archivo(s)...`);

  // Construir payload para Railway
  const documentos = lote.map(archivo => ({
    nombre:    archivo.getName(),
    contenido: Utilities.base64Encode(archivo.getBlob().getBytes())
  }));

  const payload = JSON.stringify({ documentos });

  const opciones = {
    method:      "post",
    contentType: "application/json",
    payload:     payload,
    muteHttpExceptions: true
  };

  let respuesta;
  try {
    respuesta = UrlFetchApp.fetch(CONFIG.RAILWAY_URL, opciones);
  } catch (e) {
    Logger.log("Error llamando al endpoint: " + e.message);
    return;
  }

  if (respuesta.getResponseCode() !== 200) {
    Logger.log("Error del servidor: " + respuesta.getContentText());
    return;
  }

  const resultado = JSON.parse(respuesta.getContentText());

  // Subir PDFs firmados a carpeta destino
  resultado.procesados.forEach(doc => {
    const bytes = Utilities.base64Decode(doc.contenido);
    const blob  = Utilities.newBlob(bytes, MimeType.PDF, doc.nombre);
    carpetaDestino.createFile(blob);
    Logger.log("✅ Emitido: " + doc.nombre);
  });

  // Registrar errores
  if (resultado.errores && resultado.errores.length > 0) {
    resultado.errores.forEach(e => {
      Logger.log("❌ Error en " + e.nombre + ": " + e.error);
    });
  }

  // Mover originales procesados a subcarpeta PROCESADOS
  const nombresOk = new Set(resultado.procesados.map(d => d.nombre));
  lote.forEach(archivo => {
    if (nombresOk.has(archivo.getName())) {
      carpetaProcesados.addFile(archivo);
      carpetaOrigen.removeFile(archivo);
      Logger.log("📁 Movido a PROCESADOS: " + archivo.getName());
    }
  });

  Logger.log(`Lote completado. OK: ${resultado.procesados.length} | Errores: ${resultado.errores.length}`);
}


// ------------------------------------------------------------
// HELPER — obtener o crear subcarpeta
// ------------------------------------------------------------
function obtenerOCrearSubcarpeta(padre, nombre) {
  const iter = padre.getFoldersByName(nombre);
  if (iter.hasNext()) return iter.next();
  return padre.createFolder(nombre);
}


// ------------------------------------------------------------
// TRIGGER — ejecutar automáticamente cada hora (opcional)
// Ejecuta esta función UNA VEZ para activar el trigger
// ------------------------------------------------------------
function crearTriggerHorario() {
  ScriptApp.newTrigger("procesarCertificados")
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log("Trigger horario creado.");
}
