/**
 * ============================================================================
 *  BNH Medical · Sistema Unificado de Pagos
 *  Web App (reporte + panel admin) + OCR automático de comprobantes (Cloud Vision)
 * ============================================================================

 */

// ===================== CONFIGURACIÓN GENERAL =====================
const SHEET_ID = '1cRIKmKsITlmhkCwccEIMaT7faVhYtPVagD3NIT1QbaA';
const NOMBRE_HOJA = 'PAGOS';
const DRIVE_FOLDER_ID = '1X7JIDp4FiV-8tM-9cHo1ISJ_tD4VYTKO'; // carpeta de comprobantes

// NUEVO: carpeta de Drive donde se guardan los RIF adjuntos (Inicial / Reserva).
// Reemplaza el valor por el ID real de tu carpeta de Drive (ver instrucciones
// al final de este archivo, sección "CONFIGURACIÓN DEL RIF").
const DRIVE_FOLDER_ID_RIF = '1MoxUIO5-s9MQ1M3By3CI_sDoBeQ65-CQ';

// NUEVO: configuración de la lista de vendedores. Se asume que la hoja
// "VENDEDORES" vive en el MISMO spreadsheet (SHEET_ID) que la hoja "PAGOS",
// ya que ambas pertenecen al archivo "RECIBO DE PAGOS BNH VIVA". Si en tu
// caso es un archivo distinto, cambia SHEET_ID_VENDEDORES por el ID de ese
// otro spreadsheet.
const SHEET_ID_VENDEDORES = SHEET_ID;
const NOMBRE_HOJA_VENDEDORES = 'VENDEDORES';
const COL_VENDEDOR_LISTA = 2;      // columna B
const FILA_INICIO_VENDEDORES = 2;
const FILA_FIN_VENDEDORES = 50;

// Columnas (1-indexed) — deben coincidir con los encabezados reales de la hoja
const COL_FECHA_REGISTRO = 1;  // A
const COL_NOMBRE         = 2;  // B
const COL_FECHA_PAGO     = 3;  // C
const COL_TIPO_PAGO      = 4;  // D
const COL_MONEDA         = 5;  // E
const COL_TITULAR        = 6;  // F
const COL_ARCHIVO        = 7;  // G  -> URL del comprobante (usada por el OCR)
const COL_REFERENCIA     = 8;  // H  -> resultado OCR
const COL_MONTO          = 9;  // I  -> resultado OCR
const COL_BENEFICIARIO   = 10; // J  -> resultado OCR
const COL_CORREO         = 11; // K  -> correo electrónico ingresado en el formulario
const COL_CONCEPTO_PAGO  = 12; // L  -> Concepto de pago (Inicial / Reserva / pago IVA / Cuota viva)
const COL_TELEFONO       = 13; // M  -> Teléfono ingresado en el formulario
const COL_RIF            = 14; // N  -> URL del RIF adjunto (Inicial / Reserva)
const COL_VENDEDOR       = 15; // O  -> VENDEDOR (Inicial / Reserva)
const COL_COMENTARIO     = 20; // T  -> OBSERVACIONES (comentarios de Servicio Técnico)
const MAX_COMENTARIO     = 500;
const COL_EMPRESA        = 18; // R  -> EMPRESA (usada por el filtro del panel admin)
const COL_TASA_BINANCE   = 22; // V  -> TASA BINANCE (usada por la tarjeta de Bolívares del panel admin)

// Valores de "Tipo de pago" (columna D) que activan el escaneo de SERIALES
// DE BILLETES en lugar del flujo tradicional de Referencia/Monto/Beneficiario.
const TIPOS_PAGO_EFECTIVO = ['EFECTIVO', 'CASH', 'DIVISAS', 'EFECTIVO BS', 'EFECTIVO USD'];

// Valor EXACTO de "Tipo de pago" (columna D) para el cual el Monto (columna I)
// NO se calcula por OCR, sino que viene directamente del formulario web
// (campo "montoManual"). Comparación exacta a propósito: variantes como
// "Efectivo BS" o "Efectivo USD" siguen usando extraerMontoEfectivo() normalmente.
const TIPO_PAGO_MONTO_MANUAL = 'EFECTIVO';

// Valores de "Concepto de pago" (columna L) que exigen adjuntar el RIF y
// seleccionar el Vendedor en el formulario. Comparación en mayúsculas
// contra formData.conceptoPago.
const CONCEPTOS_QUE_REQUIEREN_RIF = ['INICIAL', 'RESERVA'];
// Máximo de adjuntos de soporte por envío (una fila en PAGOS por cada uno)
const MAX_ARCHIVOS_SOPORTE = 10;

// Si el Tipo de pago es exactamente "Efectivo" (monto manual):
// false = el monto va solo en la primera fila (evita sumar el mismo monto N veces)
// true  = el monto se repite en todas las filas
const REPETIR_MONTO_MANUAL_EN_TODAS = false;

/* ============================================================================
 * VALIDACIÓN DE PAGOS DUPLICADOS (columna H - Referencia)
 * ========================================================================== */

// Valores que NUNCA deben considerarse "duplicados" entre sí, aunque coincidan
// textualmente (son placeholders de error o "sin dato", no referencias reales).
const VALORES_REFERENCIA_IGNORADOS = new Set([
  '', 'NO ENCONTRADA', 'NO ENCONTRADO', 'NO DETECTADO',
  'ERROR OCR', 'ERROR AL LEER SERIALES', 'NO SE DETECTÓ TEXTO'
]);

/**
 * Verifica si una Referencia (columna H) ya existe previamente en la hoja
 * PAGOS. Optimizada para leer ÚNICAMENTE el rango H2:H(filaTope) — nunca la
 * hoja completa — minimizando el consumo de cuota de Sheets.
 *
 * @param {Sheet} hoja Hoja "PAGOS"
 * @param {string} referencia Valor detectado en la columna H a validar
 * @param {number} [filaExcluir] Fila que debe excluirse de la comparación
 *        (la fila recién insertada, para no compararla consigo misma)
 * @return {boolean} true si ya existe una referencia igual (normalizada)
 */
function referenciaYaExiste(hoja, referencia, filaExcluir) {
  const refNormalizada = String(referencia || '').trim().toUpperCase();

  if (!refNormalizada || VALORES_REFERENCIA_IGNORADOS.has(refNormalizada)) {
    return false; // nunca bloqueamos por valores vacíos/errores de OCR
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return false;

  const filaTope = filaExcluir ? Math.min(filaExcluir - 1, ultimaFila) : ultimaFila;
  if (filaTope < 2) return false;

  const numFilas = filaTope - 1; // filas 2..filaTope
  // Lectura mínima: SOLO columna H, ninguna otra columna.
  const valores = hoja.getRange(2, COL_REFERENCIA, numFilas, 1).getValues();

  for (let i = 0; i < valores.length; i++) {
    const valor = String(valores[i][0] || '').trim().toUpperCase();
    if (valor && valor === refNormalizada) return true;
  }
  return false;
}


/* ============================================================================
 * WEB APP: doGet, procesarFormulario, obtenerPagos, obtenerVendedores
 * ========================================================================== */

function doGet(e) {
  const params = (e && e.parameter) || {};
 
  // NUEVO: prueba rápida de que la API responde -> .../exec?action=ping
  if (String(params.action || '') === 'ping') return jsonOut_({ ok: true, data: 'pong' });
 
  const page = String(params.page || '').toLowerCase().trim();
  const isAdmin = page === 'admin';
  const isMonitor = page === 'monitor';
 
  try {
    const nombrePlantilla = isMonitor ? 'monitor' : (isAdmin ? 'admin' : 'index');
    const template = HtmlService.createTemplateFromFile(nombrePlantilla);
 
    let titulo = 'Reporte de Pago · BNH Medical';
    if (isAdmin)   titulo = 'Panel Administrativo · BNH Medical';
    if (isMonitor) titulo = 'Monitoreo · BNH Medical';
 
    return template.evaluate()
      .setTitle(titulo)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    console.error('❌ Error en doGet: ' + err.message);
    return HtmlService.createHtmlOutput(
      '<p>Ocurrió un error al cargar la página. Contacta al administrador.</p>'
    ).setTitle('Error · BNH Medical');
  }
}

// Campos obligatorios que debe traer formData para poder procesar el pago.
const CAMPOS_REQUERIDOS_FORMULARIO = [
  'nombre', 'fechaPago', 'conceptoPago', 'tipoPago', 'moneda', 'titular', 'telefono'
];

// Formato esperado para el teléfono -> +58 seguido de 10 dígitos
// (código de ciudad/operadora + número), ej. +584121234567
const REGEX_TELEFONO = /^\+58\d{10}$/;

/**
 * Determina si, para el "Tipo de pago" recibido en el formulario, el Monto
 * debe venir directamente del campo manual "montoManual" en vez de
 * calcularse por OCR. Comparación EXACTA (no "contiene"): solo aplica
 * cuando el usuario eligió exactamente "Efectivo".
 * @param {string} tipoPago Valor crudo de formData.tipoPago
 * @return {boolean}
 */
function requiereMontoManual(tipoPago) {
  return String(tipoPago || '').trim().toUpperCase() === TIPO_PAGO_MONTO_MANUAL;
}

/**
 * Determina si, para el "Concepto de pago" recibido en el formulario, es
 * obligatorio adjuntar el RIF y seleccionar el Vendedor (Inicial / Reserva).
 * @param {string} conceptoPago Valor crudo de formData.conceptoPago
 * @return {boolean}
 */
function requiereRif(conceptoPago) {
  return CONCEPTOS_QUE_REQUIEREN_RIF.includes(String(conceptoPago || '').trim().toUpperCase());
}

/**
 * Indica si el concepto de pago es "Servicio Técnico" (ignora tildes y mayúsculas).
 */
function esServicioTecnico_(conceptoPago) {
  const c = String(conceptoPago || '').trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return c === 'SERVICIO TECNICO';
}

/**
 * Valida que formData traiga todos los campos obligatorios y que el
 * archivo venga en formato Data URL válido ("data:<mime>;base64,<datos>").
 * Si el Tipo de pago es exactamente "Efectivo", exige además el campo
 * "montoManual" (numérico, mayor a 0).
 * Valida además que "telefono" respete el formato +58 + 10 dígitos.
 * Si el Concepto de pago es "Inicial" o "Reserva", exige que venga el
 * archivo del RIF ("archivoRifBase64" / "nombreArchivoRif") en formato
 * Data URL válido, y también el campo "vendedor".
 * Si el Concepto es "Servicio Técnico", valida que el comentario no
 * supere MAX_COMENTARIO caracteres.
 * No modifica nada en la hoja; solo determina si procesarFormulario()
 * puede continuar con seguridad.
 * @param {Object} formData
 * @return {{valido: boolean, error: string}}
 */
/**
 * Devuelve la lista de adjuntos del formulario como [{nombreArchivo, archivoBase64}].
 * Acepta el formato nuevo (formData.archivos) y el antiguo (un solo archivo).
 */
function obtenerArchivosDeFormData_(formData) {
  if (Array.isArray(formData.archivos) && formData.archivos.length > 0) {
    return formData.archivos;
  }
  if (formData.archivoBase64 && formData.nombreArchivo) {
    return [{ nombreArchivo: formData.nombreArchivo, archivoBase64: formData.archivoBase64 }];
  }
  return [];
}



function validarFormData(formData) {
  if (!formData || typeof formData !== 'object') {
    return { valido: false, error: 'No se recibieron datos del formulario.' };
  }

  for (const campo of CAMPOS_REQUERIDOS_FORMULARIO) {
    if (!formData[campo]) {
      return { valido: false, error: 'Falta el campo obligatorio: ' + campo };
    }
  }

  const telefonoLimpio = String(formData.telefono || '').trim();
  if (!REGEX_TELEFONO.test(telefonoLimpio)) {
    return { valido: false, error: 'El teléfono debe tener el formato +58 seguido del código y número (Ej: +584121234567).' };
  }

  // Adjuntos de soporte: de 1 a MAX_ARCHIVOS_SOPORTE
  const archivos = obtenerArchivosDeFormData_(formData);
  if (archivos.length === 0) {
    return { valido: false, error: 'Debes adjuntar al menos un soporte de pago.' };
  }
  if (archivos.length > MAX_ARCHIVOS_SOPORTE) {
    return { valido: false, error: 'Máximo ' + MAX_ARCHIVOS_SOPORTE + ' archivos por envío.' };
  }
  for (let i = 0; i < archivos.length; i++) {
    const a = archivos[i] || {};
    const partesA = String(a.archivoBase64 || '').split(',');
    const matchA = partesA[0] && partesA[0].match(/data:(.*);base64/);
    if (!a.nombreArchivo || partesA.length < 2 || !matchA) {
      return { valido: false, error: 'El archivo #' + (i + 1) + ' no tiene un formato Base64 válido.' };
    }
  }

  if (requiereMontoManual(formData.tipoPago)) {
    const montoManual = Number(formData.montoManual);
    if (!formData.montoManual || isNaN(montoManual) || montoManual <= 0) {
      return { valido: false, error: 'Falta el campo obligatorio: montoManual (monto del pago en efectivo)' };
    }
  }

  if (requiereRif(formData.conceptoPago)) {
    if (!formData.archivoRifBase64 || !formData.nombreArchivoRif) {
      return { valido: false, error: 'Debes adjuntar el RIF para el concepto de pago seleccionado (' + formData.conceptoPago + ').' };
    }
    const partesRif = String(formData.archivoRifBase64).split(',');
    const matchRif = partesRif[0] && partesRif[0].match(/data:(.*);base64/);
    if (partesRif.length < 2 || !matchRif) {
      return { valido: false, error: 'El archivo del RIF no tiene un formato Base64 válido.' };
    }
    if (!formData.vendedor || !String(formData.vendedor).trim()) {
      return { valido: false, error: 'Debes seleccionar el Vendedor para el concepto de pago seleccionado (' + formData.conceptoPago + ').' };
    }
  }

  if (esServicioTecnico_(formData.conceptoPago)) {
    const comentario = String(formData.comentario || '');
    if (comentario.length > MAX_COMENTARIO) {
      return { valido: false, error: 'El comentario no puede superar los ' + MAX_COMENTARIO + ' caracteres.' };
    }
  }

  return { valido: true, error: '' };
}

/**
 * Recibe el formulario público. Por CADA adjunto de soporte (hasta 10):
 * sube el archivo a Drive, agrega una fila en PAGOS con los mismos datos,
 * corre el OCR sobre esa fila y valida duplicados por referencia.
 * El RIF (Inicial / Reserva) se sube una sola vez y su URL se repite en cada fila.
 * El comentario (Servicio Técnico) se repite en cada fila (columna T, OBSERVACIONES).
 */
function procesarFormulario(formData) {
  try {
    const validacion = validarFormData(formData);
    if (!validacion.valido) {
      console.error('❌ Formulario rechazado: ' + validacion.error);
      return { success: false, error: validacion.error };
    }

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(NOMBRE_HOJA);
    if (!sheet) {
      return { success: false, error: 'No se encontró la hoja "' + NOMBRE_HOJA + '" en el spreadsheet.' };
    }

    const archivos = obtenerArchivosDeFormData_(formData);
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

    // --- RIF: se sube UNA sola vez (fail-open) ---
    let rifUrl = '';
    if (formData.archivoRifBase64 && formData.nombreArchivoRif) {
      try {
        const partesRif = formData.archivoRifBase64.split(',');
        const contentTypeRif = partesRif[0].match(/data:(.*);base64/)[1];
        const blobRif = Utilities.newBlob(Utilities.base64Decode(partesRif[1]), contentTypeRif, formData.nombreArchivoRif);
        const fileRif = DriveApp.getFolderById(DRIVE_FOLDER_ID_RIF).createFile(blobRif);
        fileRif.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        rifUrl = fileRif.getUrl();
        console.log('✅ RIF guardado en Drive.');
      } catch (errRif) {
        console.error('⚠️ No se pudo subir el RIF: ' + errRif.message);
      }
    }

    const vendedor = requiereRif(formData.conceptoPago) ? String(formData.vendedor || '').trim() : '';
    const comentario = esServicioTecnico_(formData.conceptoPago)
      ? textoSeguro_(String(formData.comentario || '').trim().substring(0, MAX_COMENTARIO))
      : '';
    const guardados = [];
    const duplicados = [];
    const errores = [];

    for (let idx = 0; idx < archivos.length; idx++) {
      const adjunto = archivos[idx];
      let file = null;
      try {
        // 1. Subir comprobante a Drive
        const partes = adjunto.archivoBase64.split(',');
        const contentType = partes[0].match(/data:(.*);base64/)[1];
        const blob = Utilities.newBlob(Utilities.base64Decode(partes[1]), contentType, adjunto.nombreArchivo);
        file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        // 2. Fila nueva (A-G)
        sheet.appendRow([
          new Date(),
          formData.nombre,
          formData.fechaPago,
          formData.tipoPago,
          formData.moneda,
          formData.titular,
          file.getUrl()
        ]);
        const fila = sheet.getLastRow();

        // 2b. Columnas K a O en una sola escritura:
        // K correo | L concepto | M teléfono | N RIF | O vendedor
        sheet.getRange(fila, COL_CORREO, 1, 5).setValues([[
          formData.correo || '',
          formData.conceptoPago || '',
          String(formData.telefono || '').trim(),
          rifUrl,
          vendedor
        ]]);

        // 2b-bis. Comentario de Servicio Técnico -> columna T (OBSERVACIONES)
        if (comentario) {
          sheet.getRange(fila, COL_COMENTARIO).setValue(comentario);
        }

        // 2c. Monto manual (Efectivo exacto): ANTES del OCR
        if (requiereMontoManual(formData.tipoPago) && (idx === 0 || REPETIR_MONTO_MANUAL_EN_TODAS)) {
          sheet.getRange(fila, COL_MONTO).setValue(Number(formData.montoManual));
        }

        // 3. OCR sobre esta fila
        let ocrOk = true;
        try {
          procesarOCRFila(sheet, fila);
        } catch (ocrErr) {
          ocrOk = false;
          console.error('⚠️ Fila ' + fila + ' guardada pero el OCR falló: ' + ocrErr.message);
        }

        // 4. Validación de duplicado (columna H)
        if (ocrOk) {
          try {
            const referenciaDetectada = sheet.getRange(fila, COL_REFERENCIA).getValue();
            if (referenciaYaExiste(sheet, referenciaDetectada, fila)) {
              sheet.deleteRow(fila);
              try { file.setTrashed(true); } catch (errTrash) {
                console.error('⚠️ No se pudo enviar a la papelera el duplicado: ' + errTrash.message);
              }
              duplicados.push(String(referenciaDetectada).trim());
              continue;
            }
          } catch (errDup) {
            console.error('⚠️ Error al validar duplicado (fail-open): ' + errDup.message);
          }
        }

        guardados.push(fila);
      } catch (errArchivo) {
        console.error('❌ Error con el adjunto #' + (idx + 1) + ' (' + adjunto.nombreArchivo + '): ' + errArchivo.message);
        errores.push(adjunto.nombreArchivo);
        if (file) { try { file.setTrashed(true); } catch (e2) {} }
      }
    }

    // --- Resultado ---
    if (guardados.length === 0) {
      if (duplicados.length > 0) {
        return { success: false, duplicado: true, referencia: duplicados.join(', ') };
      }
      return { success: false, error: 'No se pudo guardar ningún adjunto.' };
    }

    return {
      success: true,
      guardados: guardados.length,
      total: archivos.length,
      duplicados: duplicados,
      errores: errores
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}


/**
 * Devuelve todos los pagos registrados, en el formato que
 * consume el dashboard (admin.html).
 */
function obtenerPagos() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(NOMBRE_HOJA);
    if (!sheet) {
      console.error('❌ obtenerPagos: no se encontró la hoja "' + NOMBRE_HOJA + '".');
      return [];
    }

    const filas = sheet.getDataRange().getValues();
    filas.shift(); // quitar encabezados

    return filas
      .filter(fila => fila[COL_NOMBRE - 1])
      .map(fila => ({
        fechaRegistro:   fila[COL_FECHA_REGISTRO - 1] instanceof Date ? fila[COL_FECHA_REGISTRO - 1].toISOString() : fila[COL_FECHA_REGISTRO - 1],
        nombre:           fila[COL_NOMBRE - 1],
        fechaPago:        fila[COL_FECHA_PAGO - 1] instanceof Date ? Utilities.formatDate(fila[COL_FECHA_PAGO - 1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : fila[COL_FECHA_PAGO - 1],
        conceptoPago:     fila[COL_CONCEPTO_PAGO - 1],
        tipoPago:         fila[COL_TIPO_PAGO - 1],
        moneda:           fila[COL_MONEDA - 1],
        titular:          fila[COL_TITULAR - 1],
        comprobanteUrl:   fila[COL_ARCHIVO - 1],
        referencia:       fila[COL_REFERENCIA - 1],
        monto:            fila[COL_MONTO - 1],
        beneficiario:     fila[COL_BENEFICIARIO - 1],
        correo:           fila[COL_CORREO - 1],
        telefono:         fila[COL_TELEFONO - 1],
        rifUrl:           fila[COL_RIF - 1],
        vendedor:         fila[COL_VENDEDOR - 1],
        comentario:       fila[COL_COMENTARIO - 1],
        empresa:          fila[COL_EMPRESA - 1],          // Columna R, usada por el filtro del panel admin
        tasaBinance:      fila[COL_TASA_BINANCE - 1]      // NUEVO: Columna V, usada para convertir Bs a USD en el panel admin
      }));
  } catch (err) {
    // Un fallo aquí no debe tumbar el panel admin: se registra y se
    // devuelve una lista vacía para que la UI lo maneje con normalidad.
    console.error('❌ Error en obtenerPagos: ' + err.message);
    return [];
  }
}

/**
 * NUEVO: Devuelve la lista de vendedores para poblar el <select> del
 * formulario público. Lee la hoja "VENDEDORES", columna B, filas 2-50.
 */
function obtenerVendedores() {
  try {
    const hoja = SpreadsheetApp.openById(SHEET_ID_VENDEDORES).getSheetByName(NOMBRE_HOJA_VENDEDORES);
    if (!hoja) {
      console.error('❌ obtenerVendedores: no se encontró la hoja "' + NOMBRE_HOJA_VENDEDORES + '".');
      return [];
    }

    const numFilas = FILA_FIN_VENDEDORES - FILA_INICIO_VENDEDORES + 1;
    const valores = hoja.getRange(FILA_INICIO_VENDEDORES, COL_VENDEDOR_LISTA, numFilas, 1).getValues();

    return valores
      .map(fila => String(fila[0] || '').trim())
      .filter(v => v.length > 0);
  } catch (err) {
    console.error('❌ Error en obtenerVendedores: ' + err.message);
    return [];
  }
}

/** Utilidad para incluir parciales HTML si separas el CSS/JS */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/* ============================================================================
 * OCR: núcleo compartido + entrada por trigger (Google Form nativo, opcional)
 * ========================================================================== */

/**
 * Núcleo del OCR: procesa una fila específica de la hoja PAGOS.
 * Lo llaman tanto procesarFormulario() (Web App, automático) como
 * alEnviarFormulario() (trigger de un Google Form nativo, opcional).
 */
function procesarOCRFila(hoja, fila) {
  if (!hoja || !fila || fila < 2) {
    console.error('❌ procesarOCRFila: parámetros inválidos (hoja o número de fila).');
    return;
  }

  console.log('📍 Procesando OCR de la fila: ' + fila);

  const valorCelda = hoja.getRange(fila, COL_ARCHIVO).getValue();
  console.log('🔗 Enlace detectado: ' + valorCelda);

  if (!valorCelda) {
    console.log('❌ La celda del archivo está vacía. Abortando OCR.');
    return;
  }

  const fileId = extraerIdDesdeUrl(valorCelda);
  if (!fileId) {
    console.log('❌ No se pudo extraer el ID del enlace.');
    return;
  }

  console.log('✅ ID extraído: ' + fileId + '. Enviando a Cloud Vision...');

  const archivo = DriveApp.getFileById(fileId);

  // Se pide UNA sola vez texto + posición de cada palabra (bounding box).
  // El texto (datosOCR.texto) se sigue usando exactamente igual que antes
  // para fecha/seriales/monto. Las posiciones (datosOCR.palabras) se usan
  // además para la extracción de Referencia "por posición".
  let datosOCR;
  try {
    datosOCR = extraerDatosConVision(archivo);
  } catch (errVision) {
    // La imagen puede estar borrosa, dañada o inaccesible: no detenemos el script.
    console.error('❌ Error al ejecutar OCR con Vision: ' + errVision.message);
    hoja.getRange(fila, COL_REFERENCIA).setValue('ERROR OCR');
    return;
  }
  const textoExtraido = datosOCR.texto;
  console.log('📝 TEXTO COMPLETO DEL RECIBO:\n' + textoExtraido);

  // --- validar/ajustar la Fecha de pago (columna C) contra la fecha
  // detectada en el comprobante. Fail-open: nunca interrumpe lo que sigue. ---
  validarFechaPagoConOCR_(hoja, fila, textoExtraido);

  // --- Leer el Tipo de pago (columna D) para decidir la rama a ejecutar ---
  const tipoPagoRaw = hoja.getRange(fila, COL_TIPO_PAGO).getValue();
  const tipoPagoUpper = (tipoPagoRaw || '').toString().trim().toUpperCase();

  if (esTipoPagoEfectivo(tipoPagoUpper)) {
    // ===================== CASO A: EFECTIVO =====================
    // Se escanean TODOS los seriales de billete presentes en la imagen
    // y se guardan concatenados por comas en la columna de Referencia (H).
    // Además, se calcula el monto total sumando las denominaciones detectadas
    // — EXCEPTO cuando el Tipo de pago es EXACTAMENTE "Efectivo": en ese caso
    // el Monto ya fue guardado en procesarFormulario() desde el campo manual
    // "montoManual", así que aquí NO se llama a extraerMontoEfectivo() ni se
    // sobrescribe la columna I.
    try {
      const seriales = extraerSerialesDeBilletes(textoExtraido);
      const serialesTexto = seriales.length > 0 ? seriales.join(', ') : 'No detectado';

      hoja.getRange(fila, COL_REFERENCIA).setValue(serialesTexto);

      if (requiereMontoManual(tipoPagoUpper)) {
        console.log('ℹ️ Tipo de pago "Efectivo" (exacto): se omite extraerMontoEfectivo(); el Monto ya viene del formulario (montoManual).');
      } else {
        const montoDetectado = extraerMontoEfectivo(textoExtraido, seriales.length);
        hoja.getRange(fila, COL_MONTO).setValue(montoDetectado);
        console.log(`✅ MONTO TOTAL DETECTADO (extraerMontoEfectivo): ${montoDetectado}`);
      }

      // Si el tipo de pago es "Efectivo" o "Bs", el Beneficiario se toma de
      // la columna F (Titular) en vez de 'No aplica'.
      if (debeUsarTitularComoBeneficiario(tipoPagoUpper)) {
        const titularEfectivo = hoja.getRange(fila, COL_TITULAR).getValue();
        hoja.getRange(fila, COL_BENEFICIARIO).setValue(titularEfectivo || 'No encontrado');
      } else {
        hoja.getRange(fila, COL_BENEFICIARIO).setValue('No aplica (pago en efectivo)');
      }

      console.log(`✅ SERIALES DETECTADOS (${seriales.length}): ${serialesTexto}`);
    } catch (errSeriales) {
      // Imagen borrosa o sin seriales legibles: se registra el error sin detener el script.
      console.error('❌ Error al extraer seriales de billetes: ' + errSeriales.message);
      hoja.getRange(fila, COL_REFERENCIA).setValue('ERROR AL LEER SERIALES');
      if (!requiereMontoManual(tipoPagoUpper)) {
        hoja.getRange(fila, COL_MONTO).setValue('ERROR AL LEER MONTO');
      }
    }

  } else {
    // ===================== CASO B: ZELLE / BINANCE / TRANSFERENCIA / PAGO MÓVIL =====================
    // La Referencia se busca primero "por posición" (usando las coordenadas
    // de cada palabra que ya devuelve Cloud Vision), y solo si eso no da
    // resultado se recurre al método original por texto plano.
    let referenciaDetectada = extraerReferenciaRobusta(datosOCR);
    let montoDetectado = extraerMonto(textoExtraido);
    let beneficiarioDetectado = debeUsarTitularComoBeneficiario(tipoPagoUpper)
      ? (hoja.getRange(fila, COL_TITULAR).getValue() || 'No encontrado')
      : extraerBeneficiario(textoExtraido);

    // NUEVO: aprendizaje inteligente (fail-open: si algo falla, el flujo sigue igual)
    try {
      const completado = aprenderYCompletar_(
        hoja, fila, textoExtraido,
        { referencia: referenciaDetectada, monto: montoDetectado, beneficiario: beneficiarioDetectado },
        { beneficiario: debeUsarTitularComoBeneficiario(tipoPagoUpper) }
      );
      referenciaDetectada = completado.referencia;
      montoDetectado = completado.monto;
      beneficiarioDetectado = completado.beneficiario;
    } catch (errAprende) {
      console.error('⚠️ Aprendizaje OCR (fail-open): ' + errAprende.message);
    }

    hoja.getRange(fila, COL_REFERENCIA).setValue(referenciaDetectada);
    hoja.getRange(fila, COL_MONTO).setValue(montoDetectado);
    hoja.getRange(fila, COL_BENEFICIARIO).setValue(beneficiarioDetectado);

    console.log(`✅ DATOS GUARDADOS -> Ref: ${referenciaDetectada} | Monto: ${montoDetectado} | Benef: ${beneficiarioDetectado}`);
  }
}

/**
 * Determina si el "Tipo de pago" corresponde a Efectivo/Divisas/Cash,
 * lo cual activa el escaneo de seriales de billetes en lugar del
 * flujo tradicional de Referencia/Monto/Beneficiario.
 * @param {string} tipoPagoUpper Valor de la columna D, ya en mayúsculas
 * @return {boolean}
 */
function esTipoPagoEfectivo(tipoPagoUpper) {
  return TIPOS_PAGO_EFECTIVO.some(t => tipoPagoUpper.indexOf(t) !== -1);
}

/**
 * Determina si el "Tipo de pago" (columna D) es "Efectivo" o "Bs". En ese
 * caso, el Beneficiario (columna J) NO se calcula mediante
 * extraerBeneficiario() sobre el texto OCR, sino que se toma directamente
 * de la columna F (Titular).
 * @param {string} tipoPagoUpper Valor de la columna D, ya en mayúsculas
 * @return {boolean}
 */
function debeUsarTitularComoBeneficiario(tipoPagoUpper) {
  return tipoPagoUpper.indexOf('EFECTIVO') !== -1 || tipoPagoUpper.indexOf('BS') !== -1;
}

/* ============================================================================
 * VALIDACIÓN DE FECHA DE PAGO (columna C) contra la fecha leída del OCR
 * ========================================================================== */

/**
 * Orquesta la validación de la fecha de pago para una fila ya procesada
 * por OCR. Compara la fecha ingresada en el formulario (columna C, tal como
 * quedó guardada por procesarFormulario) contra la fecha que aparece en el
 * texto del comprobante (textoExtraido de Cloud Vision).
 *
 * Reglas:
 *  - Si ambas fechas coinciden (mismo día/mes/año) -> se conserva la del formulario.
 *  - Si difieren -> prevalece la fecha detectada en el comprobante.
 *  - Si el OCR no encontró ninguna fecha -> se conserva la del formulario, sin tocar la celda.
 *
 * Es "fail-open": cualquier error se registra en consola y la función
 * simplemente no modifica la columna C, para no interrumpir el resto del
 * flujo de procesarOCRFila (referencia/monto/beneficiario siguen su curso).
 *
 * @param {Sheet} hoja Hoja "PAGOS"
 * @param {number} fila Fila recién procesada
 * @param {string} textoExtraido Texto ya obtenido por Cloud Vision en procesarOCRFila
 */
function validarFechaPagoConOCR_(hoja, fila, textoExtraido) {
  try {
    const valorActualCelda = hoja.getRange(fila, COL_FECHA_PAGO).getValue();
    const fechaFormularioDate = normalizarFechaFlexible_(valorActualCelda);
    const fechaSoporteDate = extraerFechaDeTexto_(textoExtraido);

    if (!fechaSoporteDate) {
      console.log('ℹ️ Validación de fecha: el comprobante no arrojó ninguna fecha reconocible. Se conserva la del formulario.');
      return;
    }

    if (!fechaFormularioDate) {
      console.log('⚠️ Validación de fecha: la celda C no tenía una fecha interpretable. Se usa la del comprobante.');
      hoja.getRange(fila, COL_FECHA_PAGO).setValue(fechaSoporteDate);
      console.log('✅ Fecha final (formato dd/mm/aaaa): ' + formatearFechaDDMMAAAA_(fechaSoporteDate));
      return;
    }

    const fechaFinalDate = sonMismaFecha_(fechaFormularioDate, fechaSoporteDate)
      ? fechaFormularioDate
      : fechaSoporteDate;

    hoja.getRange(fila, COL_FECHA_PAGO).setValue(fechaFinalDate);

    console.log(
      sonMismaFecha_(fechaFormularioDate, fechaSoporteDate)
        ? '✅ Fecha del formulario coincide con el comprobante.'
        : '⚠️ Fecha del formulario difiere del comprobante — se usa la fecha del comprobante.'
    );
    console.log('✅ Fecha final (formato dd/mm/aaaa): ' + formatearFechaDDMMAAAA_(fechaFinalDate));

  } catch (errFecha) {
    // Fail-open: un problema aquí nunca debe tumbar el OCR de referencia/monto/beneficiario.
    console.error('❌ Error al validar la fecha de pago con OCR: ' + errFecha.message);
  }
}

/**
 * Busca y normaliza una fecha dentro de un texto libre (resultado de OCR).
 * Soporta los formatos más comunes en comprobantes:
 *   - dd/mm/aaaa, dd-mm-aaaa, dd.mm.aaaa
 *   - aaaa/mm/dd, aaaa-mm-dd (ISO)
 *   - "5 de agosto de 2026", "05 Ago 2026" (nombre de mes en español)
 *
 * @param {string} texto
 * @return {Date|null}
 * @private
 */
function extraerFechaDeTexto_(texto) {
  if (!texto) return null;

  const meses = {
    'enero': 1, 'ene': 1, 'febrero': 2, 'feb': 2, 'marzo': 3, 'mar': 3,
    'abril': 4, 'abr': 4, 'mayo': 5, 'may': 5, 'junio': 6, 'jun': 6,
    'julio': 7, 'jul': 7, 'agosto': 8, 'ago': 8,
    'septiembre': 9, 'setiembre': 9, 'sep': 9, 'sept': 9,
    'octubre': 10, 'oct': 10, 'noviembre': 11, 'nov': 11, 'diciembre': 12, 'dic': 12
  };

  const regexNumericaYMD = /\b(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})\b/;
  const regexNumericaDMY = /\b(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})\b/;
  const regexTextoMes = /\b(\d{1,2})\s*(?:de)?\s*([A-Za-zÁÉÍÓÚáéíóúñÑ]+)\s*(?:de|,)?\s*(\d{4})\b/;

  let match;

  match = texto.match(regexNumericaYMD);
  if (match) {
    const anio = parseInt(match[1], 10), mes = parseInt(match[2], 10), dia = parseInt(match[3], 10);
    const fecha = new Date(anio, mes - 1, dia);
    if (esFechaValida_(fecha, dia, mes, anio)) return fecha;
  }

  match = texto.match(regexNumericaDMY);
  if (match) {
    const dia = parseInt(match[1], 10), mes = parseInt(match[2], 10);
    let anio = parseInt(match[3], 10);
    if (anio < 100) anio += 2000;
    const fecha = new Date(anio, mes - 1, dia);
    if (esFechaValida_(fecha, dia, mes, anio)) return fecha;
  }

  match = texto.match(regexTextoMes);
  if (match) {
    const dia = parseInt(match[1], 10);
    const nombreMes = match[2].toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const mes = meses[nombreMes];
    const anio = parseInt(match[3], 10);
    if (mes) {
      const fecha = new Date(anio, mes - 1, dia);
      if (esFechaValida_(fecha, dia, mes, anio)) return fecha;
    }
  }

  return null;
}

/**
 * Convierte el valor actual de la celda "Fecha pago" (puede llegar como
 * objeto Date -si Sheets ya la autoconvirtió- o como string crudo tipo
 * "2026-08-05") en un objeto Date normalizado.
 * @param {Date|string} valor
 * @return {Date|null}
 * @private
 */
function normalizarFechaFlexible_(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return valor;

  const intentoDirecto = new Date(valor);
  if (!isNaN(intentoDirecto.getTime())) return intentoDirecto;

  return extraerFechaDeTexto_(String(valor));
}

/** @private */
function esFechaValida_(fecha, dia, mes, anio) {
  return fecha.getFullYear() === anio && fecha.getMonth() === mes - 1 && fecha.getDate() === dia;
}

/** Compara dos fechas solo por día/mes/año (ignora horas). @private */
function sonMismaFecha_(fechaA, fechaB) {
  return fechaA.getFullYear() === fechaB.getFullYear() &&
         fechaA.getMonth() === fechaB.getMonth() &&
         fechaA.getDate() === fechaB.getDate();
}

/** Formatea una fecha en dd/mm/aaaa (con ceros a la izquierda). @private */
function formatearFechaDDMMAAAA_(fecha) {
  if (!fecha) return '';
  const dd = String(fecha.getDate()).padStart(2, '0');
  const mm = String(fecha.getMonth() + 1).padStart(2, '0');
  const aaaa = fecha.getFullYear();
  return `${dd}/${mm}/${aaaa}`;
}

/* ============================================================================
 * FUNCIONES DE EXTRACCIÓN (regex) y llamada a Vision API
 * ========================================================================== */

/**
 * Extrae el ID de archivo de Drive a partir de la URL guardada en la
 * columna "Comprobante" (columna G).
 * @param {string} valorCelda Valor crudo de la celda (URL de Drive)
 * @return {?string} ID del archivo, o null si no se pudo extraer
 */
function extraerIdDesdeUrl(valorCelda) {
  const texto = String(valorCelda);
  const match = texto.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

// ------------------------------------------------------------------
// Hace la llamada a Vision UNA sola vez y devuelve texto + posición de
// cada palabra (bounding box). Todo lo demás se construye sobre esto.
// ------------------------------------------------------------------
function extraerDatosConVision(archivoDrive) {
  if (!archivoDrive) {
    console.error('❌ extraerDatosConVision: no se recibió un archivo de Drive válido.');
    return { texto: 'Error OCR', palabras: [] };
  }

  let blob, base64;
  try {
    blob = archivoDrive.getBlob();
    base64 = Utilities.base64Encode(blob.getBytes());
  } catch (errBlob) {
    // Archivo corrupto, eliminado de Drive, o sin permisos de lectura.
    console.error('❌ Error al leer el archivo de Drive: ' + errBlob.message);
    return { texto: 'Error OCR', palabras: [] };
  }

  const url = 'https://vision.googleapis.com/v1/images:annotate';

  const payload = {
    requests: [
      {
        image: { content: base64 },
        features: [{ type: 'TEXT_DETECTION' }]
      }
    ]
  };

  const opciones = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let respuesta, json;
  try {
    respuesta = UrlFetchApp.fetch(url, opciones);
    json = JSON.parse(respuesta.getContentText());
  } catch (errFetch) {
    // Timeout, problema de red, o respuesta que no es JSON válido.
    console.error('❌ Error de red/parseo al llamar a Cloud Vision: ' + errFetch.message);
    return { texto: 'Error OCR', palabras: [] };
  }

  const codigo = respuesta.getResponseCode();

  if (codigo !== 200) {
    Logger.log('Error de Vision API (' + codigo + '): ' + respuesta.getContentText());
    return { texto: 'Error OCR', palabras: [] };
  }

  const resp = json.responses && json.responses[0];

  if (!resp || !resp.fullTextAnnotation) {
    return { texto: 'No se detectó texto', palabras: [] };
  }

  const texto = resp.fullTextAnnotation.text;

  // textAnnotations[0] = bloque de texto completo (se ignora aquí).
  // Desde el índice 1 en adelante, cada entrada es UNA palabra con su
  // boundingPoly (4 vértices) -> de ahí sacamos la posición real en la imagen.
  const palabras = [];
  const anotaciones = resp.textAnnotations || [];

  for (let i = 1; i < anotaciones.length; i++) {
    const anot = anotaciones[i];
    const vertices = anot.boundingPoly && anot.boundingPoly.vertices;
    if (!vertices || vertices.length === 0) continue;

    const xs = vertices.map(v => v.x || 0);
    const ys = vertices.map(v => v.y || 0);

    palabras.push({
      texto: anot.description,
      xMin: Math.min(...xs),
      xMax: Math.max(...xs),
      yMin: Math.min(...ys),
      yMax: Math.max(...ys),
      xCentro: (Math.min(...xs) + Math.max(...xs)) / 2,
      yCentro: (Math.min(...ys) + Math.max(...ys)) / 2
    });
  }

  return { texto, palabras };
}

// ------------------------------------------------------------------
// COMPATIBILIDAD: se mantiene igual que antes para no romper nada que
// ya la esté llamando. Internamente reusa extraerDatosConVision.
// ------------------------------------------------------------------
function extraerTextoConVision(archivoDrive) {
  return extraerDatosConVision(archivoDrive).texto;
}

// ------------------------------------------------------------------
// Agrupa las palabras por "banda" horizontal = un billete físico.
// Asume billetes apilados/extendidos en fila (variación vertical entre
// billetes >> variación vertical entre palabras del mismo billete).
// ------------------------------------------------------------------
function agruparPalabrasPorBanda(palabras) {
  if (!palabras || palabras.length === 0) return [];

  const alturas = palabras.map(p => p.yMax - p.yMin).sort((a, b) => a - b);
  const alturaMediana = alturas[Math.floor(alturas.length / 2)] || 20;

  // Margen de tolerancia entre palabras del MISMO billete. Generoso a
  // propósito: es preferible fundir dos billetes en una banda (raro, se
  // corrige por el emparejamiento de más abajo) que partir uno en dos.
  const TOLERANCIA = alturaMediana * 3;

  const ordenadas = [...palabras].sort((a, b) => a.yCentro - b.yCentro);

  const bandas = [];
  let bandaActual = [ordenadas[0]];
  let yMaxBandaActual = ordenadas[0].yMax;

  for (let i = 1; i < ordenadas.length; i++) {
    const palabra = ordenadas[i];
    if (palabra.yMin - yMaxBandaActual <= TOLERANCIA) {
      bandaActual.push(palabra);
      yMaxBandaActual = Math.max(yMaxBandaActual, palabra.yMax);
    } else {
      bandas.push(bandaActual);
      bandaActual = [palabra];
      yMaxBandaActual = palabra.yMax;
    }
  }
  bandas.push(bandaActual);

  return bandas.map(banda => {
    const ordenPorX = [...banda].sort((a, b) => a.xCentro - b.xCentro);
    return {
      texto: ordenPorX.map(p => p.texto).join(' '),
      yCentro: banda.reduce((s, p) => s + p.yCentro, 0) / banda.length
    };
  });
}

// ------------------------------------------------------------------
// Cuenta cada billete individualmente:
// ------------------------------------------------------------------

function contarBilletesIndividualmente(archivoDrive) {
  const datos = extraerDatosConVision(archivoDrive);

  if (!datos.palabras || datos.palabras.length === 0) {
    // Sin posiciones (error OCR, o Vision no devolvió textAnnotations):
    // cae de vuelta al método anterior, sobre texto completo.
    const seriales = extraerSerialesDeBilletes(datos.texto);
    const monto = extraerMontoEfectivo(datos.texto, seriales.length);
    return { billetes: [], total: monto, seriales, metodo: 'texto_completo_sin_posicion' };
  }

  const bandas = agruparPalabrasPorBanda(datos.palabras);

  const bandasProcesadas = bandas.map(banda => {
    const seriales = extraerSerialesDeBilletes(banda.texto);
    const denominacion = extraerMontoEfectivo(banda.texto); // función original, SIN CAMBIOS, aplicada por banda
    return {
      yCentro: banda.yCentro,
      seriales,
      denominacion: typeof denominacion === 'number' ? denominacion : null
    };
  });

  const bandasConDenominacion = bandasProcesadas.filter(b => b.denominacion !== null);
  const billetesPorSerial = new Map(); // serial -> { denominacion, inferida }

  bandasProcesadas.forEach(banda => {
    if (banda.seriales.length === 0) return; // banda sin serial = no es un billete contable

    let denominacionFinal = banda.denominacion;
    let inferida = false;

    if (denominacionFinal === null && bandasConDenominacion.length > 0) {
      let masCercana = bandasConDenominacion[0];
      let distMin = Math.abs(banda.yCentro - masCercana.yCentro);
      for (const candidata of bandasConDenominacion) {
        const dist = Math.abs(banda.yCentro - candidata.yCentro);
        if (dist < distMin) { distMin = dist; masCercana = candidata; }
      }
      denominacionFinal = masCercana.denominacion;
      inferida = true;
    }

    banda.seriales.forEach(serial => {
      // Si el serial ya fue registrado por otra banda (posible solape), no se duplica.
      if (!billetesPorSerial.has(serial)) {
        billetesPorSerial.set(serial, { denominacion: denominacionFinal, denominacionInferida: inferida });
      }
    });
  });

  // --- RECONCILIACIÓN: recuperar seriales perdidos por el bandeo por posición ---
  // Se extraen los seriales sobre el TEXTO COMPLETO (sin dividir en bandas) como
  // fuente de verdad adicional. Si el conteo por bandas encontró menos seriales
  // que el texto completo, se agregan los faltantes con la denominación más
  // frecuente detectada (moda) entre las bandas que sí tuvieron denominación.
  const serialesTextoCompleto = extraerSerialesDeBilletes(datos.texto);

  if (serialesTextoCompleto.length > billetesPorSerial.size) {
    // Calcular la denominación más frecuente (moda) como mejor estimación
    // para los billetes que el bandeo no logró ubicar.
    let denominacionModa = null;
    if (bandasConDenominacion.length > 0) {
      const conteoPorDenominacion = {};
      bandasConDenominacion.forEach(b => {
        conteoPorDenominacion[b.denominacion] = (conteoPorDenominacion[b.denominacion] || 0) + 1;
      });
      denominacionModa = Number(
        Object.entries(conteoPorDenominacion).sort((a, b) => b[1] - a[1])[0][0]
      );
    } else if (billetesPorSerial.size > 0) {
      // Si ninguna banda tuvo denominación propia, usar la que ya se infirió antes.
      const valores = [...billetesPorSerial.values()].map(b => b.denominacion).filter(v => v !== null);
      if (valores.length > 0) denominacionModa = valores[0];
    }

    serialesTextoCompleto.forEach(serial => {
      if (!billetesPorSerial.has(serial)) {
        billetesPorSerial.set(serial, {
          denominacion: denominacionModa,
          denominacionInferida: true,
          recuperadoPorReconciliacion: true
        });
      }
    });

    console.log(`⚠️ Reconciliación aplicada: el bandeo por posición encontró ${billetesPorSerial.size - (serialesTextoCompleto.length - billetesPorSerial.size)} seriales, pero el texto completo tenía ${serialesTextoCompleto.length}. Se recuperaron los faltantes.`);
  }

  const billetes = [...billetesPorSerial.entries()].map(([serial, info]) => ({
    serial,
    denominacion: info.denominacion,
    denominacionInferida: info.denominacionInferida,
    recuperadoPorReconciliacion: info.recuperadoPorReconciliacion || false
  }));

  const total = billetes.reduce((sum, b) => sum + (b.denominacion || 0), 0);

  console.log(`💵 ${billetes.length} billete(s) contados individualmente | Total: $${total}`);
  billetes.forEach(b => {
    let etiqueta = '';
    if (b.recuperadoPorReconciliacion) etiqueta = ' (recuperado por reconciliación, denominación inferida por moda)';
    else if (b.denominacionInferida) etiqueta = ' (denominación inferida por cercanía)';
    console.log(`   Serial ${b.serial} -> $${b.denominacion}${etiqueta}`);
  });

  return { billetes, total, metodo: 'bandas_por_posicion' };
}



/* ============================================================================
 * FUNCIONES DE EXTRACCIÓN REFERENCIA
 * ========================================================================== */

/**
 * Extrae la Referencia usando la POSICIÓN de cada palabra en la imagen
 * (bounding boxes que ya devuelve Cloud Vision), en vez de depender del
 * orden en que el texto plano quedó armado.
 *
 * Por qué esto resuelve el problema de la marca de agua: cuando
 * "www.bancamiga.com" queda superpuesta sobre "Referencia: 2510087091",
 * Vision las reporta como palabras/tokens SEPARADOS, cada uno con su propia
 * caja — nunca se mezclan letra por letra dentro de un mismo token. Así que
 * en vez de parsear una cadena de texto donde el orden de lectura pudo
 * revolverse, buscamos literalmente qué palabras están "a la derecha de
 * Referencia, en su misma fila" y descartamos por CONTENIDO las que
 * pertenezcan a la marca de agua (www / bancamiga / .com / redes sociales),
 * quedándonos solo con las que sí son parte del número.
 *
 * @param {{texto: string, palabras: Array}} datosOCR resultado de extraerDatosConVision()
 * @return {?string} Referencia detectada, o null si no se pudo (para que el
 *         llamador recurra al método por texto plano como respaldo).
 */
function extraerReferenciaPorPosicion(datosOCR) {
  if (!datosOCR || !datosOCR.palabras || datosOCR.palabras.length === 0) return null;

  const palabras = datosOCR.palabras;

  // Patrón de etiqueta: cualquier palabra que EMPIECE como "Referencia",
  // "Ref", "Confirmation", etc. (sin dos puntos, Vision los separa aparte).
  const PATRON_ETIQUETA = /^(REFERENCIA|REF\.?|CONFIRMATION|TRANSACTION)$/i;
  const idxLabel = palabras.findIndex(p => PATRON_ETIQUETA.test(p.texto.replace(/[:.,]/g, '')));
  if (idxLabel === -1) return null;

  const labelWord = palabras[idxLabel];
  const alturaLabel = Math.max(labelWord.yMax - labelWord.yMin, 1);
  // Tolerancia vertical generosa (misma "fila" del recibo, aunque el OCR
  // recorte cada palabra con un poquito de variación en Y).
  const TOLERANCIA_Y = alturaLabel * 0.8;

  // Candidatas: misma fila (Y parecido) y a la derecha del label en X.
  const candidatas = palabras.filter(p =>
    p !== labelWord &&
    Math.abs(p.yCentro - labelWord.yCentro) <= TOLERANCIA_Y &&
    p.xCentro > labelWord.xCentro
  );

  if (candidatas.length === 0) return null;

  // Descartar por CONTENIDO cualquier palabra que sea parte de la marca de
  // agua / redes sociales / web del banco, sin importar en qué posición del
  // texto plano hubiera caído.
  const PATRON_RUIDO = /WWW|BANCAMIGA|\.COM|TUBANCA|INSTAGRAM|FACEBOOK|TIKTOK/i;
  const limpias = candidatas.filter(p => !PATRON_RUIDO.test(p.texto));

  if (limpias.length === 0) return null;

  limpias.sort((a, b) => a.xCentro - b.xCentro);

  // Caso normal: la referencia suele venir como un solo token ya completo.
  for (const p of limpias) {
    const limpio = p.texto.replace(/[^A-Z0-9]/gi, '');
    if (/^[A-Z0-9]{6,20}$/i.test(limpio) && /\d/.test(limpio)) {
      return limpio.toUpperCase();
    }
  }

  // Respaldo: si Vision partió el número en varios tokens pequeños
  // (ej. "251" "0087" "091"), se concatenan en orden de X.
  const unido = limpias.map(p => p.texto).join('').replace(/[^A-Z0-9]/gi, '');
  if (/\d{6,}/.test(unido)) return unido.toUpperCase();

  return null;
}

/**
 * Punto de entrada único para obtener la Referencia. Intenta primero el
 * método por posición (más confiable ante marcas de agua superpuestas) y,
 * si no da resultado, cae al método original por texto plano — así ningún
 * comprobante se queda sin intentar todos los caminos disponibles.
 * @param {{texto: string, palabras: Array}} datosOCR resultado de extraerDatosConVision()
 * @return {string} Referencia detectada, o 'No encontrada'
 */
function extraerReferenciaRobusta(datosOCR) {
  try {
    const porPosicion = extraerReferenciaPorPosicion(datosOCR);
    if (porPosicion) {
      console.log('✅ Referencia obtenida por POSICIÓN (bounding boxes): ' + porPosicion);
      return porPosicion;
    }
  } catch (errPos) {
    // Fail-open: si el método por posición falla por lo que sea, se cae al
    // método por texto plano sin interrumpir el flujo.
    console.error('⚠️ Error en extraerReferenciaPorPosicion, se usa el método por texto: ' + errPos.message);
  }

  const porTexto = extraerReferencia(datosOCR.texto);
  console.log('ℹ️ Referencia obtenida por TEXTO PLANO (respaldo): ' + porTexto);
  return porTexto;
}

/**
 * Extrae el número de referencia/transacción de un comprobante (Zelle,
 * Binance, transferencia bancaria, pago móvil) a partir del texto OCR.
 * @param {string} texto Texto completo devuelto por Cloud Vision
 * @return {string} Referencia detectada, o 'No encontrada'
 */
function extraerReferencia(texto) {
  const lineasRaw = texto.split('\n').map(l => l.trim()).filter(l => l.length > 0);

  // 1. Etiquetas conocidas (Zelle, transferencias, pago móvil, etc.)
  const ETIQUETAS = [
    'ID DE ORDEN',
    'N[UÚ]MERO DE TRANSACCI[OÓ]N',
    'TRANSACTION\\s*NUMBER',
    'N[UÚ]MERO DE REFERENCIA',
    'N[UÚ]MERO DE CONFIRMACI[OÓ]N',
    'NRO\\.?\\s*DE\\s*REFERENCIA',
    'N[°º]\\s*DE\\s*REFERENCIA',
    'REFERENCIA BANCARIA',
    'REFERENCIA',
    'REF\\.?\\s*[:#]',
    'CONF(?:IRMATION)?\\s*(?:NUMBER|#|:)',
    'TRANSACTION\\s*ID',
    'PAYMENT ID'
  ];
  const etiquetaRegex = new RegExp('^(?:' + ETIQUETAS.join('|') + ')\\s*:?\\s*(.*)$', 'i');

  // 2. Palabras que NUNCA son una referencia, aunque pasen los regex de formato
  const PALABRAS_PROHIBIDAS = new Set([
    'ESTADO', 'ENTREGADO', 'ACEPTADO', 'PENDIENTE', 'COMPLETADO', 'CANCELADO',
    'RECHAZADO', 'PROCESANDO', 'DISPONIBLE', 'MENSAJE', 'DESTINATARIO',
    'CHECKING', 'SAVINGS', 'TOTAL', 'PAGA', 'ENVIA', 'ENVIADA', 'ENVIADO',
    'RECIBIDO', 'CONCEPTO', 'IMPORTE', 'MONTO', 'FECHA', 'HORA', 'BANCO',
    'CUENTA', 'DINERO', 'INSTANTES', 'CANCELAR', 'DETALLES', 'RESUMEN'
  ]);

  // 2b. Líneas de contexto que jamás deben aportar un candidato a referencia
  // (sello del cajero, C.I. del depositante, número de cuenta, teléfono de
  // atención al cliente, redes sociales / web del banco).
  const CONTEXTO_PROHIBIDO = /C\.?\s*I\.?\s*DEL\s*DEPOSIT|DEPOSITANTE|FIRMA|SELLO\s*DEL\s*CAJERO|CAJERO|^CAJA\b|TITULAR|TUBANCA|WWW\.|\.COM|@|INSTAGRAM|FACEBOOK|TIKTOK|S[IÍ]GUENOS|REDES\s*SOCIALES|^0500\b|N[UÚ]MERO\s*DE\s*CUENTA/i;

  const lineas = lineasRaw.map(l => ({
    texto: l,
    prohibida: CONTEXTO_PROHIBIDO.test(l)
  }));

  const esInvalido = (v) => {
    if (!v) return true;
    const vUpper = v.toUpperCase();
    if (PALABRAS_PROHIBIDAS.has(vUpper)) return true;
    if (/^\d{2}[\/\-]\d{2}[\/\-]\d{4}/.test(v)) return true;
    if (/^[a-z]{3}\s*\d{1,2},?\s*\d{4}/i.test(v)) return true;
    if (/^[VJEG]-?\d{5,9}$/i.test(v)) return true;
    if (/^\*+\d+$/.test(v)) return true;
    if (/\(\.\.\.\d+\)/.test(v)) return true;
    if (/^0\d{3}-?\d{7}$/.test(v)) return true;
    if (/^\d{1,2}:\d{2}/.test(v)) return true;
    // Patrón típico de sello de cajero, ej. "CA1983", "CA2190" (2 letras + 3-4
    // dígitos cortos, normalmente pegado a "C.I. del Depositante")
    if (/^[A-Z]{2}\d{3,4}$/i.test(v)) return true;
    // Los números de cuenta venezolanos tienen exactamente 20 dígitos.
    // Solo se descartan los de 20 dígitos o más (los IDs de Binance de 18-19 sí valen).
    if (/^\d{20,}$/.test(v)) return true;
    return false;
  };

  const pareceReferencia = (v) =>
    /\d/.test(v) && /^[A-Z0-9][A-Z0-9\-]{4,20}$/i.test(v) && !esInvalido(v);

  const limpiar = (v) => v.replace(/[^A-Z0-9]/gi, '');
  const BINANCE_REGEX = /\b\d{18}\b/;
  const BINANCE_REGEX_AMPLIO = /\b\d{17,19}\b/;

  // --- PASO 0: patrones bancarios inconfundibles primero ---
  const PATRONES_CONOCIDOS = /\b(?:JPM[A-Z0-9]{9}|WFCT[A-Z0-9]{8}|BOFA[A-Z0-9]{8,10})\b/i;
  const matchConocido = texto.match(PATRONES_CONOCIDOS);
  if (matchConocido) return limpiar(matchConocido[0]);

  // --- PASO 0b: comprobantes tipo Binance sin etiqueta visible de "ID DE ORDEN" ---
  const CONTEXTO_BINANCE = /CUENTA DE FONDOS|A[ÑN]ADIR ALIAS|USDT|BINANCE/i;
  if (CONTEXTO_BINANCE.test(texto)) {
    for (const l of lineas) {
      if (l.prohibida) continue;
      const m = l.texto.match(BINANCE_REGEX_AMPLIO);
      if (m) return m[0];
    }
  }

  // --- PASO 0c (NUEVO): BNC "Nota de operación ejecutada" ---
  // La referencia viene al final como "Ref.N°. 227837 Edo.Cta." (no al inicio
  // de línea), por eso la etiqueta normal no la detecta. El \s* también cubre
  // el caso en que el OCR deje el número en la línea siguiente.
  const BNC_REF_REGEX = /\bREF\.?\s*N\s*[°º˚oO0]?\s*\.?\s*:?\s*(\d{4,12})\b/i;
  const mBnc = texto.match(BNC_REF_REGEX);
  if (mBnc && !esInvalido(mBnc[1])) return mBnc[1];

  // --- PASO 0d (NUEVO): Bancamiga "Depósito Moneda Extranjera" ---
  // Referencia de 9-11 dígitos (ej. 1840139077). La marca de agua
  // "www.bancamiga.com" se superpone al número y el OCR mezcla líneas
  // (sobre todo si la foto está volteada), así que se limpia la marca de agua
  // y se busca el número cerca de la palabra "Referencia" y, si no, en todo el comprobante.
  const ES_BANCAMIGA = /BANCAMIG|MONEDA\s*EXTRANJERA|NOMBRE\s*DE\s*LA\s*TRANSACCI[OÓ]N/i;
  if (ES_BANCAMIGA.test(texto)) {
    // Líneas que nunca aportan referencia (la marca de agua NO cuenta aquí,
    // porque justamente ensucia la línea que sí tiene la referencia)
    const LINEA_NO_UTIL_BCM = /C\.?\s*I\.?\s*DEL\s*DEPOSIT|DEPOSITANTE|FIRMA|SELLO\s*DEL\s*CAJERO|CAJERO|^CAJA\b|TITULAR|TUBANCA|^0500\b|N[UÚ]MERO\s*DE\s*CUENTA/i;

    const limpiarMarcaAgua = (s) => s
      .replace(/(?:W{2,3}\.?)?\s*BANCAMIGA\s*\.?\s*(?:COM)?/gi, ' ')
      .replace(/\bW{3}\b\.?/gi, ' ')
      .replace(/(?<=\d)\.?COM\b/gi, ' ')
      .replace(/\.COM\b/gi, ' ');

    const validoBcm = (v) =>
      /^\d{9,11}$/.test(v) && !/^0500/.test(v) && !/^0172/.test(v) && !esInvalido(v);

    const buscarEnTexto = (s, permitirGrupos) => {
      const t = limpiarMarcaAgua(s);
      const tokens = t.match(/\b\d{9,11}\b/g) || [];
      for (const tk of tokens) if (validoBcm(tk)) return tk;
      if (permitirGrupos) {
        // El OCR a veces parte el número en 2-3 grupos (ej. "7020 781025")
        const grupos = t.match(/\d{2,8}(?:[ \-]\d{2,8}){1,2}/g) || [];
        for (const g of grupos) {
          const j = g.replace(/\D/g, '');
          if (validoBcm(j)) return j;
        }
      }
      return null;
    };

    // (a) Alrededor de la etiqueta "Referencia" (línea anterior, propia y 2 siguientes)
    for (let i = 0; i < lineasRaw.length; i++) {
      if (!/REFEREN/i.test(lineasRaw[i])) continue;
      if (/N[UÚ]MERO\s*DE\s*CUENTA/i.test(lineasRaw[i])) continue;
      const ventana = lineasRaw
        .slice(Math.max(0, i - 1), i + 3)
        .filter(l => !LINEA_NO_UTIL_BCM.test(l) || /REFEREN/i.test(l))
        .join(' ');
      const r = buscarEnTexto(ventana, true);
      if (r) return r;
    }

    // (b) Respaldo: cualquier línea útil del comprobante (por si el OCR
    // separó la etiqueta del número por la rotación)
    for (const l of lineasRaw) {
      if (LINEA_NO_UTIL_BCM.test(l)) continue;
      const r = buscarEnTexto(l, false);
      if (r) return r;
    }
    // Si no se encontró, se continúa con el PASO 1 y 2 de siempre
  }

  // --- PASO 1: búsqueda por etiqueta ---
  const VENTANA = 4;

  for (let i = 0; i < lineas.length; i++) {
    const l = lineas[i].texto;
    const m = l.match(etiquetaRegex);

    const esBinance = /^ID DE ORDEN/i.test(l);
    if (esBinance) {
      for (let j = i; j <= i + 2 && j < lineas.length; j++) {
        const match18 = lineas[j].texto.match(BINANCE_REGEX);
        if (match18) return match18[0];
      }
      const globalMatch18 = texto.match(BINANCE_REGEX);
      if (globalMatch18) return globalMatch18[0];
      continue;
    }

    if (!m) continue;

    // Si la propia línea de la etiqueta viene con señales de marca de agua
    // superpuesta, se descarta esta ocurrencia y se sigue buscando.
    const LINEA_CORRUPTA = /WWW\.|\.COM|@/i;
    if (LINEA_CORRUPTA.test(l)) continue;

    const mismaLinea = m[1].trim();

    if (mismaLinea) {
      const partes = mismaLinea.split(/\s+/);
      const candidato = partes[partes.length - 1];
      if (pareceReferencia(candidato)) return limpiar(candidato);
    }

    for (let j = i + 1; j <= i + VENTANA && j < lineas.length; j++) {
      if (lineas[j].prohibida) continue;

      const candidato = lineas[j].texto.trim();
      const primerElemento = candidato.split(/\s+/)[0];
      if (pareceReferencia(primerElemento) && !etiquetaRegex.test(candidato)) {
        return limpiar(primerElemento);
      }
    }
  }

  // --- PASO 2: fallback global (respetando líneas prohibidas) ---
  const idxReferencia = lineas.findIndex(l => /REFERENCIA/i.test(l.texto));
  const lineasDesdeReferencia = idxReferencia >= 0 ? lineas.slice(idxReferencia) : lineas;

  const lineasValidas = lineasDesdeReferencia.filter(l => !l.prohibida).map(l => l.texto).join('\n');
  const candidatosGlobales = lineasValidas.match(/\b[A-Z0-9]{6,20}\b/gi) || [];

  for (const c of candidatosGlobales) {
    if (!esInvalido(c) && /[A-Z]/i.test(c) && /[0-9]/.test(c)) {
      return limpiar(c);
    }
  }
  for (const c of candidatosGlobales) {
    if (!esInvalido(c) && /^[0-9]+$/.test(c)) {
      return limpiar(c);
    }
  }
  return 'No encontrada';
}
/* ============================================================================
 * FUNCIONES DE EXTRACCIÓN MONTO
 * ========================================================================== */
function extraerMonto(texto) {
  const textoLimpio = texto.replace(/\r?\n/g, ' ');

  // NUEVO: BNC "Nota de operación ejecutada". El monto viene en bolívares como
  // "VEB 263,040.00" (aparece dos veces en la nota). Solo se activa si el texto
  // es claramente de BNC, para no afectar a los demás comprobantes.
  const ES_BNC = /NOTA\s*DE\s*OPERACI[OÓ]N|BANCO\s*NACIONAL\s*DE\s*CR[EÉ]DITO|\bBNC\b/i;
  if (ES_BNC.test(textoLimpio)) {
    const matchBnc = textoLimpio.match(/\bVEB\s*:?\s*(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\b/i);
    if (matchBnc) return matchBnc[1].trim();
  }

  const patronAmericano = /\$\s?\d{1,3}(?:,\d{3})+(?:\.\d{2})?|\$\s?\d+\.\d{2}/;
  const matchAmerik = textoLimpio.match(patronAmericano);
  if (matchAmerik) return matchAmerik[0].trim();

  const patronesRespaldo = [
    /\$\s?\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})/i,
    /\b\d{1,3}(?:[.,]\d{3})*[.,]\d{2}\b/,
    /\b\d+\s?(?:USDT|USD|BS\.?)\b/i
  ];

  for (let patron of patronesRespaldo) {
    const match = textoLimpio.match(patron);
    if (match) return match[0].trim();
  }

  return 'No encontrado';
}
/* ============================================================================
 * FUNCIÓN DE EXTRACCIÓN DE MONTO PARA PAGOS EN EFECTIVO (billetes fotografiados)
 * ========================================================================== */

function extraerMontoEfectivo(texto, cantidadBilletes = 1) {
  if (!texto) return null;
  const t = texto.toUpperCase();
  let denominacion = null;

  // Detectar la denominación basada en el texto clásico impreso en los dólares
  if (t.includes('HUNDRED') || t.includes('100')) denominacion = 100;
  else if (t.includes('FIFTY') || t.includes('50')) denominacion = 50;
  else if (t.includes('TWENTY') || t.includes('20')) denominacion = 20;
  else if (t.includes('TEN') || /\b10\b/.test(t)) denominacion = 10;
  else if (t.includes('FIVE') || /\b5\b/.test(t)) denominacion = 5;
  else if (t.includes('TWO DOLLAR') || /\b2\b/.test(t)) denominacion = 2;
  else if (t.includes('ONE DOLLAR') || /\b1\b/.test(t)) denominacion = 1;

  // Si no se logra detectar nada legible, asumimos 0
  if (denominacion === null) return 0;

  // Si la función fue llamada desde procesarOCRFila con la cantidad de seriales detectados,
  // multiplicamos la denominación encontrada por la cantidad de billetes.
  if (cantidadBilletes && cantidadBilletes > 0) {
    return denominacion * cantidadBilletes;
  }

  // Si fue llamada desde contarBilletesIndividualmente (por banda),
  // solo devolvemos la denominación de ese billete específico.
  return denominacion;
}




/* ============================================================================
 * FUNCIONES DE EXTRACCIÓN BENEFICIARIO
 * ========================================================================== */

function extraerBeneficiario(texto) {
  const lineas = texto.split('\n').map(l => l.trim()).filter(Boolean);

  const limpiar = (valor) => valor.replace(/[.:]+$/, '').trim();

  // Etiquetas ordenadas por prioridad (nombre legal/registrado primero)
  const etiquetas = [
    /^Inscrito\(?a?\)?\s+como[\s:]*(.*)$/i,
    /^(?:BENEFICIARIO|Nombre\s+del?\s+titular)[\s:]*(.*)$/i,
    /^(?:TO|AS|ALIAS|ENVIADA\s+A|Enviado\s+a)[\s:]*(.*)$/i,
    /^A:\s*(.*)$/i, // "A:" con dos puntos, evita falsos positivos con la letra "A" suelta
  ];

  for (const patron of etiquetas) {
    for (let i = 0; i < lineas.length; i++) {
      const m = lineas[i].match(patron);
      if (m) {
        const valor = (m[1] || '').trim();
        if (valor.length > 2) return limpiar(valor);

        // Si la etiqueta está sola en su línea, el valor puede venir en la siguiente
        if (lineas[i + 1] && lineas[i + 1].length > 2) {
          return limpiar(lineas[i + 1]);
        }
      }
    }
  }

  // Frases embebidas: "Enviaste $X a Y", "Dinero enviado a Y", "sent to Y"
  const patronesTexto = [
    /Enviaste\s+\$?[\d.,]+\s+a\s+([^\n.]+?)\.?\s*$/im,
    /Dinero enviado a\s+([^\n]+)/i,
    /(?:sent to|payment of.*?to)\s+([^\n]+)/i,
  ];

  for (const patron of patronesTexto) {
    const m = texto.match(patron);
    if (m && m[1]) return limpiar(m[1]);
  }

  return 'No encontrado';
}
/* ============================================================================
 * FUNCIÓN DE PRUEBA MANUAL
 * ========================================================================== */

function pruebaManualPorFila() {
  const FILA_DE_PRUEBA = 825; // ajusta la fila a probar
  const hoja = SpreadsheetApp.openById(SHEET_ID).getSheetByName(NOMBRE_HOJA);
  procesarOCRFila(hoja, FILA_DE_PRUEBA);
}


/**========================================================================== 
 * NUEVO: punto de entrada para el frontend alojado en Vercel.
 * Recibe { fn, args, pin } y ejecuta SOLO funciones de una lista blanca.
 ========================================================================== 
 */
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = req.fn;
    const args = req.args || [];
    let data;
 
    switch (fn) {
      case 'procesarFormulario':
        data = procesarFormulario(args[0]);
        break;
      case 'obtenerVendedores':
        data = obtenerVendedores();
        break;
      case 'validarPin':
        data = pinValido_(args[0]);
        break;
      case 'obtenerPagos':
        if (!pinValido_(req.pin)) throw new Error('UNAUTHORIZED');
        data = obtenerPagos();
        break;
      default:
        throw new Error('Función no permitida: ' + fn);
    }
    return jsonOut_({ ok: true, data: data });
  } catch (err) {
    console.error('❌ Error en doPost: ' + err.message);
    return jsonOut_({ ok: false, error: err.message });
  }
}
 
/** NUEVO: respuesta JSON estándar */
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
 
/**
 * NUEVO: valida la clave del panel en el SERVIDOR (ya no queda visible en el código
 * público). Para cambiarla: Configuración del proyecto > Propiedades del script >
 * agregar ADMIN_PIN. Si no existe, se usa la clave actual 'BNH2025'.
 */
function pinValido_(pin) {
  const esperado = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || 'BNH2025';
  return String(pin || '').trim() === esperado;
}





/* ============================================================================
 * APRENDIZAJE INTELIGENTE DEL OCR
 * Hojas: OCR_FALLOS (casos no leídos / correcciones) y OCR_REGLAS (reglas aprendidas)
 * ========================================================================== */

const HOJA_OCR_FALLOS = 'OCR_FALLOS';
const HOJA_OCR_REGLAS = 'OCR_REGLAS';
const CAMPOS_APRENDIZAJE = ['REFERENCIA', 'MONTO', 'BENEFICIARIO'];
const CLAVE_CACHE_REGLAS = 'OCR_REGLAS_V1';
const CACHE_REGLAS_SEG = 600;
const MAX_TEXTO_FALLO = 45000;

// IA (opcional): la API key de Gemini se lee de la hoja OCR_REGLAS, celda A1.
// Modelo por defecto editable con la propiedad de script GEMINI_MODEL.
const IA_ACTIVADA = true;
const IA_MODELO_DEFECTO = 'gemini-3.8-flash';
const CELDA_API_KEY = 'K1';
// false = las reglas propuestas por la IA se guardan con ACTIVA = NO
// y tú las apruebas cambiando la celda a SI.
const ACTIVAR_REGLAS_IA_AUTOMATICO = true;

const ENC_FALLOS = ['FECHA', 'FILA PAGOS', 'CAMPO', 'HUELLA', 'COMPROBANTE', 'TEXTO OCR', 'VALOR CORRECTO', 'ESTADO', 'NOTAS'];
const ENC_REGLAS = ['CAMPO', 'HUELLA (regex)', 'REGEX DE EXTRACCION', 'ACTIVA', 'ORIGEN', 'FECHA', 'ACIERTOS', 'EJEMPLO'];

// Pistas para identificar el "tipo de comprobante" (de lo más a lo menos específico)
const HUELLAS_COMPROBANTE = [
  'NOTA\\s*DE\\s*OPERACI', 'BANCO\\s*NACIONAL\\s*DE\\s*CR', '\\bBNC\\b', 'BANCAMIGA', 'BANESCO',
  'MERCANTIL', 'PROVINCIAL', 'BBVA', 'BANCO\\s*DE\\s*VENEZUELA', 'BANPLUS', 'BANCO\\s*EXTERIOR',
  'BICENTENARIO', 'BANCO\\s*DEL\\s*TESORO', 'ZELLE', 'BINANCE', 'BANK\\s*OF\\s*AMERICA',
  'CHASE', 'WELLS\\s*FARGO', 'PAGO\\s*M[OÓ]VIL'
];

/* ------------------------------- INSTALACIÓN ------------------------------- */

/** Ejecutar UNA vez desde el editor: crea las hojas y el disparador horario. */
function instalarAprendizaje() {
  hojaAprendizaje_(HOJA_OCR_FALLOS, ENC_FALLOS);
  hojaAprendizaje_(HOJA_OCR_REGLAS, ENC_REGLAS);
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'aprenderDeCorrecciones')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('aprenderDeCorrecciones').timeBased().everyHours(1).create();
  console.log('✅ Aprendizaje instalado: hojas creadas y disparador horario activo.');
}

function hojaAprendizaje_(nombre, encabezados) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.getRange(1, 1, 1, encabezados.length).setValues([encabezados])
      .setFontWeight('bold').setBackground('#0b3d91').setFontColor('#ffffff');
    h.setFrozenRows(1);
  }
  return h;
}

/* ------------------------- ORQUESTADOR (lo llama el OCR) ------------------------- */

/**
 * Completa los campos que el OCR normal no pudo leer.
 * Orden: reglas aprendidas -> IA (si hay API key) -> registro para corrección humana.
 * @param {Sheet} hoja Hoja PAGOS
 * @param {number} fila Fila procesada
 * @param {string} texto Texto OCR completo
 * @param {{referencia:*, monto:*, beneficiario:*}} valores Lo que leyó el OCR normal
 * @param {{referencia?:boolean, monto?:boolean, beneficiario?:boolean}} [omitir] Campos a no tocar
 * @return {{referencia:*, monto:*, beneficiario:*}}
 */
function aprenderYCompletar_(hoja, fila, texto, valores, omitir) {
  const mapa = { REFERENCIA: 'referencia', MONTO: 'monto', BENEFICIARIO: 'beneficiario' };
  const omit = omitir || {};
  const resultado = { referencia: valores.referencia, monto: valores.monto, beneficiario: valores.beneficiario };
  const faltantes = () => CAMPOS_APRENDIZAJE.filter(c => !omit[mapa[c]] && esFallo_(c, resultado[mapa[c]]));

  let pendientes = faltantes();
  if (pendientes.length === 0) return resultado;
  if (!texto || /^(ERROR OCR|NO SE DETECT[OÓ] TEXTO)$/i.test(String(texto).trim())) return resultado;

  const url = String(hoja.getRange(fila, COL_ARCHIVO).getValue() || '');

  // 1) Reglas aprendidas
  pendientes.forEach(c => {
    const v = aplicarReglasAprendidas_(c, texto);
    if (v) {
      resultado[mapa[c]] = v;
      console.log('🧠 ' + c + ' resuelto por regla aprendida: ' + v);
    }
  });
  pendientes = faltantes();
  if (pendientes.length === 0) return resultado;

  // 2) IA (opcional)
  const deIA = consultarIA_(texto, pendientes);
  pendientes.forEach(c => {
    const v = deIA[c];
    if (!v) return;
    resultado[mapa[c]] = v;
    const g = generarRegla_(c, texto, v);
    if (g.ok) {
      registrarRegla_(c, g.huella, g.regex, 'IA', ACTIVAR_REGLAS_IA_AUTOMATICO ? 'true' : 'false', v);
    }
    registrarFallo_(fila, c, texto, url, 'RESUELTO_IA', v,
      g.ok ? 'Regla propuesta en OCR_REGLAS' : ('Sin regla: ' + g.motivo));
    console.log('🤖 ' + c + ' resuelto por IA: ' + v);
  });
  pendientes = faltantes();

  // 3) Lo que sigue sin leerse queda para corrección humana
  pendientes.forEach(c => registrarFallo_(fila, c, texto, url, 'PENDIENTE', '', ''));
  return resultado;
}

/* ----------------------------- DETECCIÓN DE FALLOS ----------------------------- */

function esFallo_(campo, valor) {
  const v = String(valor === null || valor === undefined ? '' : valor).trim().toUpperCase();
  if (!v) return true;
  if (campo === 'REFERENCIA') return VALORES_REFERENCIA_IGNORADOS.has(v);
  if (campo === 'MONTO') {
    return ['NO ENCONTRADO', 'NO ENCONTRADA', 'NO DETECTADO', 'ERROR AL LEER MONTO', '0'].indexOf(v) !== -1;
  }
  return ['NO ENCONTRADO', 'NO ENCONTRADA', 'NO DETECTADO'].indexOf(v) !== -1;
}

function registrarFallo_(fila, campo, texto, url, estado, valor, notas) {
  const h = hojaAprendizaje_(HOJA_OCR_FALLOS, ENC_FALLOS);
  h.appendRow([
    new Date(), fila, campo, calcularHuella_(texto), textoSeguro_(url),
    textoSeguro_(String(texto).substring(0, MAX_TEXTO_FALLO)),
    textoSeguro_(valor || ''), estado, textoSeguro_(notas || '')
  ]);
}

/** Evita que Sheets interprete como fórmula un texto que empiece con = + - @ */
function textoSeguro_(s) {
  const t = String(s === null || s === undefined ? '' : s);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function calcularHuella_(texto) {
  for (const patron of HUELLAS_COMPROBANTE) {
    if (new RegExp(patron, 'i').test(texto)) return patron;
  }
  return '';
}

/* ------------------------------ REGLAS APRENDIDAS ------------------------------ */

function cargarReglas_() {
  const cache = CacheService.getScriptCache();
  const enCache = cache.get(CLAVE_CACHE_REGLAS);
  if (enCache) {
    try { return JSON.parse(enCache); } catch (e) { /* se recarga abajo */ }
  }
  const h = hojaAprendizaje_(HOJA_OCR_REGLAS, ENC_REGLAS);
  const ultima = h.getLastRow();
  let reglas = [];
  if (ultima >= 2) {
    reglas = h.getRange(2, 1, ultima - 1, ENC_REGLAS.length).getValues()
      .map((r, i) => ({
        fila: i + 2,
        campo: String(r[0]).trim().toUpperCase(),
        huella: String(r[1] || '').trim(),
        regex: String(r[2] || '').trim(),
        activa: String(r[3]).trim().toUpperCase() === 'SI'
      }))
      .filter(r => r.activa && r.regex && CAMPOS_APRENDIZAJE.indexOf(r.campo) !== -1);
  }
  cache.put(CLAVE_CACHE_REGLAS, JSON.stringify(reglas), CACHE_REGLAS_SEG);
  return reglas;
}

function aplicarReglasAprendidas_(campo, texto) {
  const reglas = cargarReglas_().filter(r => r.campo === campo);
  for (const r of reglas) {
    try {
      if (r.huella && !new RegExp(r.huella, 'i').test(texto)) continue;
      const m = new RegExp(r.regex, 'i').exec(texto);
      if (!m || !m[1]) continue;
      const valor = limpiarValorAprendido_(campo, m[1]);
      if (!valor || esFallo_(campo, valor)) continue;
      sumarAcierto_(r.fila);
      return valor;
    } catch (err) {
      console.error('⚠️ Regla inválida en fila ' + r.fila + ' de ' + HOJA_OCR_REGLAS + ': ' + err.message);
    }
  }
  return null;
}

function limpiarValorAprendido_(campo, crudo) {
  const v = String(crudo || '').trim();
  if (campo === 'REFERENCIA') {
    const limpio = v.replace(/[^A-Z0-9]/gi, '').toUpperCase();
    return (limpio.length >= 4 && limpio.length <= 24 && /\d/.test(limpio)) ? limpio : null;
  }
  if (campo === 'MONTO') return /\d/.test(v) ? v : null;
  const nombre = v.replace(/\s+/g, ' ').replace(/[.:,\-\s]+$/, '').trim();
  return nombre.length >= 3 ? nombre : null;
}

function sumarAcierto_(filaRegla) {
  try {
    const celda = hojaAprendizaje_(HOJA_OCR_REGLAS, ENC_REGLAS).getRange(filaRegla, 7);
    celda.setValue((Number(celda.getValue()) || 0) + 1);
  } catch (e) { /* el contador es informativo, nunca debe interrumpir */ }
}

function registrarRegla_(campo, huella, regex, origen, activa, ejemplo) {
  const h = hojaAprendizaje_(HOJA_OCR_REGLAS, ENC_REGLAS);
  const ultima = h.getLastRow();
  if (ultima >= 2) {
    const existentes = h.getRange(2, 1, ultima - 1, 3).getValues();
    const repetida = existentes.some(r =>
      String(r[0]).trim().toUpperCase() === campo && String(r[1]) === huella && String(r[2]) === regex);
    if (repetida) return false;
  }
  h.appendRow([campo, huella, regex, activa, origen, new Date(), 0, textoSeguro_(ejemplo)]);
  CacheService.getScriptCache().remove(CLAVE_CACHE_REGLAS);
  return true;
}

/* ------------------- GENERACIÓN DE REGLAS (el "aprendizaje") ------------------- */

/**
 * A partir de un texto OCR y el valor correcto, deduce una regla
 * "ETIQUETA + valor" y la VALIDA contra ese mismo texto antes de aceptarla.
 * @return {{ok:boolean, regex?:string, huella?:string, motivo?:string}}
 */
function generarRegla_(campo, texto, valorCorrecto) {
  const valor = String(valorCorrecto || '').trim();
  if (!valor || !texto) return { ok: false, motivo: 'Falta el valor o el texto OCR.' };

  const loc = localizarValor_(campo, texto, valor);
  if (!loc) {
    return { ok: false, motivo: 'El valor correcto no aparece en el texto que leyó el OCR (la imagen no se leyó bien).' };
  }

  const inicioLinea = loc.index === 0 ? 0 : texto.lastIndexOf('\n', loc.index - 1) + 1;
  const antes = texto.substring(inicioLinea, loc.index);
  const patronValor = patronDeValor_(campo, valor);
  const prefijoMoneda = campo === 'MONTO' ? '(?:US\\$|\\$|USD|VEB|VES|BS\\.?)?[ \\t]*' : '';
  const candidatos = [];

  const etiquetaMisma = extraerEtiqueta_(antes);
  if (etiquetaMisma) {
    candidatos.push(etiquetaMisma + '[ \\t]*[:#\\-]?[ \\t]*' + prefijoMoneda + '(' + patronValor + ')');
  }

  const previas = texto.substring(0, Math.max(inicioLinea - 1, 0))
    .split('\n').map(l => l.trim()).filter(Boolean);
  const etiquetaPrevia = previas.length ? extraerEtiqueta_(previas[previas.length - 1]) : '';
  if (etiquetaPrevia) {
    candidatos.push(etiquetaPrevia + '[ \\t]*:?[ \\t]*\\n[ \\t]*' + prefijoMoneda + '(' + patronValor + ')');
  }

  for (const regex of candidatos) {
    try {
      const m = new RegExp(regex, 'i').exec(texto);
      if (m && m[1] && sonMismoValor_(campo, m[1], valor)) {
        return { ok: true, regex: regex, huella: calcularHuella_(texto) };
      }
    } catch (e) { /* candidato inválido: se prueba el siguiente */ }
  }
  return { ok: false, motivo: 'No se encontró una etiqueta estable junto al valor para generalizar la regla.' };
}

function localizarValor_(campo, texto, valor) {
  const directo = texto.toUpperCase().indexOf(valor.toUpperCase());
  if (directo !== -1) return { index: directo, length: valor.length };

  let patron;
  if (campo === 'MONTO') {
    const digitos = valor.replace(/\D/g, '');
    if (!digitos) return null;
    patron = digitos.split('').join('[.,\\s]?');
  } else if (campo === 'REFERENCIA') {
    const limpio = valor.replace(/[^A-Z0-9]/gi, '');
    if (limpio.length < 4) return null;
    patron = limpio.split('').join('[\\s\\-]?');
  } else {
    patron = valor.trim().split(/\s+/).map(escaparRegex_).join('\\s+');
  }
  const m = new RegExp(patron, 'i').exec(texto);
  return m ? { index: m.index, length: m[0].length } : null;
}

function patronDeValor_(campo, valor) {
  if (campo === 'MONTO') {
    return '\\d{1,3}(?:[.,]\\d{3})*(?:[.,]\\d{1,2})?|\\d+(?:[.,]\\d{1,2})?';
  }
  if (campo === 'REFERENCIA') {
    const limpio = valor.replace(/[^A-Z0-9]/gi, '');
    const min = Math.max(4, limpio.length - 2);
    const max = Math.min(24, limpio.length + 3);
    return (/^\d+$/.test(limpio) ? '\\d' : '[A-Z0-9]') + '{' + min + ',' + max + '}';
  }
  return "[A-ZÁÉÍÓÚÑÜ][A-ZÁÉÍÓÚÑÜ.,&' \\t-]{2,80}";
}

/** Toma las últimas palabras antes del valor como "etiqueta" del campo. */
function extraerEtiqueta_(antes) {
  let e = String(antes || '').replace(/[:\-#\s$]+$/, '').trim();
  if (!/[A-Za-zÁÉÍÓÚáéíóúÑñ]{3,}/.test(e)) return '';
  const palabras = e.split(/\s+/).slice(-4);
  while (palabras.join(' ').length > 40 && palabras.length > 1) palabras.shift();
  e = palabras.join(' ');
  return escaparRegex_(e).replace(/\s+/g, '[ \\t]+');
}

function escaparRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\\/\-]/g, '\\$&');
}

function montoANumero_(s) {
  const t = String(s).replace(/[^\d.,]/g, '');
  if (!t) return NaN;
  const ultPunto = t.lastIndexOf('.');
  const ultComa = t.lastIndexOf(',');
  const sepDec = Math.max(ultPunto, ultComa);
  if (sepDec === -1) return Number(t);
  const decimales = t.length - sepDec - 1;
  const hayAmbos = ultPunto !== -1 && ultComa !== -1;
  if (hayAmbos || decimales <= 2) {
    return Number(t.slice(0, sepDec).replace(/[.,]/g, '') + '.' + t.slice(sepDec + 1));
  }
  return Number(t.replace(/[.,]/g, ''));
}

function sonMismoValor_(campo, a, b) {
  if (campo === 'MONTO') {
    const na = montoANumero_(a), nb = montoANumero_(b);
    return !isNaN(na) && !isNaN(nb) && Math.abs(na - nb) < 0.005;
  }
  if (campo === 'REFERENCIA') {
    return String(a).replace(/[^A-Z0-9]/gi, '').toUpperCase() === String(b).replace(/[^A-Z0-9]/gi, '').toUpperCase();
  }
  const norm = s => String(s).replace(/\s+/g, ' ').replace(/[.:,\-\s]+$/, '').trim().toUpperCase();
  return norm(a) === norm(b);
}

/* -------------------- APRENDER DE TUS CORRECCIONES (disparador) -------------------- */

/**
 * Procesa las filas de OCR_FALLOS con ESTADO = PENDIENTE y VALOR CORRECTO escrito:
 * corrige la celda en PAGOS y genera/valida la regla. Corre cada hora o a mano.
 */
function aprenderDeCorrecciones() {
  const h = hojaAprendizaje_(HOJA_OCR_FALLOS, ENC_FALLOS);
  const ultima = h.getLastRow();
  if (ultima < 2) return;

  const hojaPagos = SpreadsheetApp.openById(SHEET_ID).getSheetByName(NOMBRE_HOJA);
  const datos = h.getRange(2, 1, ultima - 1, ENC_FALLOS.length).getValues();
  const colPorCampo = { REFERENCIA: COL_REFERENCIA, MONTO: COL_MONTO, BENEFICIARIO: COL_BENEFICIARIO };
  let aprendidas = 0;

  datos.forEach((r, i) => {
    const estado = String(r[7]).trim().toUpperCase();
    const corregido = String(r[6]).trim();
    if (estado !== 'PENDIENTE' || !corregido) return;

    const filaFallos = i + 2;
    const campo = String(r[2]).trim().toUpperCase();
    if (!colPorCampo[campo]) return;

    const notas = [];

    // 1) Corregir la fila en PAGOS (se ubica por URL del comprobante, no por nº de fila,
    //    porque los duplicados borrados desplazan las filas).
    const filaPagos = buscarFilaPorUrl_(hojaPagos, String(r[4]));
    if (filaPagos) {
      hojaPagos.getRange(filaPagos, colPorCampo[campo]).setValue(corregido);
      notas.push('PAGOS fila ' + filaPagos + ' corregida.');
    } else {
      notas.push('No se halló el comprobante en PAGOS (¿fila eliminada?).');
    }

    // 2) Aprender la regla
    const g = generarRegla_(campo, String(r[5]), corregido);
    let nuevoEstado = 'NO_APRENDIDO';
    if (g.ok) {
      const creada = registrarRegla_(campo, g.huella, g.regex, 'HUMANO', 'SI', corregido);
      nuevoEstado = 'APRENDIDO';
      notas.push(creada ? 'Regla nueva activa.' : 'La regla ya existía.');
      aprendidas++;
    } else {
      notas.push(g.motivo);
    }
    h.getRange(filaFallos, 8, 1, 2).setValues([[nuevoEstado, notas.join(' ')]]);
  });

  if (aprendidas > 0) CacheService.getScriptCache().remove(CLAVE_CACHE_REGLAS);
  console.log('🧠 Reglas nuevas aprendidas: ' + aprendidas);
}

function buscarFilaPorUrl_(hoja, url) {
  if (!url) return 0;
  const ultima = hoja.getLastRow();
  if (ultima < 2) return 0;
  const hallada = hoja.getRange(2, COL_ARCHIVO, ultima - 1, 1)
    .createTextFinder(url).matchEntireCell(true).findNext();
  return hallada ? hallada.getRow() : 0;
}

/* ------------------------------ IA (Gemini) ------------------------------ */

/** Lee la API key de OCR_REGLAS!A1. Devuelve '' si no hay una key válida. */
function leerApiKeyIA_() {
  try {
    const hoja = SpreadsheetApp.openById(SHEET_ID).getSheetByName(HOJA_OCR_REGLAS);
    if (!hoja) return '';
    const valor = String(hoja.getRange(CELDA_API_KEY).getValue() || '').trim();
    // Si A1 todavía tiene el encabezado original, no es una key
    return (valor && valor.toUpperCase() !== 'CAMPO') ? valor : '';
  } catch (e) {
    console.error('⚠️ No se pudo leer la API key en ' + HOJA_OCR_REGLAS + '!' + CELDA_API_KEY + ': ' + e.message);
    return '';
  }
}

/**
 * Pide a Gemini que lea los campos faltantes del texto OCR.
 * Cada valor devuelto se VERIFICA contra el texto original; si no aparece, se descarta.
 * @return {Object} p. ej. { REFERENCIA: '227837', MONTO: '263,040.00' }
 */
function consultarIA_(texto, campos) {
  const apiKey = leerApiKeyIA_();
  if (!IA_ACTIVADA || !apiKey || !campos.length) return {};

  const modelo = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || IA_MODELO_DEFECTO;
  const instruccion =
    'Este es el texto OCR de un comprobante de pago (puede venir desordenado, volteado o con marcas de agua).\n' +
    'Extrae SOLO estos campos: ' + campos.join(', ') + '.\n' +
    'Reglas: copia cada valor EXACTAMENTE como aparece en el texto; si no estás seguro, usa null; no inventes.\n' +
    '- REFERENCIA: número de referencia, confirmación u operación (NO cuentas de 20 dígitos, teléfonos, cédulas ni RIF).\n' +
    '- MONTO: importe de la operación, con su moneda o símbolo si aparece.\n' +
    '- BENEFICIARIO: persona o empresa que recibe el pago.\n' +
    'Responde SOLO un JSON con las claves en mayúsculas, sin explicaciones.\n\n' +
    'TEXTO OCR:\n' + String(texto).substring(0, 12000);

  try {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
      encodeURIComponent(modelo) + ':generateContent';
    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify({
        contents: [{ role: 'user', parts: [{ text: instruccion }] }],
        generationConfig: {
          temperature: 0,
          maxOutputTokens: 300,
          responseMimeType: 'application/json'
        }
      }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      console.error('⚠️ IA respondió ' + resp.getResponseCode() + ': ' + resp.getContentText().substring(0, 300));
      return {};
    }
    const cand = (JSON.parse(resp.getContentText()).candidates || [])[0];
    const salida = ((cand && cand.content && cand.content.parts) || [])
      .map(p => p.text || '').join('');
    const bruto = salida.match(/\{[\s\S]*\}/);
    if (!bruto) return {};
    const json = JSON.parse(bruto[0]);

    const verificados = {};
    campos.forEach(c => {
      const propuesto = json[c];
      if (propuesto === null || propuesto === undefined || String(propuesto).trim() === '') return;
      const loc = localizarValor_(c, texto, String(propuesto).trim());
      if (!loc) { console.log('⚠️ IA propuso ' + c + ' pero no aparece en el texto. Descartado.'); return; }
      let valor = texto.substr(loc.index, loc.length).trim();
      if (c === 'REFERENCIA') valor = valor.replace(/[^A-Z0-9]/gi, '').toUpperCase();
      if (!esFallo_(c, valor)) verificados[c] = valor;
    });
    return verificados;
  } catch (err) {
    console.error('⚠️ Error consultando la IA (fail-open): ' + err.message);
    return {};
  }
}
