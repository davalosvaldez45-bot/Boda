/**
 * ============================================================
 *  RSVP Boda Gustavo & Fabiola — Backend en Google Apps Script
 * ============================================================
 *
 * Este script es el ÚNICO intermediario entre la invitación (HTML)
 * y la planilla de Google Sheets, que permanece PRIVADA en todo momento.
 * Los invitados nunca ven, abren ni descargan la planilla: sólo reciben
 * las respuestas JSON que este script decide devolver.
 *
 * ---------------------------------------------------------------
 * ESTRUCTURA DE LA HOJA (pestaña llamada exactamente "Invitados")
 * ---------------------------------------------------------------
 * Fila 1 = encabezados. A partir de la fila 2, una fila por invitación:
 *
 *   A: Codigo               (ej: FG001)  -> identificador único, NO usar nombres
 *   B: Nombre                (ej: "Familia Pérez")
 *   C: CantidadPermitida     (ej: 4)
 *   D: Confirmado            (TRUE/FALSE, se completa solo)
 *   E: CantidadConfirmada    (número, se completa solo)
 *   F: FechaConfirmacion     (fecha/hora, se completa solo)
 *   G: Asistencia            (Sí/No, se completa solo, sólo referencia interna)
 *
 * Antes de enviar las invitaciones, cargá manualmente las columnas A, B y C
 * para tus 100 invitados. Dejá D, E, F y G vacías: el script las completa.
 *
 * ---------------------------------------------------------------
 * CÓMO DESPLEGAR
 * ---------------------------------------------------------------
 * 1) Abrí tu Google Sheet privado (el que NUNCA vas a compartir).
 * 2) Extensiones > Apps Script.
 * 3) Borrá el contenido de Codigo.gs y pegá todo este archivo.
 * 4) Arriba a la derecha: Implementar > Nueva implementación.
 *    - Tipo: "Aplicación web"
 *    - Ejecutar como: "Yo" (tu cuenta)
 *    - Quién tiene acceso: "Cualquier usuario" (Anyone)
 *      (esto NO da acceso a la planilla, sólo permite que la invitación
 *       llame a esta API; la planilla en sí sigue siendo privada)
 * 5) Copiá la URL que termina en /exec y pegala en index.html en la
 *    constante APPS_SCRIPT_URL.
 * 6) Cada vez que modifiques este código, tenés que crear una
 *    "Nueva implementación" (o gestionar implementaciones) para que
 *    los cambios se apliquen a la URL publicada.
 *
 * ---------------------------------------------------------------
 * SI SIENTE LENTO (checklist de rendimiento)
 * ---------------------------------------------------------------
 * - Usá la URL que termina en /exec, NUNCA la de /dev (la de pruebas es
 *   mucho más lenta y requiere que el invitado esté logueado como vos).
 * - "Quién tiene acceso" debe ser "Cualquier usuario" (Anyone), NO
 *   "Cualquier usuario con cuenta de Google": esa segunda opción exige
 *   iniciar sesión antes de responder y agrega varios segundos (o rompe
 *   la respuesta JSON directamente).
 * - La primera consulta después de un rato sin uso siempre tarda un poco
 *   más (Google "despierta" el script en frío); las siguientes son mucho
 *   más rápidas gracias a la caché de 5 minutos que ya incluye este código.
 * ---------------------------------------------------------------
 */

const SHEET_NAME = 'Invitados';
const CACHE_PREFIX = 'guest_';
const CACHE_TTL_SECONDS = 300; // 5 minutos

// Columnas (1-indexed, tal como las usa getRange)
const COL_CODIGO = 1;
const COL_NOMBRE = 2;
const COL_CANTIDAD_PERMITIDA = 3;
const COL_CONFIRMADO = 4;
const COL_CANTIDAD_CONFIRMADA = 5;
const COL_FECHA = 6;
const COL_ASISTENCIA = 7;

function getCache_() {
  return CacheService.getScriptCache();
}

function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('No existe la hoja "' + SHEET_NAME + '"');
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Busca una fila por código. Devuelve null si no existe.
function findRowByCode_(sheet, codigo) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, COL_ASISTENCIA).getValues();
  for (let i = 0; i < data.length; i++) {
    const cell = String(data[i][COL_CODIGO - 1] || '').trim().toUpperCase();
    if (cell === codigo) {
      return { rowIndex: i + 2, row: data[i] }; // +2 porque data empieza en la fila 2
    }
  }
  return null;
}

function isConfirmadoValue_(value) {
  return value === true || String(value).trim().toUpperCase() === 'TRUE';
}

// Payload público: SOLO estos campos salen del script, nunca la fila completa.
function buildGuestPayload_(row) {
  return {
    success: true,
    codigo: String(row[COL_CODIGO - 1]).trim().toUpperCase(),
    nombre: row[COL_NOMBRE - 1],
    cantidadPermitida: Number(row[COL_CANTIDAD_PERMITIDA - 1]) || 1,
    confirmado: isConfirmadoValue_(row[COL_CONFIRMADO - 1]),
    cantidadConfirmada: Number(row[COL_CANTIDAD_CONFIRMADA - 1]) || 0
  };
}

/**
 * GET — Consulta de estado (se llama UNA vez cuando el invitado abre la invitación).
 * Uso: {APPS_SCRIPT_URL}?codigo=FG001
 *
 * Se resuelve primero contra la caché (CacheService, 5 min): si dos personas de
 * la misma familia abren el mismo link casi al mismo tiempo, o alguien recarga
 * la página varias veces, sólo la primera consulta lee la planilla; el resto
 * responde al instante desde caché, sin abrir el Sheet de nuevo.
 */
function doGet(e) {
  try {
    const codigoRaw = (e && e.parameter && e.parameter.codigo) || '';
    const codigo = codigoRaw.trim().toUpperCase();

    if (!codigo) {
      return jsonOut_({ success: false, error: 'missing_code' });
    }

    const cache = getCache_();
    const cacheKey = CACHE_PREFIX + codigo;
    const cached = cache.get(cacheKey);
    if (cached) {
      return jsonOut_(JSON.parse(cached));
    }

    const sheet = getSheet_();
    const found = findRowByCode_(sheet, codigo);
    if (!found) {
      return jsonOut_({ success: false, error: 'not_found' });
    }

    const payload = buildGuestPayload_(found.row);
    cache.put(cacheKey, JSON.stringify(payload), CACHE_TTL_SECONDS);
    return jsonOut_(payload);
  } catch (err) {
    return jsonOut_({ success: false, error: 'server_error' });
  }
}

/**
 * POST — Confirmación de asistencia (se llama cuando el invitado envía el formulario).
 * Body esperado (JSON, como texto plano para evitar preflight CORS):
 *   { "codigo": "FG001", "asistencia": "Sí" | "No", "cantidad": "3" }
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // Espera hasta 6s si otro invitado está escribiendo al mismo tiempo
    // (importante con ~100 invitados accediendo el mismo día). No hace falta
    // más: la escritura en sí es una operación muy corta.
    lock.waitLock(6000);

    let body;
    try {
      body = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonOut_({ success: false, error: 'bad_request' });
    }

    const codigo = String(body.codigo || '').trim().toUpperCase();
    if (!codigo) {
      return jsonOut_({ success: false, error: 'missing_code' });
    }

    const asistencia = body.asistencia === 'Sí' ? 'Sí' : 'No';
    let cantidad = parseInt(body.cantidad, 10);
    if (isNaN(cantidad) || cantidad < 0) cantidad = 0;

    const sheet = getSheet_();
    const found = findRowByCode_(sheet, codigo);
    if (!found) {
      return jsonOut_({ success: false, error: 'not_found' });
    }

    const cantidadPermitida = Number(found.row[COL_CANTIDAD_PERMITIDA - 1]) || 1;
    // Nunca permite confirmar más personas de las asignadas a ese código.
    const cantidadFinal = asistencia === 'Sí'
      ? Math.min(cantidad || 1, cantidadPermitida)
      : 0;

    sheet.getRange(found.rowIndex, COL_CONFIRMADO).setValue(true);
    sheet.getRange(found.rowIndex, COL_CANTIDAD_CONFIRMADA).setValue(cantidadFinal);
    sheet.getRange(found.rowIndex, COL_FECHA).setValue(new Date());
    sheet.getRange(found.rowIndex, COL_ASISTENCIA).setValue(asistencia);

    const payload = {
      success: true,
      codigo: codigo,
      nombre: found.row[COL_NOMBRE - 1],
      cantidadPermitida: cantidadPermitida,
      confirmado: true,
      cantidadConfirmada: cantidadFinal
    };

    // Refresca la caché con el nuevo estado para que cualquier consulta GET
    // posterior (por ejemplo, si la persona recarga la página) sea instantánea
    // y ya muestre la confirmación, sin depender de que expire la caché vieja.
    getCache_().put(CACHE_PREFIX + codigo, JSON.stringify(payload), CACHE_TTL_SECONDS);

    return jsonOut_(payload);

  } catch (err) {
    return jsonOut_({ success: false, error: 'server_error' });
  } finally {
    lock.releaseLock();
  }
}
