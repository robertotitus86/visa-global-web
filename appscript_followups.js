// ═══════════════════════════════════════════════════════════════════
// GOOGLE APPS SCRIPT — Ejecutor de Follow-ups Automáticos
// Se ejecuta cada 30 minutos vía trigger de tiempo
// Revisa la hoja "Follow-ups" y envía los mensajes pendientes
// ═══════════════════════════════════════════════════════════════════

const BOT_URL      = 'https://visa-global-bot.onrender.com';
const SS_ID_FU     = '19yHZ5HJH5eWyFXej8ffGBT2_sttXDZtvNaCoNEzjIOU';
const HOJA_FOLLOWUP = 'Follow-ups';

const COLS = {
  TELEFONO:        1,
  PHONE_NUMBER_ID: 2,
  TIPO:            3,
  MENSAJE:         4,
  FECHA_ENVIO:     5,
  ESTADO:          6,
  CONTEXTO:        7,
};

// ── Trigger principal — corre cada 30 minutos ──────────────────────
function ejecutarFollowupsPendientes() {
  const ss    = SpreadsheetApp.openById(SS_ID_FU);
  const sheet = ss.getSheetByName(HOJA_FOLLOWUP);
  if (!sheet || sheet.getLastRow() < 2) return;

  const data  = sheet.getDataRange().getValues();
  const ahora = new Date();

  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const estado = String(row[COLS.ESTADO - 1]).trim();
    if (estado !== 'PENDIENTE') continue;

    const fechaEnvio = new Date(row[COLS.FECHA_ENVIO - 1]);
    if (fechaEnvio > ahora) continue; // Aún no es el momento

    const telefono      = String(row[COLS.TELEFONO - 1]).trim();
    const phoneNumberId = String(row[COLS.PHONE_NUMBER_ID - 1]).trim();
    const tipo          = String(row[COLS.TIPO - 1]).trim();
    const mensaje       = String(row[COLS.MENSAJE - 1]).trim();

    if (!telefono || !mensaje || !phoneNumberId) continue;

    try {
      const resp = UrlFetchApp.fetch(`${BOT_URL}/send-followup`, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({ telefono, phone_number_id: phoneNumberId, tipo, mensaje }),
        muteHttpExceptions: true,
      });

      const resultado = JSON.parse(resp.getContentText());
      const nuevoEstado = resultado.status === 'sent' ? 'ENVIADO'
        : resultado.status === 'skip' ? 'CANCELADO'
        : 'ERROR';

      sheet.getRange(i + 1, COLS.ESTADO).setValue(nuevoEstado);
      sheet.getRange(i + 1, COLS.ESTADO).setBackground(
        nuevoEstado === 'ENVIADO'   ? '#F0FDF4' :
        nuevoEstado === 'CANCELADO' ? '#FEF9C3' : '#FEF2F2'
      );

    } catch (e) {
      console.error(`Error enviando follow-up fila ${i + 1}:`, e.toString());
      sheet.getRange(i + 1, COLS.ESTADO).setValue('ERROR');
    }
  }
}

// ── doPost: recibe acciones de guardar/cancelar follow-ups ─────────
function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const accion = data.accion;

    if (accion === 'guardar_followup') {
      guardarFollowup(data);
    } else if (accion === 'cancelar_followups') {
      cancelarFollowups(data.telefono, data.tipo_prefijo || '');
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', error: e.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function guardarFollowup(data) {
  const ss    = SpreadsheetApp.openById(SS_ID_FU);
  const sheet = getOrCreateFollowupSheet(ss);
  sheet.appendRow([
    data.telefono,
    data.phone_number_id,
    data.tipo,
    data.mensaje,
    data.fecha_envio,
    'PENDIENTE',
    data.contexto || '{}',
  ]);
}

function cancelarFollowups(telefono, tipoPrefijo) {
  const ss    = SpreadsheetApp.openById(SS_ID_FU);
  const sheet = ss.getSheetByName(HOJA_FOLLOWUP);
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const tel    = String(data[i][COLS.TELEFONO - 1]).trim();
    const tipo   = String(data[i][COLS.TIPO - 1]).trim();
    const estado = String(data[i][COLS.ESTADO - 1]).trim();

    if (tel !== telefono) continue;
    if (estado !== 'PENDIENTE') continue;
    if (tipoPrefijo && !tipo.startsWith(tipoPrefijo)) continue;

    sheet.getRange(i + 1, COLS.ESTADO).setValue('CANCELADO');
    sheet.getRange(i + 1, COLS.ESTADO).setBackground('#FEF9C3');
  }
}

function getOrCreateFollowupSheet(ss) {
  let sheet = ss.getSheetByName(HOJA_FOLLOWUP);
  if (!sheet) {
    sheet = ss.insertSheet(HOJA_FOLLOWUP);
    const headers = ['Telefono', 'Phone Number ID', 'Tipo', 'Mensaje', 'Fecha Envio', 'Estado', 'Contexto'];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#060E1F')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
  }
  return sheet;
}

// ── Crear trigger automático (ejecutar UNA VEZ manualmente) ────────
function crearTrigger() {
  // Eliminar triggers existentes del mismo nombre
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'ejecutarFollowupsPendientes')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // Crear nuevo trigger cada 30 minutos
  ScriptApp.newTrigger('ejecutarFollowupsPendientes')
    .timeBased()
    .everyMinutes(30)
    .create();

  console.log('Trigger creado: ejecutarFollowupsPendientes cada 30 minutos');
}
