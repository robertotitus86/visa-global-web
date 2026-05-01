// ═══════════════════════════════════════════════════════════════════
// GOOGLE APPS SCRIPT — DS-160 Collector
// Asesoría Visa Global — Roberto Acosta
//
// INSTRUCCIONES DE INSTALACIÓN:
// 1. Abre Google Sheets → crea un archivo nuevo → nómbralo "DS-160 Clientes"
// 2. En el menú: Extensiones → Apps Script
// 3. Borra todo el código existente y pega este archivo completo
// 4. Guarda (Ctrl+S)
// 5. Clic en "Implementar" → "Nueva implementación"
//    → Tipo: Aplicación web
//    → Ejecutar como: Yo (tu cuenta)
//    → Quién puede acceder: Cualquier persona
// 6. Copia la URL generada
// 7. Pégala en ds160.html donde dice REEMPLAZAR_CON_TU_URL
// 8. Sube ds160.html a GitHub (git push)
// ═══════════════════════════════════════════════════════════════════

const SHEET_RESUMEN   = 'Clientes DS-160';   // pestaña resumen
const SHEET_DETALLE   = 'DS-160 Detalle';    // pestaña con todos los campos
const SHEET_ID        = SpreadsheetApp.getActiveSpreadsheet().getId();

// Columnas del resumen
const COL_RESUMEN = [
  'Fecha',
  'ID Grupo',
  'Personas',
  'Teléfono (persona 1)',
  'Email (persona 1)',
  'Fecha viaje',
  'Estado',
  'Reporte completo',
];

// ─── PUNTO DE ENTRADA POST ────────────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const grupoId = 'DS' + Date.now();
    const fecha   = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRes = getOrCreateSheet(ss, SHEET_RESUMEN, COL_RESUMEN);
    const sheetDet = getOrCreateSheet(ss, SHEET_DETALLE, null);

    // Escribir detalle (una fila por campo por persona)
    const detalleStartRow = sheetDet.getLastRow() + 1;
    payload.personas.forEach((persona, idx) => {
      Object.entries(persona.datos).forEach(([campo, valor]) => {
        sheetDet.appendRow([
          grupoId,
          fecha,
          idx + 1,
          persona.nombre,
          campo,
          valor,
        ]);
      });
    });

    // URL del reporte filtrado (rango en detalle)
    const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit#gid=${sheetDet.getSheetId()}&range=A1`;
    const reporteUrl     = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit#gid=${sheetDet.getSheetId()}`;

    // Nombres de personas
    const nombres = payload.personas.map(p => p.nombre).join(', ');

    // Datos persona 1 para el resumen
    const p1     = payload.personas[0].datos;
    const tel1   = p1['Teléfono celular'] || '';
    const email1 = p1['Correo electrónico'] || '';
    const fecha_viaje = p1['Fecha de llegada a USA'] || '';

    // Fila del resumen con hipervínculo al reporte
    const row = [
      fecha,
      grupoId,
      nombres,
      tel1,
      email1,
      fecha_viaje,
      'NUEVO',
      reporteUrl,
    ];
    const rowNum = sheetRes.getLastRow() + 1;
    sheetRes.appendRow(row);

    // Convertir la celda del reporte en hipervínculo clicable
    const reporteCell = sheetRes.getRange(rowNum, COL_RESUMEN.length);
    reporteCell.setFormula(`=HYPERLINK("${reporteUrl}","Ver DS-160 de ${payload.personas[0].nombre.split(' ')[0]}")`);

    // Dar formato a la fila
    const nuevaFila = sheetRes.getRange(rowNum, 1, 1, COL_RESUMEN.length);
    nuevaFila.setBackground('#F0FDF4');
    sheetRes.getRange(rowNum, 7).setFontColor('#166534').setFontWeight('bold'); // Estado

    // Cabeceras del detalle si es la primera vez
    if (sheetDet.getLastRow() === Object.keys(payload.personas[0].datos).length) {
      sheetDet.getRange(1,1,1,6).setValues([['ID Grupo','Fecha','#Persona','Nombre','Campo DS-160','Respuesta']]);
      sheetDet.getRange(1,1,1,6).setBackground('#060E1F').setFontColor('#FFFFFF').setFontWeight('bold');
    }

    // Notificar a Roberto por email
    notificarRoberto(grupoId, nombres, tel1, email1, fecha_viaje, reporteUrl);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', grupoId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error(err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', msg: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── GET: ver datos de un grupo ───────────────────────────────────────────────
function doGet(e) {
  const grupoId = e.parameter.id;
  if (!grupoId) {
    return ContentService.createTextOutput('Falta el parámetro id').setMimeType(ContentService.MimeType.TEXT);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DETALLE);
  if (!sheet) return ContentService.createTextOutput('Sin datos').setMimeType(ContentService.MimeType.TEXT);

  const data = sheet.getDataRange().getValues();
  const rows = data.filter(r => r[0] === grupoId);

  let html = `<html><head>
    <meta charset="UTF-8">
    <title>DS-160 ${grupoId}</title>
    <style>
      body{font-family:sans-serif;max-width:800px;margin:40px auto;padding:0 16px;color:#060E1F}
      h1{color:#060E1F;border-bottom:3px solid #F0B429;padding-bottom:10px}
      h2{margin-top:28px;color:#374151;font-size:1rem;text-transform:uppercase;letter-spacing:.05em}
      table{width:100%;border-collapse:collapse;margin-top:8px;font-size:.9rem}
      th{background:#060E1F;color:white;padding:8px 12px;text-align:left}
      td{padding:8px 12px;border-bottom:1px solid #E5E7EB}
      tr:hover{background:#FFF3CD}
      .badge{display:inline-block;background:#F0FDF4;color:#166534;padding:3px 10px;border-radius:99px;font-size:.75rem;font-weight:700}
    </style></head><body>
    <h1>📋 DS-160 — ${grupoId}</h1>
    <p><span class="badge">✓ Formulario completo</span></p>`;

  // Agrupar por persona
  const personas = {};
  rows.forEach(r => {
    const nombre = r[3];
    if (!personas[nombre]) personas[nombre] = [];
    personas[nombre].push({ campo: r[4], valor: r[5] });
  });

  Object.entries(personas).forEach(([nombre, campos]) => {
    html += `<h2>👤 ${nombre}</h2><table><tr><th>Campo</th><th>Respuesta</th></tr>`;
    campos.forEach(c => {
      html += `<tr><td>${c.campo}</td><td>${c.valor || '—'}</td></tr>`;
    });
    html += `</table>`;
  });

  html += `</body></html>`;
  return HtmlService.createHtmlOutput(html);
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#060E1F')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function notificarRoberto(grupoId, nombres, telefono, email, fechaViaje, reporteUrl) {
  try {
    MailApp.sendEmail({
      to: 'nanotiendaec@gmail.com',
      subject: `🆕 DS-160 nuevo — ${nombres}`,
      htmlBody: `
        <div style="font-family:sans-serif;max-width:600px;margin:0 auto">
          <h2 style="color:#060E1F;border-bottom:3px solid #F0B429;padding-bottom:10px">
            Nuevo formulario DS-160 recibido
          </h2>
          <table style="width:100%;border-collapse:collapse">
            <tr><td style="padding:8px;color:#6B7280">Personas:</td><td style="padding:8px;font-weight:bold">${nombres}</td></tr>
            <tr><td style="padding:8px;color:#6B7280">Teléfono:</td><td style="padding:8px">${telefono}</td></tr>
            <tr><td style="padding:8px;color:#6B7280">Email:</td><td style="padding:8px">${email}</td></tr>
            <tr><td style="padding:8px;color:#6B7280">Fecha viaje:</td><td style="padding:8px">${fechaViaje}</td></tr>
            <tr><td style="padding:8px;color:#6B7280">ID:</td><td style="padding:8px">${grupoId}</td></tr>
          </table>
          <br>
          <a href="${reporteUrl}" style="background:#060E1F;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:bold;display:inline-block">
            Ver DS-160 completo →
          </a>
          <p style="color:#9CA3AF;font-size:.85rem;margin-top:24px">Asesoría Visa Global — Sistema automatizado</p>
        </div>
      `
    });
  } catch(e) {
    console.log('Email no enviado:', e);
  }
}
