// ═══════════════════════════════════════════════════════════════════════
// GOOGLE APPS SCRIPT — Portal Unificado v2.0
// Asesoría Visa Global — Roberto Acosta
// Maneja: USA DS-160 | Schengen | Reino Unido | Caso de Rechazo
// Guardar como nuevo despliegue en Apps Script. Reemplaza appscript_ds160.js
// ═══════════════════════════════════════════════════════════════════════

const ANTHROPIC_KEY = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
const EMAIL_ROBERTO = 'nanotiendaec@gmail.com';
const SS_ID        = '19yHZ5HJH5eWyFXej8ffGBT2_sttXDZtvNaCoNEzjIOU';

// Hojas del Spreadsheet por tipo de visa
const HOJAS = {
  'USA DS-160':       'USA DS-160',
  'Schengen':         'Schengen',
  'Reino Unido':      'Reino Unido',
  'Caso de Rechazo':  'Casos Rechazo',
};

// Columnas del resumen (comunes a todos los tipos)
const COL_RESUMEN = [
  'Fecha', 'Ref ID', 'Tipo Visa', 'Nombre', 'Telefono', 'Email',
  'Estado', 'Probabilidad', 'Paquete Recomendado', 'Puntos Fuertes',
  'Puntos Debiles', 'Estrategia', 'Ver Expediente'
];

// ── doPost: recibe datos del formulario ────────────────────────────
function doPost(e) {
  try {
    const payload   = JSON.parse(e.postData.contents);
    const tipoVisa  = payload.tipo_visa || 'USA DS-160';
    const refId     = payload.ref || ('REF' + Date.now().toString().slice(-6));
    const fecha     = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');
    const ss        = SpreadsheetApp.openById(SS_ID);
    const personas  = payload.personas || [];

    if (!personas.length) return ok({ error: 'Sin datos de personas' });

    const nombre    = personas[0].nombre || '';
    const datos     = personas[0].datos  || {};
    const telefono  = datos.telefono  || datos.telefono_celular || '';
    const email     = datos.email     || '';
    const nombreHoja = HOJAS[tipoVisa] || 'Otros';

    // 1. Guardar detalle en hoja específica
    const sheetDet = getOrCreateSheet(ss, nombreHoja + ' Detalle', ['Ref ID','Fecha','Nombre','Campo','Respuesta']);
    personas.forEach(p => {
      Object.entries(p.datos || {}).forEach(([campo, valor]) => {
        if (valor && valor !== '' && valor !== 'undefined') {
          sheetDet.appendRow([refId, fecha, p.nombre, campo, String(valor)]);
        }
      });
    });
    formatearCabecera(sheetDet, 1);

    // 2. Análisis IA
    const analisis = tipoVisa === 'Caso de Rechazo'
      ? analizarRechazo(datos, nombre)
      : analizarPerfil(datos, tipoVisa, nombre, personas.length);

    // 3. Guardar resumen en hoja general
    const sheetRes = getOrCreateSheet(ss, nombreHoja, COL_RESUMEN);
    const rowNum   = sheetRes.getLastRow() + 1;
    const reporteUrl = generarUrlReporte(refId, nombreHoja);

    sheetRes.appendRow([
      fecha, refId, tipoVisa, nombre, telefono, email,
      'NUEVO', analisis.probabilidad || '—', analisis.paquete || '—',
      analisis.fuertes || '—', analisis.debiles || '—',
      analisis.estrategia || '—', reporteUrl
    ]);
    formatearCabecera(sheetRes, 1);

    // Formato visual de la fila
    sheetRes.getRange(rowNum, 1, 1, COL_RESUMEN.length).setBackground('#F0FDF4');
    sheetRes.getRange(rowNum, 7).setFontColor('#166534').setFontWeight('bold');
    sheetRes.getRange(rowNum, COL_RESUMEN.length).setFormula(
      `=HYPERLINK("${reporteUrl}","Ver expediente")`
    );
    try { sheetRes.autoResizeColumns(1, COL_RESUMEN.length); } catch(e) {}

    // 4. Notificar a Roberto por email
    notificarRoberto(refId, tipoVisa, nombre, telefono, email, reporteUrl, analisis, datos);

    return ok({ refId, status: 'ok' });

  } catch(err) {
    console.error('doPost error:', err.toString());
    return ok({ error: err.toString() });
  }
}

// ── doGet: reporte HTML del expediente ────────────────────────────
function doGet(e) {
  const refId = e.parameter.id;
  const tipo  = e.parameter.tipo || 'USA DS-160';
  if (!refId) return ContentService.createTextOutput('Falta id').setMimeType(ContentService.MimeType.TEXT);

  const ss       = SpreadsheetApp.openById(SS_ID);
  const nombreHoja = HOJAS[tipo] || tipo;
  const det      = ss.getSheetByName(nombreHoja + ' Detalle');
  const res      = ss.getSheetByName(nombreHoja);
  if (!det) return ContentService.createTextOutput('Sin datos para ' + tipo).setMimeType(ContentService.MimeType.TEXT);

  const rows    = det.getDataRange().getValues().filter(r => r[0] === refId);
  const resRows = res ? res.getDataRange().getValues().filter(r => r[1] === refId) : [];
  const analisis = resRows[0] || [];

  const personas = {};
  rows.forEach(r => {
    if (!personas[r[2]]) personas[r[2]] = [];
    personas[r[2]].push({ campo: r[3], valor: r[4] });
  });

  const html = generarHTMLReporte(refId, tipo, personas, analisis);
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Análisis IA — Perfil general ──────────────────────────────────
function analizarPerfil(datos, tipoVisa, nombre, numPersonas) {
  try {
    const perfil = construirPerfil(datos, nombre, numPersonas);
    const prompt = `Eres un experto asesor de visas con 15 años de experiencia. Analiza el perfil de este solicitante de ${tipoVisa} y proporciona:

${perfil}

Responde en JSON exacto sin markdown:
{
  "probabilidad": "porcentaje estimado de aprobacion (ej: 75%)",
  "paquete": "ESENCIAL $97 o PROFESIONAL $197 o VIP $397",
  "razon_paquete": "una linea explicando por que ese paquete",
  "fuertes": "3-5 puntos fuertes del perfil separados por punto y coma",
  "debiles": "2-4 riesgos o debilidades separados por punto y coma",
  "estrategia": "5-7 pasos numerados para maximizar probabilidad de aprobacion",
  "documentos_clave": "3-5 documentos mas importantes para este caso separados por punto y coma"
}`;

    return llamarClaudeIA(prompt);
  } catch(e) {
    return { fuertes: 'Revisar manualmente', debiles: 'Revisar manualmente', estrategia: 'Revisar manualmente', probabilidad: '—', paquete: 'PROFESIONAL $197' };
  }
}

// ── Análisis IA — Caso de Rechazo ────────────────────────────────
function analizarRechazo(datos, nombre) {
  try {
    const prompt = `Eres un experto asesor de visas especializado en casos de rechazo. Analiza este caso:

CLIENTE: ${nombre}
TIPO DE VISA SOLICITADA: ${datos.tipo_visa || '—'}
CONSULADO: ${datos.consulado || '—'}
FECHA DE RECHAZO: ${datos.fecha_rechazo || '—'}
MOTIVO INDICADO: ${datos.motivo_rechazo || '—'}
TEXTO DE LA CARTA: ${datos.texto_rechazo || '—'}
RECHAZOS PREVIOS: ${datos.rechazos_previos || '—'}
CAMBIOS DESDE EL RECHAZO: ${datos.cambios_desde_rechazo || '—'}
CUANDO NECESITA VIAJAR: ${datos.fecha_viaje_objetivo || '—'}

Responde en JSON exacto sin markdown:
{
  "probabilidad": "probabilidad de exito en segunda solicitud (ej: 80%)",
  "paquete": "siempre VIP $397 para casos de rechazo",
  "razon_paquete": "explicacion breve",
  "causa_probable": "causa real mas probable del rechazo basada en el perfil",
  "fuertes": "3-5 factores favorables para revertir el rechazo",
  "debiles": "2-4 factores que dificultan la aprobacion",
  "estrategia": "plan de accion especifico de 6-8 pasos para revertir el rechazo",
  "documentos_clave": "4-6 documentos criticos que debe preparar para la segunda solicitud"
}`;

    return llamarClaudeIA(prompt);
  } catch(e) {
    return { causa_probable: 'Requiere revision manual', fuertes: '—', debiles: '—', estrategia: '—', probabilidad: '—', paquete: 'VIP $397' };
  }
}

// ── Llamar a Claude ───────────────────────────────────────────────
function llamarClaudeIA(prompt) {
  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': ANTHROPIC_KEY,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json',
    },
    payload: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 900,
      messages: [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  const json = JSON.parse(resp.getContentText());
  const text = json.content?.[0]?.text || '{}';
  const clean = text.replace(/```json|```/g, '').trim();
  return JSON.parse(clean);
}

// ── Construir perfil para el prompt ──────────────────────────────
function construirPerfil(datos, nombre, numPersonas) {
  const campos = [
    'ocupacion', 'empleador_nombre', 'salario', 'tiempo_trabajo',
    'fondos_disponibles', 'quien_paga', 'estado_civil', 'educacion',
    'viajes_previos_usa', 'visa_usa_anterior', 'rechazo_usa',
    'rechazo_otros_paises', 'viajes_internacionales',
    'visa_schengen_anterior', 'rechazo_schengen',
    'familiares_usa', 'familiares_europa', 'familiares_uk',
    'padre_en_usa', 'madre_en_usa', 'peticion_inmigrante',
    'seg_arrestado', 'seg_deportado',
  ];
  const lineas = campos
    .filter(c => datos[c] && datos[c] !== '' && datos[c] !== 'No aplica')
    .map(c => `- ${c}: ${datos[c]}`);
  return `CLIENTE: ${nombre}\nNUMERO DE VIAJEROS: ${numPersonas}\n${lineas.join('\n')}`;
}

// ── Notificar a Roberto por email ────────────────────────────────
function notificarRoberto(refId, tipoVisa, nombre, telefono, email, url, analisis, datos) {
  try {
    const esRechazo = tipoVisa === 'Caso de Rechazo';
    const colorTema = esRechazo ? '#7C3AED' : '#060E1F';
    const etiquetaTema = esRechazo ? 'CASO DE RECHAZO' : tipoVisa.toUpperCase();

    const htmlEmail = `
<div style="font-family:Calibri,Arial,sans-serif;max-width:700px;margin:0 auto;background:#F4F6F9;padding:20px">
  <div style="background:${colorTema};color:white;padding:24px 28px;border-radius:12px 12px 0 0">
    <div style="font-size:11px;font-weight:700;letter-spacing:.1em;color:rgba(255,255,255,.6);text-transform:uppercase;margin-bottom:8px">${etiquetaTema}</div>
    <h1 style="font-size:20px;font-weight:700;margin:0">${esRechazo ? 'Caso de Rechazo recibido' : 'Nuevo expediente recibido'} — ${nombre}</h1>
  </div>

  <div style="background:white;padding:24px 28px;border:1px solid #E2E8F0;border-top:none">
    <table style="width:100%;border-collapse:collapse;margin-bottom:20px">
      <tr><td style="padding:7px 0;color:#64748B;font-size:13px;width:35%">Referencia</td><td style="font-weight:700;font-size:13px">${refId}</td></tr>
      <tr><td style="padding:7px 0;color:#64748B;font-size:13px">Tipo de visa</td><td style="font-size:13px">${tipoVisa}</td></tr>
      <tr><td style="padding:7px 0;color:#64748B;font-size:13px">Nombre</td><td style="font-weight:700;font-size:13px">${nombre}</td></tr>
      <tr><td style="padding:7px 0;color:#64748B;font-size:13px">Telefono</td><td style="font-size:13px">${telefono || '—'}</td></tr>
      <tr><td style="padding:7px 0;color:#64748B;font-size:13px">Email</td><td style="font-size:13px">${email || '—'}</td></tr>
      ${esRechazo ? `<tr><td style="padding:7px 0;color:#64748B;font-size:13px">Motivo rechazo</td><td style="font-size:13px">${datos.motivo_rechazo || '—'}</td></tr>` : ''}
      ${esRechazo ? `<tr><td style="padding:7px 0;color:#64748B;font-size:13px">Cuando viaja</td><td style="font-size:13px">${datos.fecha_viaje_objetivo || '—'}</td></tr>` : ''}
    </table>

    <div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:10px;padding:16px;margin-bottom:12px">
      <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#166534;margin-bottom:8px">Analisis IA — Probabilidad: ${analisis.probabilidad || '—'} — Paquete: ${analisis.paquete || '—'}</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
        <div><div style="font-size:10px;font-weight:700;color:#166534;margin-bottom:4px">PUNTOS FUERTES</div><div style="font-size:12px;color:#1A2940">${(analisis.fuertes||'—').replace(/;/g,'<br>')}</div></div>
        <div><div style="font-size:10px;font-weight:700;color:#991B1B;margin-bottom:4px">PUNTOS DEBILES</div><div style="font-size:12px;color:#1A2940">${(analisis.debiles||'—').replace(/;/g,'<br>')}</div></div>
      </div>
    </div>

    ${esRechazo ? `<div style="background:#FEF3C7;border:1px solid #FCD34D;border-radius:10px;padding:14px;margin-bottom:12px"><div style="font-size:10px;font-weight:700;color:#92400E;margin-bottom:6px">CAUSA PROBABLE DEL RECHAZO</div><div style="font-size:12px;color:#1A2940">${analisis.causa_probable || '—'}</div></div>` : ''}

    <div style="background:#EFF6FF;border:1px solid #93C5FD;border-radius:10px;padding:14px;margin-bottom:20px">
      <div style="font-size:10px;font-weight:700;color:#1E40AF;margin-bottom:6px">ESTRATEGIA RECOMENDADA</div>
      <div style="font-size:12px;color:#1A2940;line-height:1.7">${(analisis.estrategia||'—').replace(/\n/g,'<br>')}</div>
    </div>

    <div style="background:#F4F6F9;border:1px solid #E2E8F0;border-radius:10px;padding:14px;margin-bottom:24px">
      <div style="font-size:10px;font-weight:700;color:#64748B;margin-bottom:6px">DOCUMENTOS CLAVE A SOLICITAR</div>
      <div style="font-size:12px;color:#1A2940">${(analisis.documentos_clave||'—').replace(/;/g,'<br>')}</div>
    </div>

    <a href="${url}" style="display:inline-block;background:${colorTema};color:white;padding:13px 28px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px">
      Ver expediente completo
    </a>
  </div>

  <div style="text-align:center;padding:16px;font-size:11px;color:#94A3B8">
    Asesoria Visa Global — Sistema automatizado — ${Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm')}
  </div>
</div>`;

    const asunto = esRechazo
      ? `CASO RECHAZO — ${nombre} — ${datos.tipo_visa || 'visa'} [${refId}]`
      : `Nuevo expediente ${tipoVisa} — ${nombre} [${refId}]`;

    MailApp.sendEmail({ to: EMAIL_ROBERTO, subject: asunto, htmlBody: htmlEmail });
  } catch(e) {
    console.log('Email error:', e.toString());
  }
}

// ── Generar URL del reporte HTML ──────────────────────────────────
function generarUrlReporte(refId, tipo) {
  const scriptUrl = ScriptApp.getService().getUrl();
  return `${scriptUrl}?id=${refId}&tipo=${encodeURIComponent(tipo)}`;
}

// ── Generar HTML del reporte ──────────────────────────────────────
function generarHTMLReporte(refId, tipo, personas, analisis) {
  const esRechazo = tipo === 'Caso de Rechazo';
  const prob    = analisis[7] || '—';
  const paquete = analisis[8] || '—';
  const fuertes = analisis[9] || '—';
  const debiles = analisis[10] || '—';
  const estrategia = analisis[11] || '—';

  const personasHTML = Object.entries(personas).map(([nombre, campos]) => {
    const filas = campos.map(c =>
      `<tr><td style="color:#64748B;padding:8px 0;font-size:13px;width:42%;border-bottom:1px solid #F4F6F9;padding-right:16px">${c.campo}</td><td style="font-weight:500;padding:8px 0;font-size:13px;border-bottom:1px solid #F4F6F9">${c.valor||'—'}</td></tr>`
    ).join('');
    return `<div style="background:white;border:1px solid #E2E8F0;border-radius:12px;padding:24px;margin-bottom:12px">
      <div style="font-size:15px;font-weight:700;color:#060E1F;padding-bottom:12px;margin-bottom:16px;border-bottom:2px solid #F0B429">${nombre}</div>
      <table style="width:100%;border-collapse:collapse">${filas}</table>
    </div>`;
  }).join('');

  return `<!DOCTYPE html><html lang="es"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Expediente ${refId}</title>
<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Calibri,Arial,sans-serif;background:#F4F6F9;color:#1A2940;padding:20px}.wrap{max-width:860px;margin:0 auto}</style>
</head><body><div class="wrap">
  <div style="background:#060E1F;color:white;padding:22px 28px;border-radius:12px 12px 0 0;margin-bottom:0">
    <div style="font-size:10px;letter-spacing:.1em;color:rgba(255,255,255,.5);text-transform:uppercase;margin-bottom:6px">${tipo}</div>
    <div style="font-size:18px;font-weight:700">Expediente ${refId}</div>
  </div>
  ${analisis.length ? `<div style="background:white;border:1px solid #E2E8F0;border-top:none;padding:22px 28px;margin-bottom:4px">
    <div style="display:grid;grid-template-columns:1fr 1fr 2fr;gap:14px">
      <div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:10px;padding:14px"><div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#166534;margin-bottom:6px">Puntos fuertes</div><div style="font-size:12px;color:#1A2940">${fuertes.replace(/;/g,'<br>')}</div></div>
      <div style="background:#FEF2F2;border:1px solid #FCA5A5;border-radius:10px;padding:14px"><div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#991B1B;margin-bottom:6px">Puntos debiles</div><div style="font-size:12px;color:#1A2940">${debiles.replace(/;/g,'<br>')}</div></div>
      <div style="background:#EFF6FF;border:1px solid #93C5FD;border-radius:10px;padding:14px"><div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#1E40AF;margin-bottom:6px">Estrategia — ${prob} — ${paquete}</div><div style="font-size:12px;color:#1A2940;line-height:1.6">${estrategia.replace(/\n/g,'<br>')}</div></div>
    </div>
  </div>` : ''}
  ${personasHTML}
</div></body></html>`;
}

// ── Helpers ───────────────────────────────────────────────────────
function getOrCreateSheet(ss, name, headers) {
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    if (headers) {
      s.appendRow(headers);
      s.setFrozenRows(1);
      formatearCabecera(s, 1);
    }
  }
  return s;
}

function formatearCabecera(sheet, fila) {
  try {
    const last = sheet.getLastColumn();
    if (last < 1) return;
    sheet.getRange(fila, 1, 1, last)
      .setBackground('#060E1F')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
  } catch(e) {}
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', ...data }))
    .setMimeType(ContentService.MimeType.JSON);
}
