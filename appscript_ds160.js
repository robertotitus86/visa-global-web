// ═══════════════════════════════════════════════════════════════════
// GOOGLE APPS SCRIPT — DS-160 Collector + Análisis IA
// Asesoría Visa Global — Roberto Acosta
// ═══════════════════════════════════════════════════════════════════

// La API key se guarda en: Apps Script → Configuración del proyecto → Propiedades de script
// Clave: ANTHROPIC_KEY  Valor: sk-ant-api03-AYP5F-...
const ANTHROPIC_KEY = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
const SHEET_RESUMEN   = 'Clientes DS-160';
const SHEET_DETALLE   = 'DS-160 Detalle';
const SHEET_ID        = SpreadsheetApp.getActiveSpreadsheet().getId();

const COL_RESUMEN = [
  'Fecha','ID Grupo','Personas','Teléfono','Email','Fecha viaje',
  'Estado','Puntos fuertes','Puntos débiles','Estrategia IA','Ver DS-160 completo'
];

// ── doPost: recibe datos del formulario ───────────────────────────
function doPost(e) {
  try {
    const payload   = JSON.parse(e.postData.contents);
    const grupoId   = 'DS' + Date.now();
    const fecha     = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRes  = getOrCreateSheet(ss, SHEET_RESUMEN, COL_RESUMEN);
    const sheetDet  = getOrCreateSheet(ss, SHEET_DETALLE, ['ID Grupo','Fecha','#','Nombre','Campo DS-160','Respuesta']);

    // ── 1. Escribir detalle ─────────────────────────────────────
    payload.personas.forEach((persona, idx) => {
      Object.entries(persona.datos).forEach(([campo, valor]) => {
        sheetDet.appendRow([grupoId, fecha, idx+1, persona.nombre, campo, valor]);
      });
    });
    formatearCabecera(sheetDet, 1);

    // ── 2. Generar análisis con IA ──────────────────────────────
    const analisis  = analizarConIA(payload.personas, grupoId);

    // ── 3. Datos resumen ────────────────────────────────────────
    const p1        = payload.personas[0].datos;
    const nombres   = payload.personas.map(p => p.nombre).join(', ');
    const tel1      = p1['Teléfono celular'] || '';
    const email1    = p1['Correo electrónico'] || '';
    const fViaje    = p1['Fecha llegada USA'] || '';
    const reporteUrl = `https://script.google.com/macros/s/AKfycbw50KSbad11eId3-poH8TMNuOe-KvV8tTHPWmc8lVDHfqQezc-zZ5mfyiMkEuFv3Q87Pw/exec?id=${grupoId}`;

    // ── 4. Fila en resumen ──────────────────────────────────────
    const rowNum = sheetRes.getLastRow() + 1;
    sheetRes.appendRow([
      fecha, grupoId, nombres, tel1, email1, fViaje,
      'NUEVO',
      analisis.fuertes, analisis.debiles, analisis.estrategia,
      reporteUrl
    ]);
    formatearCabecera(sheetRes, 1);

    // Hipervínculo en última columna
    sheetRes.getRange(rowNum, COL_RESUMEN.length)
      .setFormula(`=HYPERLINK("${reporteUrl}","Ver expediente")`);

    // Color fondo verde suave + negrita en Estado
    sheetRes.getRange(rowNum, 1, 1, COL_RESUMEN.length).setBackground('#F0FDF4');
    sheetRes.getRange(rowNum, 7).setFontColor('#166534').setFontWeight('bold');

    // Wrap text en columnas largas
    sheetRes.getRange(rowNum, 8, 1, 3).setWrap(true);

    // Autofit columnas
    sheetRes.autoResizeColumns(1, COL_RESUMEN.length);

    // ── 5. Notificar a Roberto ──────────────────────────────────
    notificarRoberto(grupoId, nombres, tel1, email1, fViaje, reporteUrl, analisis);

    return ok({ grupoId });
  } catch(err) {
    console.error(err.toString());
    return ok({ error: err.toString() });
  }
}

// ── doGet: reporte HTML de un grupo ──────────────────────────────
function doGet(e) {
  const grupoId = e.parameter.id;
  if(!grupoId) return ContentService.createTextOutput('Falta el parámetro id').setMimeType(ContentService.MimeType.TEXT);

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const det   = ss.getSheetByName(SHEET_DETALLE);
  const res   = ss.getSheetByName(SHEET_RESUMEN);
  if(!det) return ContentService.createTextOutput('Sin datos').setMimeType(ContentService.MimeType.TEXT);

  const rows     = det.getDataRange().getValues().filter(r => r[0] === grupoId);
  const resRows  = res ? res.getDataRange().getValues().filter(r => r[1] === grupoId) : [];
  const analisis = resRows[0] || [];

  // Agrupar por persona
  const personas = {};
  rows.forEach(r => {
    const n = r[3];
    if(!personas[n]) personas[n] = [];
    personas[n].push({ campo: r[4], valor: r[5] });
  });

  const html = `<!DOCTYPE html><html lang="es"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>DS-160 ${grupoId}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',sans-serif;background:#F3F4F6;color:#111827;padding:24px}
  .wrap{max-width:900px;margin:0 auto}
  .header{background:#060E1F;color:white;padding:24px 28px;border-radius:16px 16px 0 0;display:flex;align-items:center;justify-content:space-between}
  .header h1{font-size:1.25rem;font-weight:700}
  .header h1 span{color:#F0B429}
  .badge{background:rgba(240,180,41,.2);color:#F0B429;border:1px solid rgba(240,180,41,.3);padding:5px 14px;border-radius:99px;font-size:.75rem;font-weight:700}
  .analisis{background:white;border:1px solid #E5E7EB;border-top:none;padding:28px;margin-bottom:4px}
  .analisis h2{font-size:1rem;font-weight:700;color:#060E1F;margin-bottom:16px;display:flex;align-items:center;gap:8px}
  .cols{display:grid;grid-template-columns:1fr 1fr 2fr;gap:16px}
  .box{padding:16px;border-radius:12px;font-size:.875rem;line-height:1.6}
  .box-green{background:#F0FDF4;border:1px solid #86EFAC;color:#166534}
  .box-red{background:#FEF2F2;border:1px solid #FCA5A5;color:#991B1B}
  .box-blue{background:#EFF6FF;border:1px solid #93C5FD;color:#1E40AF}
  .box h3{font-size:.75rem;text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px;opacity:.7}
  .persona-card{background:white;border:1px solid #E5E7EB;border-top:none;padding:24px 28px;margin-bottom:4px}
  .persona-card:last-child{border-radius:0 0 16px 16px;margin-bottom:0}
  .persona-title{font-size:1rem;font-weight:700;color:#060E1F;margin-bottom:20px;padding-bottom:12px;border-bottom:2px solid #F0B429;display:flex;align-items:center;gap:10px}
  .grupo-title{font-size:.7rem;text-transform:uppercase;letter-spacing:.08em;color:#9CA3AF;font-weight:700;margin:20px 0 10px;padding-top:12px;border-top:1px solid #F3F4F6}
  .grupo-title:first-of-type{margin-top:0;border-top:none}
  table{width:100%;border-collapse:collapse}
  tr:hover{background:#F9FAFB}
  td{padding:9px 0;border-bottom:1px solid #F3F4F6;font-size:.875rem;vertical-align:top}
  td:first-child{color:#6B7280;width:42%;padding-right:16px}
  td:last-child{font-weight:500;color:#111827}
  @media(max-width:680px){.cols{grid-template-columns:1fr}.cols .box:last-child{grid-column:1}}
  @media print{.no-print{display:none}.persona-card{page-break-inside:avoid}}
</style>
</head><body>
<div class="wrap">
  <div class="header">
    <h1>📋 DS-160 <span>${grupoId}</span></h1>
    <span class="badge">✓ Expediente completo</span>
  </div>

  ${analisis.length ? `
  <div class="analisis">
    <h2>🤖 Análisis IA — Estrategia de visado</h2>
    <div class="cols">
      <div class="box box-green"><h3>✅ Puntos fuertes</h3>${(analisis[7]||'—').replace(/\n/g,'<br>')}</div>
      <div class="box box-red"><h3>⚠️ Puntos débiles</h3>${(analisis[8]||'—').replace(/\n/g,'<br>')}</div>
      <div class="box box-blue"><h3>🎯 Estrategia recomendada</h3>${(analisis[9]||'—').replace(/\n/g,'<br>')}</div>
    </div>
  </div>` : ''}

  ${Object.entries(personas).map(([nombre, campos], i) => {
    // Agrupar campos por sección DS-160
    const grupos_ds = {
      'DATOS PERSONALES':['Apellidos','Nombres','Otros nombres','Fecha de nacimiento','Ciudad y país de nacimiento','Nacionalidad','Otra nacionalidad','Número de cédula','Sexo','Estado civil'],
      'PASAPORTE':['Número de pasaporte','Fecha de emisión','Fecha de vencimiento','Ciudad de emisión','Pasaporte perdido'],
      'EL VIAJE':['Fecha llegada USA','Días de estadía','Dirección en USA','Contacto en USA','Quién paga','Fondos disponibles'],
      'CONTACTO':['Dirección en Ecuador','Teléfono celular','Correo electrónico','Redes sociales'],
      'FAMILIA':['Nombre del padre','Nacimiento del padre','Padre en USA','Nombre de la madre','Nacimiento de la madre','Madre en USA','Otros familiares en USA','Datos de la pareja'],
      'TRABAJO / EDUCACIÓN':['Ocupación','Empresa o institución','Dirección del empleador','Teléfono del empleador','Ingreso mensual USD','Tiempo en el trabajo','Trabajo anterior','Nivel de educación'],
      'HISTORIAL DE VIAJES':['Visitas previas USA','Visa USA anterior','Rechazo visa USA','Rechazo otros países','Viajes internacionales','Petición inmigrante'],
      'SEGURIDAD':['Enfermedad contagiosa','Consumo de drogas','Antecedentes penales','Deportación previa','Actividad terrorista'],
    };

    const mapa = {};
    campos.forEach(c => mapa[c.campo] = c.valor);

    let html_persona = `<div class="persona-card"><div class="persona-title">👤 Persona ${i+1}: ${nombre}</div>`;
    Object.entries(grupos_ds).forEach(([g, fields]) => {
      const rows_g = fields.filter(f => mapa[f] !== undefined);
      if(!rows_g.length) return;
      html_persona += `<div class="grupo-title">${g}</div><table>`;
      rows_g.forEach(f => {
        html_persona += `<tr><td>${f}</td><td>${mapa[f]||'—'}</td></tr>`;
      });
      // Campos extra no clasificados
      html_persona += `</table>`;
    });
    html_persona += `</div>`;
    return html_persona;
  }).join('')}

</div>
<script>window.onload=()=>{if(window.location.search.includes('print'))window.print();}</script>
</body></html>`;

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Análisis con Claude ───────────────────────────────────────────
function analizarConIA(personas, grupoId) {
  try {
    // Construir resumen del perfil del cliente
    const p1 = personas[0].datos;
    const perfil = `
PERFIL DEL SOLICITANTE PRINCIPAL:
- Nombre: ${personas[0].nombre}
- Estado civil: ${p1['Estado civil']||'—'}
- Ocupación: ${p1['Ocupación']||'—'}
- Empresa: ${p1['Empresa o institución']||'—'}
- Ingreso mensual: ${p1['Ingreso mensual USD']||'—'}
- Tiempo en trabajo: ${p1['Tiempo en el trabajo']||'—'}
- Fondos disponibles: ${p1['Fondos disponibles']||'—'}
- Días de estadía: ${p1['Días de estadía']||'—'}
- Quién paga: ${p1['Quién paga']||'—'}
- Visitas previas USA: ${p1['Visitas previas USA']||'—'}
- Visa USA anterior: ${p1['Visa USA anterior']||'—'}
- Rechazo visa USA: ${p1['Rechazo visa USA']||'—'}
- Rechazo otros países: ${p1['Rechazo otros países']||'—'}
- Viajes internacionales: ${p1['Viajes internacionales']||'—'}
- Petición inmigrante: ${p1['Petición inmigrante']||'—'}
- Padre en USA: ${p1['Padre en USA']||'—'}
- Madre en USA: ${p1['Madre en USA']||'—'}
- Otros familiares en USA: ${p1['Otros familiares en USA']||'—'}
- Enfermedad: ${p1['Enfermedad contagiosa']||'No'}
- Drogas: ${p1['Consumo de drogas']||'No'}
- Antecedentes: ${p1['Antecedentes penales']||'No'}
- Deportación: ${p1['Deportación previa']||'No'}
- Total viajeros: ${personas.length}
${personas.length>1?'- Viajan con: '+personas.slice(1).map(p=>p.nombre).join(', '):''}
`;

    const prompt = `Eres un experto asesor de visas con 15 años de experiencia. Analiza el perfil de este solicitante de visa B1/B2 (turismo USA) y proporciona:

${perfil}

Responde en JSON exacto con esta estructura (sin markdown, solo JSON puro):
{
  "fuertes": "lista de 3-5 puntos fuertes del perfil, cada uno en una línea con • al inicio",
  "debiles": "lista de 2-4 puntos débiles o riesgos, cada uno en una línea con • al inicio",
  "estrategia": "estrategia detallada de 5-7 pasos para maximizar probabilidad de aprobación, cada paso numerado"
}

Si no hay suficiente información para algún punto, indícalo brevemente. Sé específico y práctico.`;

    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': ANTHROPIC_KEY,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 800,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });

    const json = JSON.parse(resp.getContentText());
    const text = json.content?.[0]?.text || '{}';
    const parsed = JSON.parse(text.replace(/```json|```/g,'').trim());
    return {
      fuertes:   parsed.fuertes   || 'Análisis no disponible',
      debiles:   parsed.debiles   || 'Análisis no disponible',
      estrategia:parsed.estrategia|| 'Análisis no disponible',
    };
  } catch(e) {
    console.error('Error IA:', e.toString());
    return {
      fuertes:   'Revisar manualmente',
      debiles:   'Revisar manualmente',
      estrategia:'Revisar manualmente',
    };
  }
}

// ── Email a Roberto ───────────────────────────────────────────────
function notificarRoberto(grupoId, nombres, tel, email, fViaje, url, analisis) {
  try {
    MailApp.sendEmail({
      to: 'nanotiendaec@gmail.com',
      subject: `🆕 DS-160 nuevo — ${nombres}`,
      htmlBody: `
<div style="font-family:'Segoe UI',sans-serif;max-width:680px;margin:0 auto;background:#F3F4F6;padding:24px">
  <div style="background:#060E1F;color:white;padding:24px 28px;border-radius:12px 12px 0 0">
    <h1 style="font-size:1.1rem;font-weight:700;margin:0">📋 Nuevo DS-160 recibido</h1>
  </div>
  <div style="background:white;padding:24px 28px;border:1px solid #E5E7EB;border-top:none">
    <table style="width:100%;border-collapse:collapse;margin-bottom:20px">
      <tr><td style="padding:8px 0;color:#6B7280;font-size:.875rem">Personas</td><td style="padding:8px 0;font-weight:700">${nombres}</td></tr>
      <tr><td style="padding:8px 0;color:#6B7280;font-size:.875rem">Teléfono</td><td style="padding:8px 0">${tel}</td></tr>
      <tr><td style="padding:8px 0;color:#6B7280;font-size:.875rem">Email</td><td style="padding:8px 0">${email}</td></tr>
      <tr><td style="padding:8px 0;color:#6B7280;font-size:.875rem">Fecha de viaje</td><td style="padding:8px 0">${fViaje}</td></tr>
    </table>

    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:20px">
      <div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:10px;padding:14px">
        <div style="font-size:.7rem;text-transform:uppercase;letter-spacing:.06em;color:#166534;font-weight:700;margin-bottom:8px">✅ Puntos fuertes</div>
        <div style="font-size:.85rem;color:#166534;line-height:1.6">${(analisis.fuertes||'—').replace(/\n/g,'<br>')}</div>
      </div>
      <div style="background:#FEF2F2;border:1px solid #FCA5A5;border-radius:10px;padding:14px">
        <div style="font-size:.7rem;text-transform:uppercase;letter-spacing:.06em;color:#991B1B;font-weight:700;margin-bottom:8px">⚠️ Puntos débiles</div>
        <div style="font-size:.85rem;color:#991B1B;line-height:1.6">${(analisis.debiles||'—').replace(/\n/g,'<br>')}</div>
      </div>
    </div>
    <div style="background:#EFF6FF;border:1px solid #93C5FD;border-radius:10px;padding:14px;margin-bottom:24px">
      <div style="font-size:.7rem;text-transform:uppercase;letter-spacing:.06em;color:#1E40AF;font-weight:700;margin-bottom:8px">🎯 Estrategia recomendada</div>
      <div style="font-size:.85rem;color:#1E40AF;line-height:1.7">${(analisis.estrategia||'—').replace(/\n/g,'<br>')}</div>
    </div>

    <a href="${url}" style="display:inline-block;background:#060E1F;color:white;padding:14px 28px;border-radius:10px;text-decoration:none;font-weight:700;font-size:.9rem">
      Ver expediente DS-160 completo →
    </a>
  </div>
  <div style="text-align:center;padding:16px;font-size:.75rem;color:#9CA3AF">
    Asesoría Visa Global · Sistema automatizado
  </div>
</div>`
    });
  } catch(e) { console.log('Email error:', e); }
}

// ── Helpers ───────────────────────────────────────────────────────
function getOrCreateSheet(ss, name, headers) {
  let s = ss.getSheetByName(name);
  if(!s) {
    s = ss.insertSheet(name);
    if(headers) {
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
    if(last < 1) return;
    const r = sheet.getRange(fila, 1, 1, last);
    r.setBackground('#060E1F').setFontColor('#FFFFFF').setFontWeight('bold');
  } catch(e) {}
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ status:'ok', ...data }))
    .setMimeType(ContentService.MimeType.JSON);
}
