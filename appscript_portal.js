// ═══════════════════════════════════════════════════════════════════════
// GOOGLE APPS SCRIPT — Portal Unificado v3.0
// Asesoría Visa Global — Roberto Acosta
// Maneja: USA DS-160 | Schengen | Reino Unido | Caso de Rechazo
//         + intake_familiar (nuevo intake.html multi-viajero DS-160)
// ═══════════════════════════════════════════════════════════════════════

const ANTHROPIC_KEY = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
const ADMIN_PIN    = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || 'visa2026';
const EMAIL_ROBERTO = 'nanotiendaec@gmail.com';
const SS_ID        = '19yHZ5HJH5eWyFXej8ffGBT2_sttXDZtvNaCoNEzjIOU';

// Columnas del CRM de casos
const COL_CRM = [
  'Ref ID','Fecha','Estado','Tipo Visa','Nombre Principal','Num Viajeros',
  'Telefono','Email','Probabilidad','Paquete','Llegada USA','Dias Estancia',
  'Pago','Cita','Notas','Ver Expediente'
];

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

// ── doPost: enruta según subtipo ──────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.subtipo === 'save_draft')      return saveDraftAction(payload);
    if (payload.subtipo === 'intake_familiar') return manejarIntakeFamiliar(payload);
    if (payload.subtipo === 'extractDS160')    return extractarDS160(payload);
    return manejarPortalLegacy(payload);
  } catch(err) {
    console.error('doPost error:', err.toString());
    return ok({ error: err.toString() });
  }
}

// ── BORRADORES EN LA NUBE (para intake.html cross-device) ─────────
function getDraftAction(e) {
  const ref = (e.parameter.ref || '').trim();
  if (!ref) return ok({ ok: false, msg: 'no_ref' });
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName('Borradores');
    if (!sh) return ok({ ok: false, msg: 'no_sheet' });
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === ref) {
        return ContentService
          .createTextOutput(rows[i][1])
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ok({ ok: false, msg: 'not_found' });
  } catch(err) {
    return ok({ ok: false, msg: err.toString() });
  }
}

function saveDraftAction(payload) {
  const ref  = (payload.ref  || '').trim();
  const data = payload.data  || {};
  if (!ref) return ok({ ok: false, msg: 'no_ref' });
  try {
    const ss  = SpreadsheetApp.openById(SS_ID);
    let sh    = ss.getSheetByName('Borradores');
    if (!sh) {
      sh = ss.insertSheet('Borradores');
      sh.appendRow(['Ref', 'Datos JSON', 'Actualizado', 'Paso']);
      formatearCabecera(sh, 1);
    }
    const now  = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');
    const json = JSON.stringify(data);
    const paso = data.currentStep || 1;
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === ref) {
        sh.getRange(i + 1, 2, 1, 3).setValues([[json, now, paso]]);
        return ok({ ok: true });
      }
    }
    sh.appendRow([ref, json, now, paso]);
    return ok({ ok: true });
  } catch(err) {
    return ok({ ok: false, msg: err.toString() });
  }
}

// ── NUEVO: Handler intake_familiar (intake.html multi-viajero) ────
function manejarIntakeFamiliar(payload) {
  const ss       = SpreadsheetApp.openById(SS_ID);
  const personas = payload.personas || [];
  const shared   = payload.shared   || {};
  const refId    = payload.ref      || ('INK' + Date.now().toString().slice(-6));
  const fecha    = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');
  const numViaj  = personas.length;

  if (!numViaj) return ok({ error: 'Sin viajeros' });

  const primerViajero = personas[0];
  const nombrePrincipal = primerViajero.nombre || '';
  const datosP0 = primerViajero.datos || {};
  const emailCliente = datosP0.email || '';
  const telCliente   = datosP0.primaryPhone || '';

  // 1. Guardar detalle completo en hoja "Intake DS-160 Detalle"
  const sheetDet = getOrCreateSheet(ss, 'Intake DS-160 Detalle',
    ['Ref ID','Fecha','Viajero','Rol','Campo','Respuesta']);
  // Shared info
  Object.entries(shared).forEach(([campo, valor]) => {
    if (valor && String(valor).trim()) {
      sheetDet.appendRow([refId, fecha, 'COMPARTIDO', '—', campo, String(valor)]);
    }
  });
  // Per-traveler data
  personas.forEach(p => {
    const datos = p.datos || {};
    Object.entries(datos).forEach(([campo, valor]) => {
      if (valor && String(valor).trim() && String(valor) !== 'N/A') {
        sheetDet.appendRow([refId, fecha, p.nombre, p.rol, campo, String(valor)]);
      }
    });
    // Security flags
    if (p.banderas_seguridad && p.banderas_seguridad.length) {
      sheetDet.appendRow([refId, fecha, p.nombre, p.rol, 'BANDERAS_SEGURIDAD',
        p.banderas_seguridad.join(', ')]);
    }
  });
  formatearCabecera(sheetDet, 1);

  // 2. Análisis IA familiar
  const analisis = analizarPerfilFamiliar(shared, personas, refId);

  // 3. Guardar/actualizar en CRM de casos
  const sheetCRM   = getOrCreateSheet(ss, 'CASOS CRM', COL_CRM);
  const reporteUrl  = generarUrlReporte(refId, 'USA DS-160');
  const crmData     = sheetCRM.getDataRange().getValues();
  const crmHeaders  = crmData[0].map(h => String(h).trim());
  const refColIdx   = crmHeaders.indexOf('Ref ID');
  const existingRow = crmData.findIndex((row, i) => i > 0 && String(row[refColIdx]) === refId);

  if (existingRow > 0) {
    // Actualizar fila existente (caso creado desde el CRM)
    const rowNum = existingRow + 1;
    const fieldsToUpdate = {
      'Estado':        'Formulario Recibido',
      'Num Viajeros':  numViaj,
      'Telefono':      telCliente || crmData[existingRow][crmHeaders.indexOf('Telefono')],
      'Email':         emailCliente || crmData[existingRow][crmHeaders.indexOf('Email')],
      'Probabilidad':  analisis.probabilidad || '—',
      'Paquete':       analisis.paquete || '—',
      'Llegada USA':   shared.intendedArrival || '—',
      'Dias Estancia': shared.lengthOfStayDays || '—',
      'Ver Expediente': reporteUrl,
    };
    Object.entries(fieldsToUpdate).forEach(([campo, valor]) => {
      const colIdx = crmHeaders.indexOf(campo);
      if (colIdx >= 0) sheetCRM.getRange(rowNum, colIdx + 1).setValue(valor);
    });
    sheetCRM.getRange(rowNum, 1, 1, COL_CRM.length).setBackground('#EFF6FF');
    sheetCRM.getRange(rowNum, crmHeaders.indexOf('Estado') + 1)
      .setFontColor('#1E40AF').setFontWeight('bold');
  } else {
    // Crear fila nueva (cliente llegó directo sin pasar por CRM)
    const crmRow = sheetCRM.getLastRow() + 1;
    sheetCRM.appendRow([
      refId, fecha, 'Formulario Recibido', 'USA DS-160',
      nombrePrincipal, numViaj, telCliente, emailCliente,
      analisis.probabilidad || '—', analisis.paquete || '—',
      shared.intendedArrival || '—', shared.lengthOfStayDays || '—',
      'Pendiente', 'Por agendar', '', reporteUrl
    ]);
    sheetCRM.getRange(crmRow, 1, 1, COL_CRM.length).setBackground('#F0FDF4');
    sheetCRM.getRange(crmRow, 3).setFontColor('#166534').setFontWeight('bold');
  }
  formatearCabecera(sheetCRM, 1);
  try { sheetCRM.autoResizeColumns(1, COL_CRM.length); } catch(e) {}

  // 4. Email a Roberto con análisis completo
  notificarRobertoIntakeFamiliar(refId, fecha, nombrePrincipal, numViaj,
    telCliente, emailCliente, shared, personas, analisis, reporteUrl);

  // 5. Email al cliente con próximos pasos
  if (emailCliente) {
    notificarClienteIntake(emailCliente, nombrePrincipal, refId, numViaj, shared);
  }

  return ok({ refId, status: 'ok', analisis: analisis.probabilidad });
}

// ── Análisis IA — Familia completa ───────────────────────────────
function analizarPerfilFamiliar(shared, personas, refId) {
  try {
    const prompt = construirPromptFamiliar(shared, personas);
    return llamarClaudeIA(prompt);
  } catch(e) {
    console.error('analizarPerfilFamiliar error:', e.toString());
    return {
      probabilidad: '—', paquete: 'PROFESIONAL $197',
      fuertes: 'Revisar manualmente', debiles: 'Revisar manualmente',
      estrategia: 'Revisar caso manualmente', documentos_clave: '—',
      proximos_pasos: '1. Contactar al cliente por WhatsApp\n2. Revisar pasaportes\n3. Solicitar documentos laborales'
    };
  }
}

function construirPromptFamiliar(shared, personas) {
  const numViaj = personas.length;
  const perfilesDetalle = personas.map((p, i) => {
    const d = p.datos || {};
    const rol    = p.rol || 'Adulto';
    const edadAprox = p.fecha_nacimiento
      ? Math.floor((Date.now() - new Date(p.fecha_nacimiento)) / (365.25*24*3600*1000))
      : '?';
    const esAdulto = !rol.includes('Menor') && edadAprox >= 14;
    const flags    = p.banderas_seguridad || [];
    let lineas = [
      `VIAJERO ${i+1}: ${p.nombre} | Rol: ${rol} | Edad aprox: ${edadAprox}`,
      `  Estado civil: ${d.maritalStatus || '—'}`,
      `  Nacionalidad: ${d.nationality || 'Ecuador'}`,
      `  Ciudad residencia: ${d.homeCity || '—'}`,
    ];
    if (esAdulto) {
      lineas = lineas.concat([
        `  Situacion laboral: ${d.employmentStatus || '—'}`,
        `  Cargo: ${d.currentOccupation || '—'}`,
        `  Empleador: ${d.currentEmployerName || '—'}`,
        `  Salario mensual (USD): ${d.monthlySalary || '—'}`,
        `  Tiempo en empleo actual: desde ${d.currentStartDate || '—'}`,
      ]);
    }
    lineas = lineas.concat([
      `  Ha estado en USA antes: ${d.hasBeenInUS || 'no'}`,
      `  Ha tenido visa USA: ${d.hasHadUSVisa || 'no'}`,
      `  Le han negado visa USA: ${d.hasBeenRefused || 'no'}`,
      `  ${d.hasBeenRefused==='si' ? 'DETALLE RECHAZO: '+d.refusalDetails : ''}`,
      `  Paises visitados 5 anos: ${d.countriesVisited5Years || 'Ninguno'}`,
      `  Familiares en USA: ${d.hasRelativesInUS==='si' ? d.relativesInUSDetails : 'No'}`,
      `  Padre en USA: ${d.fatherInUS || 'No'}`,
      `  Madre en USA: ${d.motherInUS || 'No'}`,
      `  Banderas de seguridad marcadas: ${flags.length ? flags.join(', ') : 'Ninguna'}`,
    ]);
    return lineas.filter(l => l.trim() && !l.endsWith(': —') && !l.endsWith(': ')).join('\n');
  }).join('\n\n');

  return `Eres el mejor asesor de visas B1/B2 USA del mundo, con 20 anos de experiencia logrando aprobaciones para ciudadanos ecuatorianos. Tu objetivo es siempre llevar la probabilidad de aprobacion al rango 80-100%. Conoces exactamente como piensan los consules americanos en Guayaquil y Quito.

Analiza este caso y dame una estrategia CONCRETA Y ESPECIFICA para maximizar la aprobacion:

VIAJE:
- Viajeros: ${numViaj} personas
- Tipo de grupo: ${shared.tipoGrupo || '—'}
- Llegada USA: ${shared.intendedArrival || '—'} por ${shared.lengthOfStayDays || '—'} dias
- Alojamiento: ${shared.usStayType || '—'} — ${shared.usStayName || '—'}, ${shared.usStayCity || '—'}, ${shared.usStayState || '—'}
- Paga: ${shared.whoIsPaying || '—'}
- Contacto USA: ${shared.usContactName || '—'} (${shared.usContactRelationship || '—'})
- Proposito especifico: ${shared.purposeNote || 'Turismo vacacional'}
- Motivacion real: ${shared.motivacionReal || '—'}
- HAY MENORES: ${shared.hayMenores || 'no'} | Detalles: ${shared.menoresDetalles || '—'}
- Autorizacion otro progenitor lista: ${shared.autorizacionLista || '—'}
- RECHAZO PREVIO EN GRUPO: ${shared.hayRechazo || 'no'} | Detalles y que cambio: ${shared.rechazoDetalles || '—'}
- PAREJA SIN MATRIMONIO EN GRUPO: ${shared.hayPareja || 'no'} | Detalles: ${shared.parejaDetalles || '—'}

PERFILES:
${perfilesDetalle}

CONTEXTO DEL CONSULADO USA EN ECUADOR:
- Los consules rechazan principalmente por: falta de lazos con Ecuador, fondos insuficientes, historial de overstay, parientes en USA sin explicacion clara, proposito vago del viaje.
- Aprueban cuando ven: empleo estable con carta del empleador, propiedad en Ecuador, hijos en escuela, fondos bancarios de al menos 3 meses de salario, itinerario especifico, boleto de regreso confirmado.
- Para familias: el consul evalua al titular principal. Si el titular es solido, la familia se aprueba con el.
- Duraciones de 10-21 dias son las mas seguras para primera visa.

Responde en JSON exacto sin markdown ni bloques de codigo:
{
  "probabilidad": "porcentaje estimado de aprobacion CON la estrategia aplicada (objetivo: 80-100%)",
  "probabilidad_sin_estrategia": "probabilidad actual sin preparacion adicional",
  "paquete": "ESENCIAL $97 o PROFESIONAL $197 o VIP $397",
  "razon_paquete": "una linea explicando por que ese paquete segun la complejidad del caso",
  "fuertes": "3-5 puntos fuertes del perfil que el consul vera positivamente, separados por punto y coma",
  "debiles": "2-4 debilidades reales del perfil que pueden causar rechazo, separados por punto y coma",
  "estrategia": "plan CONCRETO de 7-10 pasos numerados para llevar este caso al 80-100%. Cada paso debe ser especifico para ESTE caso, no generico. Incluir: que documentos preparar exactamente, como presentarlos, que decir en la entrevista sobre cada punto debil, que enfatizar.",
  "documentos_exactos": "lista de 5-8 documentos especificos con descripcion de como deben estar (ej: 'Carta laboral en papel membretado con cargo, salario y fecha de inicio, firmada por RRHH o gerente general' — NO solo 'carta laboral'), separados por punto y coma",
  "guion_entrevista": "3-5 preguntas que el consul hara SEGURO a este perfil y la respuesta ideal para cada una. Formato: PREGUNTA: [pregunta] | RESPUESTA IDEAL: [respuesta]",
  "proximos_pasos": "5 acciones especificas que el asesor debe hacer esta semana con este cliente, en orden de urgencia, numeradas",
  "alerta_principal": "el riesgo mas critico de este caso especifico y como neutralizarlo. Si no hay riesgo critico, escribir PERFIL LIMPIO",
  "campos_faltantes": "lista de preguntas ESPECIFICAS que el asesor debe hacerle al cliente porque no estan en el formulario y son necesarias para completar el perfil. Formato: '1. [PERSONA] — [PREGUNTA EXACTA A HACERLE]'. Si el perfil esta completo, escribir PERFIL COMPLETO. Ejemplos: '1. Carlos — Cuanto lleva en su empleo actual y cuanto gana mensualmente. 2. Sofia — Tiene la autorizacion notariada del padre para viajar.'"
}`;
}

// ── Email a Roberto — Intake Familiar ────────────────────────────
function notificarRobertoIntakeFamiliar(refId, fecha, nombre, numViaj, tel, email, shared, personas, analisis, reporteUrl) {
  try {
    const esAlerta = analisis.alerta_principal && !analisis.alerta_principal.includes('LIMPIO');
    const colorAlerta = esAlerta ? '#7C3AED' : '#060E1F';

    const viajerosList = personas.map((p, i) =>
      `<tr><td style="padding:5px 8px;font-size:12px;border-bottom:1px solid #F4F6F9">${i+1}. ${p.nombre}</td>
       <td style="padding:5px 8px;font-size:12px;border-bottom:1px solid #F4F6F9;color:#64748B">${p.rol}</td>
       <td style="padding:5px 8px;font-size:12px;border-bottom:1px solid #F4F6F9;color:#64748B">${(p.datos||{}).employmentStatus||'—'}</td>
       <td style="padding:5px 8px;font-size:12px;border-bottom:1px solid #F4F6F9">${(p.banderas_seguridad||[]).length?'<span style="color:#ef4444">'+p.banderas_seguridad.length+' flag(s)</span>':'<span style="color:#22c55e">Limpio</span>'}</td></tr>`
    ).join('');

    const html = `
<div style="font-family:Calibri,Arial,sans-serif;max-width:720px;margin:0 auto;background:#F4F6F9;padding:20px">

  <div style="background:${colorAlerta};color:white;padding:22px 28px;border-radius:12px 12px 0 0">
    <div style="font-size:10px;font-weight:700;letter-spacing:.1em;color:rgba(255,255,255,.6);text-transform:uppercase;margin-bottom:6px">
      INTAKE FAMILIAR — USA TURISMO B1/B2
    </div>
    <h1 style="font-size:20px;font-weight:700;margin:0 0 4px">Nuevo caso — ${nombre}</h1>
    <div style="font-size:13px;opacity:.8">${numViaj} viajero${numViaj>1?'s':''} · Llegada aprox: ${shared.intendedArrival||'—'} · ${shared.lengthOfStayDays||'—'} dias · ${shared.usStayCity||'?'}</div>
  </div>

  <div style="background:white;border:1px solid #E2E8F0;border-top:none;padding:22px 28px">

    <!-- ALERTA si hay bandera roja -->
    ${esAlerta ? `<div style="background:#FEF2F2;border:1px solid #FCA5A5;border-radius:10px;padding:14px;margin-bottom:18px">
      <div style="font-size:10px;font-weight:700;color:#991B1B;text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px">ALERTA — REVISAR</div>
      <div style="font-size:13px;color:#7F1D1D">${analisis.alerta_principal||''}</div>
    </div>` : `<div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:10px;padding:10px 14px;margin-bottom:18px;font-size:12px;color:#166534;font-weight:600">
      Perfil limpio — sin banderas criticas detectadas
    </div>`}

    <!-- DATOS RÁPIDOS -->
    <table style="width:100%;border-collapse:collapse;margin-bottom:20px;font-size:12px">
      <tr><td style="padding:6px 0;color:#64748B;width:38%">Referencia</td><td style="font-weight:700">${refId}</td></tr>
      <tr><td style="padding:6px 0;color:#64748B">Nombre principal</td><td style="font-weight:700">${nombre}</td></tr>
      <tr><td style="padding:6px 0;color:#64748B">Telefono</td><td>${tel||'—'}</td></tr>
      <tr><td style="padding:6px 0;color:#64748B">Email</td><td>${email||'—'}</td></tr>
      <tr><td style="padding:6px 0;color:#64748B">Alojamiento USA</td><td>${shared.usStayName||'—'}, ${shared.usStayCity||'—'}, ${shared.usStayState||'—'}</td></tr>
      <tr><td style="padding:6px 0;color:#64748B">Quien paga</td><td>${shared.whoIsPaying||'—'}</td></tr>
    </table>

    <!-- VIAJEROS -->
    <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#64748B;margin-bottom:8px">Viajeros</div>
    <table style="width:100%;border-collapse:collapse;margin-bottom:20px;background:#F8FAFC;border-radius:8px;overflow:hidden">
      <tr style="background:#E2E8F0">
        <td style="padding:6px 8px;font-size:11px;font-weight:700;color:#1A2940">#</td>
        <td style="padding:6px 8px;font-size:11px;font-weight:700;color:#1A2940">Rol</td>
        <td style="padding:6px 8px;font-size:11px;font-weight:700;color:#1A2940">Empleo</td>
        <td style="padding:6px 8px;font-size:11px;font-weight:700;color:#1A2940">Seguridad</td>
      </tr>
      ${viajerosList}
    </table>

    <!-- ANÁLISIS IA -->
    <div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:10px;padding:16px;margin-bottom:12px">
      <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#166534;margin-bottom:10px">
        ANALISIS IA — ${analisis.probabilidad||'—'} probabilidad — ${analisis.paquete||'—'}
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
        <div>
          <div style="font-size:10px;font-weight:700;color:#166534;margin-bottom:5px">PUNTOS FUERTES</div>
          <div style="font-size:12px;color:#1A2940;line-height:1.6">${(analisis.fuertes||'—').split(';').map(s=>'• '+s.trim()).join('<br>')}</div>
        </div>
        <div>
          <div style="font-size:10px;font-weight:700;color:#991B1B;margin-bottom:5px">PUNTOS DEBILES</div>
          <div style="font-size:12px;color:#1A2940;line-height:1.6">${(analisis.debiles||'—').split(';').map(s=>'• '+s.trim()).join('<br>')}</div>
        </div>
      </div>
    </div>

    <div style="background:#EFF6FF;border:1px solid #93C5FD;border-radius:10px;padding:14px;margin-bottom:12px">
      <div style="font-size:10px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em;margin-bottom:7px">ESTRATEGIA RECOMENDADA</div>
      <div style="font-size:12px;color:#1A2940;line-height:1.8">${(analisis.estrategia||'—').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong>$1</strong>')}</div>
    </div>

    <div style="background:#FEF9EE;border:1px solid #FCD34D;border-radius:10px;padding:14px;margin-bottom:12px">
      <div style="font-size:10px;font-weight:700;color:#92400E;text-transform:uppercase;letter-spacing:.07em;margin-bottom:7px">DOCUMENTOS EXACTOS A PREPARAR</div>
      <div style="font-size:12px;color:#1A2940;line-height:1.8">${(analisis.documentos_exactos||analisis.documentos_clave||'—').split(';').map((s,i)=>`<div style="padding:5px 0;border-bottom:1px solid #FEF3C7"><strong style="color:#92400E">${i+1}.</strong> ${s.trim()}</div>`).join('')}</div>
    </div>

    <div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:10px;padding:14px;margin-bottom:12px">
      <div style="font-size:10px;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:.07em;margin-bottom:7px">GUION DE ENTREVISTA — PREGUNTAS Y RESPUESTAS IDEALES</div>
      <div style="font-size:12px;color:#1A2940;line-height:1.8">${(analisis.guion_entrevista||'—').replace(/PREGUNTA:/g,'<br><strong style="color:#166534">PREGUNTA:</strong>').replace(/RESPUESTA IDEAL:/g,'<strong style="color:#1E40AF">RESPUESTA IDEAL:</strong>')}</div>
    </div>

    <div style="background:#F4F6F9;border:1px solid #E2E8F0;border-radius:10px;padding:14px;margin-bottom:12px">
      <div style="font-size:10px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:.07em;margin-bottom:7px">PROBABILIDAD — CON ESTRATEGIA VS SIN PREPARAR</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">
        <div style="text-align:center;background:#F0FDF4;border:1px solid #86EFAC;border-radius:8px;padding:10px">
          <div style="font-size:22px;font-weight:700;color:#166534">${analisis.probabilidad||'—'}</div>
          <div style="font-size:10px;color:#166534;font-weight:600">CON ESTRATEGIA APLICADA</div>
        </div>
        <div style="text-align:center;background:#FEF2F2;border:1px solid #FCA5A5;border-radius:8px;padding:10px">
          <div style="font-size:22px;font-weight:700;color:#991B1B">${analisis.probabilidad_sin_estrategia||'—'}</div>
          <div style="font-size:10px;color:#991B1B;font-weight:600">SIN PREPARACION</div>
        </div>
      </div>
    </div>

    <div style="background:#EFF6FF;border:1px solid #93C5FD;border-radius:10px;padding:14px;margin-bottom:24px">
      <div style="font-size:10px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em;margin-bottom:7px">PROXIMOS PASOS — ESTA SEMANA (EN ORDEN DE URGENCIA)</div>
      <div style="font-size:12px;color:#1A2940;line-height:1.8">${(analisis.proximos_pasos||'—').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong style="color:#1E40AF">$1</strong>')}</div>
    </div>

    ${analisis.campos_faltantes && !analisis.campos_faltantes.includes('COMPLETO') ? `
    <div style="background:#FEF2F2;border:2px solid #FCA5A5;border-radius:10px;padding:16px;margin-bottom:20px">
      <div style="font-size:11px;font-weight:700;color:#991B1B;text-transform:uppercase;letter-spacing:.07em;margin-bottom:8px">INFORMACION FALTANTE — PREGUNTAR AL CLIENTE</div>
      <div style="font-size:13px;color:#7F1D1D;line-height:1.9">${(analisis.campos_faltantes||'').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong>$1</strong>')}</div>
      <div style="margin-top:12px">
        <a href="${reporteUrl.replace('/exec?','/exec?action=verFormAdmin&')}&ref=${refId}&pin=admin" style="display:inline-block;background:#991B1B;color:white;padding:9px 18px;border-radius:6px;text-decoration:none;font-weight:700;font-size:12px">Completar datos faltantes en el admin</a>
      </div>
    </div>` : `<div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:8px;padding:10px 14px;margin-bottom:20px;font-size:12px;color:#166534;font-weight:600">Perfil completo — no faltan datos criticos</div>`}

    <a href="${reporteUrl}" style="display:inline-block;background:#060E1F;color:white;padding:13px 28px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px;margin-right:10px">
      Ver expediente completo
    </a>
    <a href="https://wa.me/${(tel||'').replace(/\D/g,'')}" style="display:inline-block;background:#22c55e;color:white;padding:13px 28px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px">
      WhatsApp cliente
    </a>
  </div>

  <div style="text-align:center;padding:14px;font-size:11px;color:#94A3B8">
    Asesoria Visa Global — ${fecha} — Sistema automatico intake.html
  </div>
</div>`;

    MailApp.sendEmail({
      to: EMAIL_ROBERTO,
      subject: `NUEVO CASO ${numViaj} viajero${numViaj>1?'s':''} — ${nombre} — [${refId}]${esAlerta?' ⚠ REVISAR':''}`,
      htmlBody: html
    });
  } catch(e) {
    console.error('notificarRobertoIntakeFamiliar error:', e.toString());
  }
}

// ── Email al cliente — Confirmación y próximos pasos ─────────────
function notificarClienteIntake(emailCliente, nombre, refId, numViaj, shared) {
  try {
    const primerNombre = nombre.split(' ')[0] || nombre;
    const llegada = shared.intendedArrival || '(por confirmar)';
    const dias    = shared.lengthOfStayDays || '';
    const ciudad  = shared.usStayCity || 'USA';

    const html = `
<div style="font-family:Calibri,Arial,sans-serif;max-width:620px;margin:0 auto;background:#F4F6F9;padding:20px">

  <div style="background:#060E1F;padding:22px 28px;border-radius:12px 12px 0 0;border-bottom:3px solid #F0B429">
    <div style="font-size:12px;color:rgba(255,255,255,.5);margin-bottom:6px">Asesoria Visa Global</div>
    <h1 style="font-size:20px;font-weight:700;color:white;margin:0">Su solicitud fue recibida</h1>
    <div style="font-size:12px;color:#F0B429;margin-top:4px">Referencia: ${refId}</div>
  </div>

  <div style="background:white;border:1px solid #E2E8F0;border-top:none;padding:24px 28px">
    <p style="font-size:14px;color:#1A2940;line-height:1.7;margin-bottom:20px">
      Estimado/a <strong>${primerNombre}</strong>, recibimos correctamente la informacion de su grupo (${numViaj} viajero${numViaj>1?'s':''}). Su asesor Roberto revisara el expediente y en menos de <strong>24 horas</strong> recibira un correo con el analisis personalizado de su caso.
    </p>

    <!-- PASOS INMEDIATOS -->
    <div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:10px;padding:18px;margin-bottom:20px">
      <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#64748B;margin-bottom:14px">Que ocurre ahora</div>

      <div style="display:flex;gap:12px;margin-bottom:14px;align-items:flex-start">
        <div style="width:28px;height:28px;border-radius:50%;background:#060E1F;color:#F0B429;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;flex-shrink:0">1</div>
        <div><strong style="font-size:13px;color:#1A2940">Analisis de perfil (hoy)</strong><br><span style="font-size:12px;color:#64748B">Revisamos toda su informacion y le enviamos el analisis completo con la lista de documentos necesarios para su caso.</span></div>
      </div>

      <div style="display:flex;gap:12px;margin-bottom:14px;align-items:flex-start">
        <div style="width:28px;height:28px;border-radius:50%;background:#060E1F;color:#F0B429;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;flex-shrink:0">2</div>
        <div><strong style="font-size:13px;color:#1A2940">Foto para la visa (esta semana)</strong><br><span style="font-size:12px;color:#64748B">Recibira las especificaciones exactas. Por favor espere nuestras instrucciones antes de sacar la foto.</span></div>
      </div>

      <div style="display:flex;gap:12px;align-items:flex-start">
        <div style="width:28px;height:28px;border-radius:50%;background:#060E1F;color:#F0B429;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;flex-shrink:0">3</div>
        <div><strong style="font-size:13px;color:#1A2940">Documentos de respaldo</strong><br><span style="font-size:12px;color:#64748B">Le pediremos solo los documentos relevantes para su perfil. Iremos paso a paso.</span></div>
      </div>
    </div>

    <!-- SPECS FOTO (importante, darselos ya) -->
    <div style="background:#FEF9EE;border:1px solid #FCD34D;border-left:4px solid #F0B429;border-radius:0 10px 10px 0;padding:14px;margin-bottom:20px">
      <div style="font-size:11px;font-weight:700;color:#92400E;text-transform:uppercase;letter-spacing:.07em;margin-bottom:8px">Especificaciones de la foto para la visa USA</div>
      <div style="font-size:12px;color:#1A2940;line-height:1.8">
        Fondo blanco liso (no gris, no crema)<br>
        Tamano: 5x5 cm (2x2 pulgadas)<br>
        Cara descubierta, frente a la camara, expression neutral<br>
        Sin lentes (ni de sol ni de graduacion)<br>
        Sin gorras, sombreros ni accesorios que cubran la cabeza<br>
        Foto reciente (menos de 6 meses)<br>
        Alta resolucion (300 DPI minimo si es digital)<br>
        <strong style="color:#92400E">Necesitara una foto por cada viajero</strong>
      </div>
    </div>

    <div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:10px;padding:16px;margin-bottom:24px;line-height:1.8">
      <div style="font-size:11px;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:.07em;margin-bottom:8px">Su numero de seguimiento</div>
      <div style="font-size:22px;font-weight:700;color:#166534;letter-spacing:.05em;margin-bottom:6px">${refId}</div>
      <div style="font-size:12px;color:#166534">
        Guarde este numero. Puede escribirnos al WhatsApp +593 99 444 2512 indicando este numero y le daremos el estado de su caso al instante.<br>
        <strong>Ejemplo:</strong> "Hola, mi caso es ${refId}"
      </div>
    </div>
    <div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:10px;padding:14px;margin-bottom:24px;font-size:12px;color:#64748B;line-height:1.6">
      <strong style="color:#1A2940">Viaje estimado:</strong> ${llegada}${dias?' · '+dias+' dias':''} · ${ciudad}
    </div>

    <div style="border-top:1px solid #E2E8F0;padding-top:18px;font-size:12px;color:#64748B;line-height:1.7">
      Cualquier pregunta, escribanos directamente por WhatsApp:<br>
      <a href="https://wa.me/593994442512" style="color:#060E1F;font-weight:700;text-decoration:none;font-size:14px">+593 99 444 2512</a>
      <br><br>
      <strong style="color:#1A2940">Asesoria Visa Global</strong><br>
      Roberto Acosta · asesoriadevisadosglobal.com
    </div>
  </div>

  <div style="text-align:center;padding:14px;font-size:11px;color:#94A3B8">
    Este es un mensaje automatico. Puede responder a este correo o contactarnos por WhatsApp.
  </div>
</div>`;

    MailApp.sendEmail({
      to: emailCliente,
      replyTo: EMAIL_ROBERTO,
      subject: `Solicitud recibida — Visa USA — ${primerNombre} [${refId}]`,
      htmlBody: html
    });
  } catch(e) {
    console.error('notificarClienteIntake error:', e.toString());
  }
}

// ── Handler legacy (portal.html antiguo) ─────────────────────────
function manejarPortalLegacy(payload) {
  try {
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

// ── doGet: reporte HTML del expediente ─────────────────────────────
function doGet(e) {
  // Admin endpoints
  const action = e.parameter.action;
  if (action === 'getCases')           return adminGetCases(e);
  if (action === 'updateCase')         return adminUpdateCase(e);
  if (action === 'newCase')            return adminNewCase(e);
  if (action === 'getExtraction')      return getExtraction(e);
  if (action === 'reanalizar')         return reanalizar(e);
  if (action === 'buscarPorTelefono')  return buscarPorTelefono(e);
  if (action === 'get_draft')          return getDraftAction(e);

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
      max_tokens: 2000,
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

// ── Re-analizar caso con info adicional del asesor ────────────────
function reanalizar(e) {
  if (e.parameter.pin !== ADMIN_PIN) {
    return ContentService.createTextOutput(JSON.stringify({error:'PIN incorrecto'})).setMimeType(ContentService.MimeType.JSON);
  }
  try {
    const refId     = e.parameter.ref;
    const infoExtra = e.parameter.info || '';
    const ss        = SpreadsheetApp.openById(SS_ID);

    // Leer datos existentes del expediente
    const detSheet = ss.getSheetByName('Intake DS-160 Detalle');
    let perfilTexto = '';
    if (detSheet) {
      const rows = detSheet.getDataRange().getValues().filter((r,i) => i > 0 && r[0] === refId);
      perfilTexto = rows.map(r => `${r[2]} — ${r[4]}: ${r[5]}`).join('\n');
    }

    // Leer datos del CRM
    const crmSheet = ss.getSheetByName('CASOS CRM');
    let nombreCliente = refId;
    if (crmSheet) {
      const crmData = crmSheet.getDataRange().getValues();
      const row = crmData.find((r,i) => i > 0 && r[0] === refId);
      if (row) nombreCliente = row[4] || refId;
    }

    const prompt = `Eres el mejor asesor de visas B1/B2 USA del mundo. Re-analiza este caso con la informacion adicional que el asesor acaba de obtener del cliente.

DATOS ORIGINALES DEL EXPEDIENTE (${nombreCliente}):
${perfilTexto.substring(0, 4000)}

INFORMACION ADICIONAL RECABADA POR EL ASESOR:
${infoExtra}

Con esta informacion completa, genera un nuevo analisis. Responde en JSON exacto sin markdown:
{
  "probabilidad": "nueva probabilidad con info completa",
  "probabilidad_sin_estrategia": "sin preparacion",
  "paquete": "ESENCIAL $97 o PROFESIONAL $197 o VIP $397",
  "fuertes": "puntos fuertes separados por punto y coma",
  "debiles": "debilidades separadas por punto y coma",
  "estrategia": "plan actualizado de 7-10 pasos ESPECIFICOS para este caso",
  "documentos_exactos": "documentos especificos con descripcion detallada separados por punto y coma",
  "guion_entrevista": "3-5 preguntas que hara el consul y respuesta ideal. Formato: PREGUNTA: [pregunta] | RESPUESTA IDEAL: [respuesta]",
  "proximos_pasos": "5 acciones urgentes esta semana numeradas",
  "alerta_principal": "riesgo critico o PERFIL LIMPIO",
  "campos_faltantes": "preguntas que aun faltan o PERFIL COMPLETO"
}`;

    const analisis = llamarClaudeIA(prompt);

    // Actualizar CRM con nuevo analisis
    if (crmSheet) {
      const crmData  = crmSheet.getDataRange().getValues();
      const headers  = crmData[0].map(h => String(h));
      const rowIdx   = crmData.findIndex((r,i) => i > 0 && r[0] === refId);
      if (rowIdx > 0) {
        const updates = { 'Probabilidad': analisis.probabilidad||'—', 'Paquete': analisis.paquete||'—' };
        Object.entries(updates).forEach(([campo, valor]) => {
          const col = headers.indexOf(campo);
          if (col >= 0) crmSheet.getRange(rowIdx+1, col+1).setValue(valor);
        });
      }
    }

    // Enviar nuevo email a Roberto
    const fecha = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');
    const htmlReanalisis = `<div style="font-family:Calibri,Arial,sans-serif;max-width:700px;margin:0 auto">
      <div style="background:#7C3AED;color:white;padding:20px 28px;border-radius:12px 12px 0 0">
        <div style="font-size:10px;color:rgba(255,255,255,.6);text-transform:uppercase;letter-spacing:.1em;margin-bottom:6px">RE-ANALISIS CON DATOS ACTUALIZADOS</div>
        <h2 style="font-size:18px;font-weight:700;margin:0">${nombreCliente} — [${refId}]</h2>
      </div>
      <div style="background:white;border:1px solid #E2E8F0;border-top:none;padding:22px 28px">
        <div style="background:#F5F0FF;border:1px solid #C4B5FD;border-radius:10px;padding:14px;margin-bottom:16px">
          <div style="font-size:10px;font-weight:700;color:#6D28D9;text-transform:uppercase;margin-bottom:6px">INFO ADICIONAL QUE PROPORCIONASTE</div>
          <div style="font-size:13px;color:#1A2940;line-height:1.7">${infoExtra.replace(/\n/g,'<br>')}</div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px">
          <div style="text-align:center;background:#F0FDF4;border:1px solid #86EFAC;border-radius:8px;padding:12px">
            <div style="font-size:24px;font-weight:700;color:#166534">${analisis.probabilidad||'—'}</div>
            <div style="font-size:10px;color:#166534;font-weight:600">CON ESTRATEGIA</div>
          </div>
          <div style="text-align:center;background:#FEF2F2;border:1px solid #FCA5A5;border-radius:8px;padding:12px">
            <div style="font-size:24px;font-weight:700;color:#991B1B">${analisis.probabilidad_sin_estrategia||'—'}</div>
            <div style="font-size:10px;color:#991B1B;font-weight:600">SIN PREPARACION</div>
          </div>
        </div>
        <div style="background:#EFF6FF;border:1px solid #93C5FD;border-radius:10px;padding:14px;margin-bottom:12px">
          <div style="font-size:10px;font-weight:700;color:#1E40AF;text-transform:uppercase;margin-bottom:7px">NUEVA ESTRATEGIA</div>
          <div style="font-size:12px;color:#1A2940;line-height:1.8">${(analisis.estrategia||'—').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong>$1</strong>')}</div>
        </div>
        <div style="background:#FEF9EE;border:1px solid #FCD34D;border-radius:10px;padding:14px;margin-bottom:12px">
          <div style="font-size:10px;font-weight:700;color:#92400E;text-transform:uppercase;margin-bottom:7px">DOCUMENTOS EXACTOS</div>
          <div style="font-size:12px;color:#1A2940;line-height:1.8">${(analisis.documentos_exactos||'—').split(';').map((s,i)=>`<strong>${i+1}.</strong> ${s.trim()}`).join('<br>')}</div>
        </div>
        <div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:10px;padding:14px;margin-bottom:12px">
          <div style="font-size:10px;font-weight:700;color:#166534;text-transform:uppercase;margin-bottom:7px">GUION DE ENTREVISTA ACTUALIZADO</div>
          <div style="font-size:12px;color:#1A2940;line-height:1.8">${(analisis.guion_entrevista||'—').replace(/PREGUNTA:/g,'<br><strong style="color:#166534">PREGUNTA:</strong>').replace(/RESPUESTA IDEAL:/g,'<strong style="color:#1E40AF">RESPUESTA IDEAL:</strong>')}</div>
        </div>
        ${analisis.campos_faltantes && !analisis.campos_faltantes.includes('COMPLETO') ? `
        <div style="background:#FEF2F2;border:2px solid #FCA5A5;border-radius:10px;padding:14px;margin-bottom:16px">
          <div style="font-size:10px;font-weight:700;color:#991B1B;text-transform:uppercase;margin-bottom:7px">AUN FALTAN ESTOS DATOS</div>
          <div style="font-size:12px;color:#7F1D1D;line-height:1.9">${(analisis.campos_faltantes||'').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong>$1</strong>')}</div>
        </div>` : '<div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:8px;padding:10px 14px;margin-bottom:16px;font-size:12px;color:#166534;font-weight:600">Perfil ahora completo</div>'}
      </div>
      <div style="text-align:center;padding:12px;font-size:11px;color:#94A3B8">Re-analisis — ${fecha}</div>
    </div>`;

    MailApp.sendEmail({
      to: EMAIL_ROBERTO,
      subject: `RE-ANALISIS ${nombreCliente} — ${analisis.probabilidad||'—'} — [${refId}]`,
      htmlBody: htmlReanalisis
    });

    return ContentService.createTextOutput(JSON.stringify({
      status: 'ok',
      probabilidad: analisis.probabilidad,
      campos_faltantes: analisis.campos_faltantes
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    console.error('reanalizar error:', err.toString());
    return ContentService.createTextOutput(JSON.stringify({error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Extraer datos de DS-160 PDF ───────────────────────────────────
function extractarDS160(payload) {
  try {
    const extractId   = payload.extractId;
    const travelerIdx = payload.travelerIdx || 0;
    const pdfText     = payload.pdfText || '';

    const prompt = `Eres un experto en formularios DS-160 de visa USA para Ecuador. Analiza este texto extraido de un DS-160 completado anteriormente y extrae TODOS los datos que puedas identificar. Devuelve SOLO un JSON valido sin markdown con estos campos (null si no encuentras el dato):
{"surnames":"apellidos en mayusculas como en pasaporte","givenNames":"nombres en mayusculas","dob":"YYYY-MM-DD","cityOfBirth":"ciudad","stateOfBirth":"provincia","countryOfBirth":"Ecuador","sex":"Masculino o Femenino","maritalStatus":"estado civil","nationality":"Ecuador","nationalId":"cedula 10 digitos","passportNumber":"numero pasaporte","passportType":"Ordinario","passportCountry":"Ecuador","passportIssueDate":"YYYY-MM-DD","passportExpiry":"YYYY-MM-DD","homeStreet":"direccion","homeCity":"ciudad residencia","homeProvince":"provincia residencia","primaryPhone":"telefono con codigo","email":"correo@email.com","employmentStatus":"situacion laboral","currentOccupation":"cargo","currentEmployerName":"empresa","currentEmployerCity":"ciudad empresa","monthlySalary":"salario USD","hasBeenInUS":"si o no","usVisitDetails":"detalles visita anterior","hasHadUSVisa":"si o no","previousVisaNumber":"numero visa anterior","previousVisaIssueDate":"YYYY-MM-DD","hasBeenRefused":"si o no","refusalDetails":"motivo rechazo","fatherSurname":"apellido padre","fatherGivenName":"nombre padre","motherSurname":"apellido madre","motherGivenName":"nombre madre","languages":"Espanol","organizations":"Ninguna","countriesVisited5Years":"paises visitados"}

TEXTO DEL DS-160:
${pdfText.substring(0, 6000)}`;

    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': ANTHROPIC_KEY,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
      },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 1500,
        messages: [{ role: 'user', content: prompt }],
      }),
      muteHttpExceptions: true,
    });

    const json     = JSON.parse(resp.getContentText());
    const text     = json.content?.[0]?.text || '{}';
    const clean    = text.replace(/```json|```/g, '').trim();
    const extracted = JSON.parse(clean);

    // Save to Extractions sheet
    const ss    = SpreadsheetApp.openById(SS_ID);
    const sheet = getOrCreateSheet(ss, 'Extractions DS-160', ['Extract ID','Fecha','Viajero','Datos JSON']);
    sheet.appendRow([extractId, new Date().toISOString(), travelerIdx, JSON.stringify(extracted)]);

    return ok({ status: 'ok', extractId });
  } catch(err) {
    console.error('extractarDS160 error:', err.toString());
    return ok({ error: err.toString() });
  }
}

// ── Recuperar extraccion por ID ───────────────────────────────────
function getExtraction(e) {
  try {
    const extractId = e.parameter.extractId;
    const ss        = SpreadsheetApp.openById(SS_ID);
    const sheet     = ss.getSheetByName('Extractions DS-160');
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({status:'pending'})).setMimeType(ContentService.MimeType.JSON);

    const data = sheet.getDataRange().getValues();
    const row  = data.find((r, i) => i > 0 && r[0] === extractId);
    if (!row) return ContentService.createTextOutput(JSON.stringify({status:'pending'})).setMimeType(ContentService.MimeType.JSON);

    return ContentService.createTextOutput(JSON.stringify({
      status: 'ok',
      datos: JSON.parse(row[3])
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Bot WhatsApp: buscar caso por telefono ────────────────────────
function buscarPorTelefono(e) {
  try {
    const telefonoBruto = (e.parameter.telefono || '').replace(/\D/g, '');
    if (!telefonoBruto) return ContentService.createTextOutput(JSON.stringify({status:'sin_caso'})).setMimeType(ContentService.MimeType.JSON);

    const ss    = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('CASOS CRM');
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({status:'sin_caso'})).setMimeType(ContentService.MimeType.JSON);

    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const telIdx  = headers.indexOf('Telefono');

    // Buscar por telefono (comparar solo digitos)
    for (let i = 1; i < data.length; i++) {
      const telSheets = String(data[i][telIdx] || '').replace(/\D/g, '');
      if (telSheets && (telSheets.endsWith(telefonoBruto) || telefonoBruto.endsWith(telSheets))) {
        const caso = {};
        headers.forEach((h, j) => { caso[h] = String(data[i][j] || ''); });

        // Obtener proximos pasos segun estado
        const estado   = caso['Estado'] || '';
        const siguientePaso = obtenerSiguientePaso(estado, caso);

        return ContentService.createTextOutput(JSON.stringify({
          status: 'ok',
          caso,
          siguiente_paso: siguientePaso
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status:'sin_caso'})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error', error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function obtenerSiguientePaso(estado, caso) {
  const e = estado.toLowerCase();
  if (e.includes('enviado') || e.includes('formulario enviado')) {
    return `Pendiente: llenar el formulario de datos en el link que le enviamos. Una vez completado, analizamos su perfil en menos de 24 horas.`;
  }
  if (e.includes('recibido') || e.includes('formulario recibido')) {
    return `Su formulario fue recibido. Estamos analizando su perfil. En menos de 24 horas recibe el analisis completo con los documentos que necesita.`;
  }
  if (e.includes('documento')) {
    const notas = caso['Notas'] || '';
    return `Estamos recopilando sus documentos. ${notas ? 'Nota del asesor: ' + notas : 'Por favor tenga listos los documentos solicitados.'}`;
  }
  if (e.includes('ds-160') || e.includes('ds160')) {
    return `Estamos llenando su formulario DS-160. Le avisamos en cuanto este listo para revisar.`;
  }
  if (e.includes('pago pendiente')) {
    return `Su analisis esta listo. El siguiente paso es confirmar el pago para formalizar la asesoria. Le enviamos el link de pago.`;
  }
  if (e.includes('pago recibido')) {
    return `Pago confirmado. Estamos preparando su expediente completo. Le contactamos pronto con los detalles.`;
  }
  if (e.includes('cita por agendar')) {
    return `Su DS-160 esta listo. El siguiente paso es agendar la cita en el consulado. Le guiamos en ese proceso.`;
  }
  if (e.includes('cita agendada')) {
    const cita = caso['Cita'] || '';
    return `Tiene cita en el consulado${cita && cita !== 'Por agendar' ? ' el ' + cita : ''}. Le enviamos el acceso al simulador de entrevista para prepararse.`;
  }
  if (e.includes('simulador')) {
    return `Tiene acceso al simulador de entrevista en asesoriadevisadosglobal.com/simulador.html. Practique antes de su cita.`;
  }
  if (e.includes('entrevista realizada')) {
    return `Ya realizo su entrevista. Esperamos el resultado del consulado juntos.`;
  }
  if (e.includes('aprobado')) {
    return `Felicitaciones! Su visa fue aprobada. Si necesita asesoria para proximos viajes, con gusto le ayudamos.`;
  }
  return `Su caso esta en proceso. Le contactamos pronto con actualizaciones.`;
}

// ── Admin: GET /getCases ──────────────────────────────────────────
function adminGetCases(e) {
  if (e.parameter.pin !== ADMIN_PIN) {
    return ContentService.createTextOutput(JSON.stringify({error:'PIN incorrecto'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    const ss    = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('CASOS CRM');
    if (!sheet || sheet.getLastRow() < 2) {
      return ContentService.createTextOutput(JSON.stringify({status:'ok', cases:[]}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const cases   = data.slice(1).map((row, i) => {
      const obj = { _row: i + 2 };
      headers.forEach((h, j) => { obj[h] = row[j] !== undefined ? String(row[j]) : ''; });
      return obj;
    });
    return ContentService.createTextOutput(JSON.stringify({status:'ok', cases}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Admin: GET /updateCase ────────────────────────────────────────
function adminUpdateCase(e) {
  if (e.parameter.pin !== ADMIN_PIN) {
    return ContentService.createTextOutput(JSON.stringify({error:'PIN incorrecto'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    const refId    = e.parameter.ref;
    const campo    = e.parameter.campo;   // 'Estado', 'Notas', 'Cita', 'Pago'
    const valor    = e.parameter.valor;
    const ss       = SpreadsheetApp.openById(SS_ID);
    const sheet    = ss.getSheetByName('CASOS CRM');
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({error:'Sin hoja CRM'})).setMimeType(ContentService.MimeType.JSON);

    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const colIdx  = headers.indexOf(campo);
    if (colIdx < 0) return ContentService.createTextOutput(JSON.stringify({error:'Campo no encontrado: '+campo})).setMimeType(ContentService.MimeType.JSON);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === refId) {
        sheet.getRange(i + 1, colIdx + 1).setValue(valor);
        // Color fila según estado
        if (campo === 'Estado') {
          const colores = {
            'Formulario Enviado':    '#FFF9E6',
            'Formulario Recibido':   '#EFF6FF',
            'En Proceso':            '#F5F0FF',
            'Cita Agendada':         '#FFF1F2',
            'Aprobado':              '#F0FDF4',
            'Rechazado':             '#FFF1F2',
            'Cerrado':               '#F4F6F9',
          };
          const bg = colores[valor] || '#FFFFFF';
          sheet.getRange(i + 1, 1, 1, headers.length).setBackground(bg);
        }
        return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({error:'Caso no encontrado'})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Admin: GET /newCase ───────────────────────────────────────────
function adminNewCase(e) {
  if (e.parameter.pin !== ADMIN_PIN) {
    return ContentService.createTextOutput(JSON.stringify({error:'PIN incorrecto'})).setMimeType(ContentService.MimeType.JSON);
  }
  try {
    const ss      = SpreadsheetApp.openById(SS_ID);
    const sheet   = getOrCreateSheet(ss, 'CASOS CRM', COL_CRM);
    const refId   = 'VG-' + Date.now().toString().slice(-6);
    const fecha   = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');
    const nombre  = e.parameter.nombre  || '';
    const tipo    = e.parameter.tipo    || 'USA DS-160';
    const viaj    = e.parameter.viaj    || '1';
    const tel     = e.parameter.tel     || '';
    const email   = e.parameter.email   || '';
    const paquete = e.parameter.paquete || '—';
    const intakeUrl = 'https://www.asesoriadevisadosglobal.com/intake.html?ref=' + refId;
    sheet.appendRow([
      refId, fecha, 'Formulario Enviado', tipo,
      nombre, viaj, tel, email,
      '—', paquete, '—', '—', 'Pendiente', 'Por agendar', '', intakeUrl
    ]);
    formatearCabecera(sheet, 1);

    // Notificar por email
    try {
      const waLink = tel ? 'https://wa.me/593' + tel.replace(/^0/,'').replace(/\D/g,'') : '';
      MailApp.sendEmail({
        to: EMAIL_ROBERTO,
        subject: '📋 Nuevo caso creado — ' + nombre + ' [' + refId + ']',
        htmlBody: `
          <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
            <div style="background:#060E1F;padding:20px 24px;border-radius:8px 8px 0 0">
              <h2 style="color:#F0B429;margin:0;font-size:1.1rem">Nuevo caso creado en el CRM</h2>
            </div>
            <div style="background:#f8faff;padding:24px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 8px 8px">
              <table style="width:100%;border-collapse:collapse;font-size:.9rem">
                <tr><td style="padding:6px 0;color:#718096;width:140px">Referencia</td><td style="padding:6px 0;font-weight:700;color:#060E1F">${refId}</td></tr>
                <tr><td style="padding:6px 0;color:#718096">Cliente</td><td style="padding:6px 0;color:#060E1F">${nombre}</td></tr>
                <tr><td style="padding:6px 0;color:#718096">Tipo visa</td><td style="padding:6px 0;color:#060E1F">${tipo}</td></tr>
                <tr><td style="padding:6px 0;color:#718096">Viajeros</td><td style="padding:6px 0;color:#060E1F">${viaj}</td></tr>
                <tr><td style="padding:6px 0;color:#718096">Teléfono</td><td style="padding:6px 0;color:#060E1F">${tel || '—'}</td></tr>
                <tr><td style="padding:6px 0;color:#718096">Email</td><td style="padding:6px 0;color:#060E1F">${email || '—'}</td></tr>
                <tr><td style="padding:6px 0;color:#718096">Paquete</td><td style="padding:6px 0;color:#060E1F">${paquete}</td></tr>
                <tr><td style="padding:6px 0;color:#718096">Creado</td><td style="padding:6px 0;color:#060E1F">${fecha}</td></tr>
              </table>
              <div style="margin-top:20px;padding:14px;background:#fffbeb;border:1px solid #f0b429;border-radius:6px">
                <div style="font-size:.8rem;color:#78350f;margin-bottom:6px">Link del formulario para el cliente:</div>
                <a href="${intakeUrl}" style="color:#060E1F;font-weight:700;font-size:.85rem">${intakeUrl}</a>
              </div>
              <div style="margin-top:12px;display:flex;gap:10px">
                ${waLink ? `<a href="${waLink}?text=${encodeURIComponent('Hola '+nombre.split(' ')[0]+', le comparto el link para completar su formulario de visa: '+intakeUrl)}" style="display:inline-block;background:#25D366;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:700;font-size:.85rem">Abrir WhatsApp</a>` : ''}
                <a href="https://www.asesoriadevisadosglobal.com/admin.html" style="display:inline-block;background:#060E1F;color:#F0B429;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:700;font-size:.85rem">Ir al CRM</a>
              </div>
            </div>
          </div>`
      });
    } catch(mailErr) { console.log('Email error:', mailErr.toString()); }

    return ContentService.createTextOutput(JSON.stringify({status:'ok', refId, intakeUrl})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
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
