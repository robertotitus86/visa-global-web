// ═══════════════════════════════════════════════════════════════════════
// GOOGLE APPS SCRIPT — Portal Unificado v3.0
// Asesoría Visa Global — Roberto Acosta
// Maneja: USA DS-160 | Schengen | Reino Unido | Caso de Rechazo
//         + intake_familiar (nuevo intake.html multi-viajero DS-160)
// ═══════════════════════════════════════════════════════════════════════

// ANTHROPIC: reservado SOLO para el bot WhatsApp. No gastar aqui.
const ANTHROPIC_KEY = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
// GEMINI: gratis. Usar para todos los analisis del portal.
const GEMINI_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY')
  || 'AIzaSyCphVM6rvGL68pKcdC39v_ikwKOB2VLgx8';
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
    if (payload.subtipo === 'save_draft')               return saveDraftAction(payload);
    if (payload.subtipo === 'intake_familiar')          return manejarIntakeFamiliar(payload);
    if (payload.subtipo === 'extractDS160')             return extractarDS160(payload);
    if (payload.subtipo === 'analizar_ds160_anteriores') return analizarDs160AnterioresAuto(payload);
    if (payload.subtipo === 'submit_screening')         return manejarScreening(payload);
    return manejarPortalLegacy(payload);
  } catch(err) {
    console.error('doPost error:', err.toString());
    return ok({ error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════════════════
// SCREENING — Formulario de calificación de prospectos
// Recibe el payload del screening.html, guarda en Sheets y envía emails
// ═══════════════════════════════════════════════════════════════════════
function manejarScreening(payload) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  const fecha = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');

  // ── 1. RECALCULAR SCORE (verificación servidor) ──────────────────
  const LABORAL = {
    empleado_publico: 35, empleado_privado_plus2: 30, negocio_propio_plus3: 28,
    empleado_privado_menos2: 15, negocio_propio_menos1: 10, desempleado: 0
  };
  const BIEN     = { casa: 15, terreno: 12, vehiculo: 5, ninguno: 0 };
  const INGRESOS = { mas_2000: 10, '1000_2000': 7, '500_999': 3, menos_500: 0 };
  const CREDITO  = { si_hipotecario: 10, si_consumo: 5, no: 0 };
  const CIVIL    = { casado: 10, union_hecho: 8, divorciado_hijos: 6, soltero: 0 };
  const HIJOS    = { dos_o_mas: 10, uno: 7, ninguno: 0 };
  const VIAJES   = { sello_entrada: 20, visa_sin_viajar: 10, ninguno: 0 };
  const RECHAZOS = { nunca: 10, mas_3_anos: 5, menos_1_ano: -5, dos_o_mas: -15 };
  const FAM_IND  = { no: 0, si: -10 };

  const sLab = LABORAL[payload.situacion_laboral] || 0;
  const sFin = (BIEN[payload.bien_principal] || 0)
             + (INGRESOS[payload.ingresos_rango] || 0)
             + (CREDITO[payload.credito_activo] || 0);
  const sFam = (CIVIL[payload.estado_civil] || 0)
             + (HIJOS[payload.hijos_ecuador] || 0);
  const sHis = (VIAJES[payload.viajes_previos] || 0)
             + (RECHAZOS[payload.historial_rechazos] || 0)
             + (FAM_IND[payload.familiares_indocumentados] || 0);

  const scoreRaw   = sLab + sFin + sFam + sHis;
  const scoreTotal = Math.max(0, Math.min(100, Math.round((scoreRaw / 100) * 100)));

  let ruta, riesgo, prioridad;
  if      (scoreTotal >= 70) { ruta = 'C'; riesgo = 'BAJO';  prioridad = 'HOT';     }
  else if (scoreTotal >= 40) { ruta = 'B'; riesgo = 'MEDIO'; prioridad = 'WARM';    }
  else                       { ruta = 'A'; riesgo = 'ALTO';  prioridad = 'NURTURE'; }

  // ── 2. GUARDAR EN HOJA "Prospectos Screening" ────────────────────
  const COLS = [
    'Fecha','Nombre','WhatsApp','Email','Visa','Laboral','Bien','Ingresos',
    'Crédito','Civil','Hijos','Viajes','Rechazos','Fam Indoc',
    'Score Laboral','Score Financiero','Score Familiar','Score Historial',
    'Score Total','Ruta','Riesgo 214b','Prioridad','Fuente','Estado'
  ];
  let sh = ss.getSheetByName('Prospectos Screening');
  if (!sh) {
    sh = ss.insertSheet('Prospectos Screening');
    sh.appendRow(COLS);
    sh.getRange(1, 1, 1, COLS.length)
      .setBackground('#060E1F').setFontColor('#F0B429').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.appendRow([
    fecha,
    payload.nombre   || '',
    payload.whatsapp || '',
    payload.email    || '',
    payload.visa_tipo || '',
    payload.situacion_laboral || '',
    payload.bien_principal    || '',
    payload.ingresos_rango    || '',
    payload.credito_activo    || '',
    payload.estado_civil      || '',
    payload.hijos_ecuador     || '',
    payload.viajes_previos    || '',
    payload.historial_rechazos || '',
    payload.familiares_indocumentados || '',
    sLab, sFin, sFam, sHis,
    scoreTotal, ruta, riesgo, prioridad,
    payload.utm_ref || 'directo',
    'Nuevo'
  ]);

  // ── 3. EMAIL A ROBERTO (alerta de nuevo prospecto) ───────────────
  const visaNombre = payload.visa_tipo === 'USA_B1B2' ? 'USA (B1/B2)' : 'Schengen';
  const emojiRuta  = ruta === 'C' ? '🟢' : ruta === 'B' ? '🟡' : '🔴';
  const asuntoRob  = `${emojiRuta} Nuevo prospecto ${ruta} — ${payload.nombre || 'sin nombre'} — Score ${scoreTotal}/100`;
  const cuerpoRob  =
    `Nuevo lead calificado desde el formulario de screening.\n\n` +
    `━━━━━━━━━━━━━━━━━━━━━━\n` +
    `NOMBRE:   ${payload.nombre}\n` +
    `WHATSAPP: ${payload.whatsapp}\n` +
    `EMAIL:    ${payload.email}\n` +
    `VISA:     ${visaNombre}\n` +
    `━━━━━━━━━━━━━━━━━━━━━━\n` +
    `SCORE TOTAL:  ${scoreTotal}/100\n` +
    `RUTA:         ${ruta} (${prioridad})\n` +
    `RIESGO 214b:  ${riesgo}\n` +
    `━━━━━━━━━━━━━━━━━━━━━━\n` +
    `DESGLOSE:\n` +
    `  Laboral:    ${sLab} pts  (${payload.situacion_laboral})\n` +
    `  Financiero: ${sFin} pts  (bien: ${payload.bien_principal} / ingresos: ${payload.ingresos_rango})\n` +
    `  Familiar:   ${sFam} pts  (${payload.estado_civil} / hijos: ${payload.hijos_ecuador})\n` +
    `  Historial:  ${sHis} pts  (viajes: ${payload.viajes_previos} / rechazos: ${payload.historial_rechazos})\n` +
    `━━━━━━━━━━━━━━━━━━━━━━\n` +
    `ACCIÓN RECOMENDADA:\n` +
    (ruta === 'C'
      ? '→ CIERRE DIRECTO. Escríbele hoy mismo — perfil ideal, tiempo es dinero.'
      : ruta === 'B'
      ? '→ VENDER CONSULTA ESTRATÉGICA. Tiene viabilidad pero necesita armado de expediente.'
      : '→ NUTRIR. Enviar guía de mejora de perfil. No gastar tiempo de asesoría.') +
    `\n\nVer todos los prospectos en Google Sheets → Hoja "Prospectos Screening"`;

  try {
    GmailApp.sendEmail(EMAIL_ROBERTO, asuntoRob, cuerpoRob);
  } catch(e) {
    console.error('Error email Roberto:', e.toString());
  }

  // ── 4. EMAIL AL PROSPECTO ─────────────────────────────────────────
  const emailProspecto = payload.email || '';
  if (emailProspecto) {
    const EMAILS = {
      A: {
        asunto: `Tu evaluación de perfil — pasos para fortalecer tu arraigo`,
        cuerpo: `Hola ${payload.nombre || ''},\n\n` +
          `Revisamos tu perfil para la ${visaNombre} y tu score de viabilidad es ${scoreTotal}/100.\n\n` +
          `Tu perfil tiene un riesgo elevado de rechazo bajo la Sección 214(b) en este momento. ` +
          `Esto no significa que no puedas obtener la visa — significa que necesitas fortalecer ` +
          `tu arraigo antes de aplicar para no desperdiciar la tasa consular ($185).\n\n` +
          `Hay acciones concretas que puedes tomar para mejorar tu perfil:\n\n` +
          `• Fortalecer tu historial laboral documentado (mínimo 1 año continuo)\n` +
          `• Demostrar vínculos financieros sólidos (cuentas, propiedades, créditos)\n` +
          `• Consolidar vínculos familiares verificables en Ecuador\n\n` +
          `Cuando tu perfil esté más sólido, contáctanos para una evaluación completa. ` +
          `Tu primera consulta siempre será gratis.\n\n` +
          `WhatsApp: +593 98 784 6751\n` +
          `Web: asesoriadevisadosglobal.com\n\n` +
          `— Asesoría Visa Global`
      },
      B: {
        asunto: `Tu perfil tiene viabilidad — necesitamos estructurar tu expediente`,
        cuerpo: `Hola ${payload.nombre || ''},\n\n` +
          `Analizamos tu perfil para la ${visaNombre} y tu score de viabilidad es ${scoreTotal}/100.\n\n` +
          `Tu caso tiene puntos fuertes pero también vulnerabilidades que un oficial consular ` +
          `puede usar en tu contra. Los casos como el tuyo no se aprueban solos — ` +
          `requieren una estrategia de expediente bien estructurada y preparación específica.\n\n` +
          `Podemos trabajar en eso juntos. La primera consulta es completamente gratis.\n\n` +
          `Escríbenos por WhatsApp para agendar:\n` +
          `→ wa.me/593987846751\n\n` +
          `O visítanos en: asesoriadevisadosglobal.com\n\n` +
          `— Asesoría Visa Global`
      },
      C: {
        asunto: `Tu perfil está listo para aplicar — iniciemos tu trámite`,
        cuerpo: `Hola ${payload.nombre || ''},\n\n` +
          `Excelente noticia: tu score de viabilidad es ${scoreTotal}/100 para la ${visaNombre}.\n\n` +
          `Tu perfil de arraigo cumple con los criterios que evalúan los consulados. ` +
          `Los tiempos de espera para citas consulares están aumentando — ` +
          `te conviene iniciar el proceso cuanto antes para asegurar tu fecha.\n\n` +
          `Podemos iniciar esta semana. Escríbenos:\n\n` +
          `WhatsApp: wa.me/593987846751\n` +
          `Email: info@asesoriadevisadosglobal.com\n\n` +
          `Primera consulta gratis. Empezamos cuando tú quieras.\n\n` +
          `— Asesoría Visa Global`
      }
    };

    try {
      const em = EMAILS[ruta];
      GmailApp.sendEmail(emailProspecto, em.asunto, em.cuerpo, {
        name: 'Asesoría Visa Global',
        replyTo: 'info@asesoriadevisadosglobal.com'
      });
    } catch(e) {
      console.error('Error email prospecto:', e.toString());
    }
  }

  return ok({ ok: true, ruta, score: scoreTotal, riesgo });
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
  const CAMPOS_INTERNOS = ['_pendientes']; // campos del sistema, no guardar en Sheets
  // Shared info
  Object.entries(shared).forEach(([campo, valor]) => {
    if (CAMPOS_INTERNOS.includes(campo)) return;
    if (valor && String(valor).trim()) {
      sheetDet.appendRow([refId, fecha, 'COMPARTIDO', '—', campo, String(valor)]);
    }
  });
  // Per-traveler data
  personas.forEach(p => {
    const datos = p.datos || {};
    Object.entries(datos).forEach(([campo, valor]) => {
      if (CAMPOS_INTERNOS.includes(campo)) return;
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

  // 2. Recopilar campos pendientes del formulario (dejados en blanco)
  const pendientesFormulario = [];
  const pendShared = shared._pendientes || [];
  if (pendShared.length) {
    pendShared.forEach(p => pendientesFormulario.push('DATOS DEL VIAJE — ' + p));
  }
  personas.forEach(p => {
    const pend = (p.datos || {})._pendientes || [];
    pend.forEach(f => pendientesFormulario.push((p.nombre || 'Viajero') + ' — ' + f));
  });

  // 3. Análisis IA familiar
  const analisis = analizarPerfilFamiliar(shared, personas, refId);
  // Inyectar pendientes del formulario al analisis para el email
  if (pendientesFormulario.length) {
    const listaPend = pendientesFormulario.map((p, i) => `${i+1}. ${p}`).join('\n');
    analisis.campos_faltantes_formulario = listaPend;
  }

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

  // AUTO: Generar y guardar guía DS-160
  try {
    const guiaDS160 = generarDS160PreFill(personas, shared, analisis);
    guardarGuiaDS160(refId, guiaDS160, analisis);
  } catch(e) {
    console.error('Error generando guia DS-160:', e.toString());
  }

  // AUTO: Generar y guardar carta del cliente
  try {
    const cartaHTML = generarCartaHTML(refId, personas, shared, analisis);
    guardarCarta(refId, cartaHTML);
  } catch(e) {
    console.error('Error generando carta:', e.toString());
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
  "consulado": "GUAYAQUIL o QUITO — razon concreta para este perfil (ciudad de residencia, disponibilidad de citas, historial del consulado con este tipo de caso)",
  "motivo_ds160": "texto EXACTO en espanol para poner en el campo purpose of trip del DS-160 — 2-3 oraciones especificas, con actividades concretas y fechas. NO usar 'turismo' a secas. Ejemplo: 'Visita turistica familiar para conocer Orlando (parques Disney y Universal) y Miami Beach del [fecha] al [fecha]. El grupo incluye menores de edad que desean conocer los parques tematicos. Regresamos a Ecuador donde los ninos continuan sus estudios.'",
  "fecha_viaje_ideal": "mejor epoca del ano para ESTE perfil especifico y por que (considerar: temporada de menor rechazo en el consulado, costos de vuelo, actividad en destinos mencionados, edad de menores si aplica)",
  "fecha_cita_sugerida": "cuando reservar la cita consular — dia/semana especifica si es posible — cuanto tiempo antes del viaje ideal — y por que ese timing le da margen para reunir documentos",
  "tiempo_preparacion": "cuantas semanas necesita este caso especifico desde hoy hasta estar listo para la cita, paso a paso: semana 1 hacer X, semana 2 hacer Y...",
  "fuertes": "3-5 puntos fuertes del perfil separados por punto y coma",
  "debiles": "2-4 debilidades reales que pueden causar rechazo separadas por punto y coma",
  "estrategia": "plan CONCRETO de 8-12 pasos numerados para llevar este caso al 80-100%. Especifico para ESTE caso: que documentos preparar exactamente, como presentarlos, que argumentar para cada punto debil, que enfatizar del perfil",
  "documentos_exactos": "lista de 6-10 documentos especificos con descripcion exacta de como deben estar preparados (formato, quien firma, que debe decir, que NO debe decir), separados por punto y coma",
  "checklist_pre_cita": "lista de 10-15 items que deben estar listos ANTES de ir al consulado, en orden de urgencia. Incluir: documentos fisicos, copias, fotos, comprobantes, DS-160 impreso, confirmacion de cita, etc.",
  "guion_entrevista": "4-6 preguntas que el consul hara SEGURO a ESTE perfil especifico y la respuesta ideal para cada una. Formato: PREGUNTA: [pregunta] | RESPUESTA IDEAL: [respuesta exacta de 1-3 frases]",
  "proximos_pasos": "5-7 acciones concretas que el asesor debe hacer ESTA SEMANA con este cliente, en orden de urgencia, numeradas",
  "alerta_principal": "el riesgo mas critico de este caso y como neutralizarlo. Si no hay riesgo critico escribir PERFIL LIMPIO",
  "campos_faltantes": "preguntas ESPECIFICAS que el asesor debe hacerle al cliente porque no estan en el formulario. Formato: '1. [PERSONA] — [PREGUNTA EXACTA]'. Si el perfil esta completo escribir PERFIL COMPLETO"
}`;
}

// ── Email a Roberto — Intake Familiar ────────────────────────────
function notificarRobertoIntakeFamiliar(refId, fecha, nombre, numViaj, tel, email, shared, personas, analisis, reporteUrl) {
  try {
    const esAlerta = analisis.alerta_principal && !analisis.alerta_principal.includes('LIMPIO');
    const colorAlerta = esAlerta ? '#7C3AED' : '#060E1F';

    const viajerosList = personas.map((p, i) => {
      const d = p.datos || {};
      return `<tr>
        <td style="padding:6px 8px;font-size:12px;border-bottom:1px solid #F4F6F9;font-weight:600">${p.nombre}</td>
        <td style="padding:6px 8px;font-size:11px;border-bottom:1px solid #F4F6F9;color:#64748B">${p.rol}</td>
        <td style="padding:6px 8px;font-size:11px;border-bottom:1px solid #F4F6F9;color:#64748B">${d.passportNumber||'Sin pasaporte'}</td>
        <td style="padding:6px 8px;font-size:11px;border-bottom:1px solid #F4F6F9;color:#64748B">${d.employmentStatus||'—'}</td>
        <td style="padding:6px 8px;font-size:11px;border-bottom:1px solid #F4F6F9">${(p.banderas_seguridad||[]).length?'<span style="color:#ef4444;font-weight:700">'+p.banderas_seguridad.length+' alerta(s)</span>':'<span style="color:#22c55e">Limpio</span>'}</td>
      </tr>`;
    }).join('');

    const blk = (color,label,content) =>
      `<div style="background:${color};border-radius:10px;padding:14px;margin-bottom:12px">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#1A2940;margin-bottom:8px;opacity:.7">${label}</div>
        <div style="font-size:12px;color:#1A2940;line-height:1.85">${content}</div>
      </div>`;

    const html = `
<div style="font-family:Calibri,Arial,sans-serif;max-width:720px;margin:0 auto;background:#F4F6F9;padding:20px">

  <!-- CABECERA -->
  <div style="background:${colorAlerta};color:white;padding:22px 28px;border-radius:12px 12px 0 0">
    <div style="font-size:10px;font-weight:700;letter-spacing:.1em;color:rgba(255,255,255,.6);text-transform:uppercase;margin-bottom:6px">EXPEDIENTE VISA USA B1/B2 — FAMILIA</div>
    <h1 style="font-size:22px;font-weight:700;margin:0 0 6px">${nombre} + ${numViaj-1} acompañante${numViaj-1!==1?'s':''}</h1>
    <div style="font-size:13px;opacity:.8">Ref: ${refId} · Llegada: ${shared.intendedArrival||'por definir'} · ${shared.lengthOfStayDays||'?'} dias · ${shared.usStayCity||'?'}, ${shared.usStayState||'?'}</div>
    <div style="margin-top:12px;display:inline-block;background:rgba(255,255,255,.15);border-radius:6px;padding:6px 14px;font-size:18px;font-weight:700">${analisis.probabilidad||'—'} aprobacion &nbsp;→&nbsp; ${analisis.paquete||'—'}</div>
  </div>

  <div style="background:white;border:1px solid #E2E8F0;border-top:none;padding:24px 28px">

    <!-- ALERTA -->
    ${esAlerta ? `<div style="background:#FEF2F2;border:2px solid #FCA5A5;border-radius:10px;padding:14px;margin-bottom:18px">
      <div style="font-size:11px;font-weight:700;color:#991B1B;text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px">ALERTA CRITICA — ATENDER PRIMERO</div>
      <div style="font-size:13px;color:#7F1D1D;line-height:1.6">${analisis.alerta_principal||''}</div>
    </div>` : `<div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:10px;padding:10px 14px;margin-bottom:18px;font-size:12px;color:#166534;font-weight:600">Perfil limpio — sin alertas criticas</div>`}

    <!-- VIAJEROS -->
    <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#64748B;margin-bottom:8px">Viajeros del grupo</div>
    <table style="width:100%;border-collapse:collapse;margin-bottom:20px;font-size:12px;background:#F8FAFC;border-radius:8px;overflow:hidden">
      <tr style="background:#E2E8F0">
        <td style="padding:6px 8px;font-weight:700;color:#1A2940">Nombre</td>
        <td style="padding:6px 8px;font-weight:700;color:#1A2940">Rol</td>
        <td style="padding:6px 8px;font-weight:700;color:#1A2940">Pasaporte</td>
        <td style="padding:6px 8px;font-weight:700;color:#1A2940">Empleo</td>
        <td style="padding:6px 8px;font-weight:700;color:#1A2940">Seguridad</td>
      </tr>
      ${viajerosList}
    </table>

    <!-- BLOQUE 1: CONSULADO Y LOGISTICA -->
    <div style="background:#EFF6FF;border:1px solid #93C5FD;border-radius:10px;padding:16px;margin-bottom:12px">
      <div style="font-size:11px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em;margin-bottom:10px">CONSULADO Y LOGISTICA</div>
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <tr><td style="padding:5px 0;color:#64748B;width:35%;vertical-align:top">Consulado recomendado</td><td style="padding:5px 0;font-weight:600;color:#1A2940">${analisis.consulado||'—'}</td></tr>
        <tr><td style="padding:5px 0;color:#64748B;vertical-align:top">Fecha viaje ideal</td><td style="padding:5px 0;font-weight:600;color:#1A2940">${analisis.fecha_viaje_ideal||'—'}</td></tr>
        <tr><td style="padding:5px 0;color:#64748B;vertical-align:top">Cuando reservar cita</td><td style="padding:5px 0;font-weight:600;color:#1A2940">${analisis.fecha_cita_sugerida||'—'}</td></tr>
        <tr><td style="padding:5px 0;color:#64748B;vertical-align:top">Tiempo de preparacion</td><td style="padding:5px 0;color:#1A2940">${analisis.tiempo_preparacion||'—'}</td></tr>
      </table>
    </div>

    <!-- BLOQUE 2: MOTIVO EXACTO DS-160 -->
    <div style="background:#FFF7ED;border:2px solid #F0B429;border-radius:10px;padding:16px;margin-bottom:12px">
      <div style="font-size:11px;font-weight:700;color:#92400E;text-transform:uppercase;letter-spacing:.07em;margin-bottom:8px">TEXTO EXACTO PARA DS-160 — CAMPO "PURPOSE OF TRIP"</div>
      <div style="font-size:13px;color:#78350F;line-height:1.7;font-style:italic;">"${analisis.motivo_ds160||'—'}"</div>
      <div style="font-size:11px;color:#92400E;margin-top:8px">Copiar este texto exactamente en el DS-160 de cada adulto del grupo.</div>
    </div>

    <!-- BLOQUE 3: PROBABILIDAD -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px">
      <div style="text-align:center;background:#F0FDF4;border:1px solid #86EFAC;border-radius:10px;padding:14px">
        <div style="font-size:28px;font-weight:700;color:#166534">${analisis.probabilidad||'—'}</div>
        <div style="font-size:10px;color:#166534;font-weight:700;text-transform:uppercase">Con estrategia aplicada</div>
      </div>
      <div style="text-align:center;background:#FEF2F2;border:1px solid #FCA5A5;border-radius:10px;padding:14px">
        <div style="font-size:28px;font-weight:700;color:#991B1B">${analisis.probabilidad_sin_estrategia||'—'}</div>
        <div style="font-size:10px;color:#991B1B;font-weight:700;text-transform:uppercase">Sin preparacion</div>
      </div>
    </div>

    <!-- BLOQUE 4: PUNTOS FUERTES Y DEBILES -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px">
      ${blk('#F0FDF4','Puntos fuertes del perfil',(analisis.fuertes||'—').split(';').map(s=>'<div style="padding:3px 0;border-bottom:1px solid #DCFCE7">✓ '+s.trim()+'</div>').join(''))}
      ${blk('#FEF2F2','Puntos debiles a neutralizar',(analisis.debiles||'—').split(';').map(s=>'<div style="padding:3px 0;border-bottom:1px solid #FEE2E2;color:#991B1B">✗ '+s.trim()+'</div>').join(''))}
    </div>

    <!-- BLOQUE 5: ESTRATEGIA COMPLETA -->
    <div style="background:#EFF6FF;border:1px solid #93C5FD;border-radius:10px;padding:16px;margin-bottom:12px">
      <div style="font-size:11px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em;margin-bottom:10px">ESTRATEGIA COMPLETA — PLAN DE ACCION</div>
      <div style="font-size:12px;color:#1A2940;line-height:1.9">${(analisis.estrategia||'—').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong style="color:#1E40AF">$1</strong>')}</div>
    </div>

    <!-- BLOQUE 6: DOCUMENTOS EXACTOS -->
    <div style="background:#FEF9EE;border:1px solid #FCD34D;border-radius:10px;padding:16px;margin-bottom:12px">
      <div style="font-size:11px;font-weight:700;color:#92400E;text-transform:uppercase;letter-spacing:.07em;margin-bottom:10px">DOCUMENTOS EXACTOS A PREPARAR</div>
      <div style="font-size:12px;color:#1A2940;line-height:1.8">${(analisis.documentos_exactos||analisis.documentos_clave||'—').split(';').map((s,i)=>`<div style="padding:6px 0;border-bottom:1px solid #FEF3C7"><strong style="color:#92400E">${i+1}.</strong> ${s.trim()}</div>`).join('')}</div>
    </div>

    <!-- BLOQUE 7: CHECKLIST PRE-CITA -->
    <div style="background:#F8F9FF;border:1px solid #C7D2FE;border-radius:10px;padding:16px;margin-bottom:12px">
      <div style="font-size:11px;font-weight:700;color:#3730A3;text-transform:uppercase;letter-spacing:.07em;margin-bottom:10px">CHECKLIST — LO QUE DEBEN LLEVAR EL DIA DE LA CITA</div>
      <div style="font-size:12px;color:#1A2940;line-height:1.9">${(analisis.checklist_pre_cita||'—').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong style="color:#3730A3">$1</strong>')}</div>
    </div>

    <!-- BLOQUE 8: GUION ENTREVISTA -->
    <div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:10px;padding:16px;margin-bottom:12px">
      <div style="font-size:11px;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:.07em;margin-bottom:10px">GUION DE ENTREVISTA — PREGUNTAS Y RESPUESTAS IDEALES</div>
      <div style="font-size:12px;color:#1A2940;line-height:1.9">${(analisis.guion_entrevista||'—').replace(/PREGUNTA:/g,'<br><strong style="color:#166534">PREGUNTA:</strong>').replace(/RESPUESTA IDEAL:/g,'<strong style="color:#1E40AF">RESPUESTA IDEAL:</strong>')}</div>
    </div>

    <!-- BLOQUE 9: PROXIMOS PASOS -->
    <div style="background:#EFF6FF;border:1px solid #93C5FD;border-radius:10px;padding:16px;margin-bottom:16px">
      <div style="font-size:11px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em;margin-bottom:10px">PROXIMOS PASOS — ESTA SEMANA (EN ORDEN DE URGENCIA)</div>
      <div style="font-size:12px;color:#1A2940;line-height:1.9">${(analisis.proximos_pasos||'—').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong style="color:#1E40AF">$1</strong>')}</div>
    </div>

    <!-- CAMPOS FALTANTES -->
    ${analisis.campos_faltantes_formulario ? `
    <div style="background:#FFF7ED;border:2px solid #F0B429;border-radius:10px;padding:16px;margin-bottom:12px">
      <div style="font-size:11px;font-weight:700;color:#92400E;text-transform:uppercase;letter-spacing:.07em;margin-bottom:8px">INFORMACION PENDIENTE — SOLICITAR AL CLIENTE</div>
      <div style="font-size:12px;color:#78350F;line-height:1.9">${analisis.campos_faltantes_formulario.replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong>$1</strong>')}</div>
    </div>` : ''}
    ${analisis.campos_faltantes && !analisis.campos_faltantes.includes('COMPLETO') ? `
    <div style="background:#FEF2F2;border:1px solid #FCA5A5;border-radius:10px;padding:14px;margin-bottom:16px">
      <div style="font-size:11px;font-weight:700;color:#991B1B;text-transform:uppercase;letter-spacing:.07em;margin-bottom:8px">DATOS ADICIONALES IDENTIFICADOS POR IA</div>
      <div style="font-size:12px;color:#7F1D1D;line-height:1.9">${(analisis.campos_faltantes||'').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong>$1</strong>')}</div>
    </div>` : ''}

    <!-- BOTONES -->
    <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:8px">
    <a href="${reporteUrl}" style="display:inline-block;background:#060E1F;color:white;padding:13px 28px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px;margin-right:10px">
      Ver expediente completo
    </a>
    <a href="https://wa.me/${(tel||'').replace(/\D/g,'')}" style="display:inline-block;background:#22c55e;color:white;padding:13px 24px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px">
      WhatsApp cliente
    </a>
    <a href="${WEBHOOK}?action=reconstruct&ref=${refId}" style="display:inline-block;background:#F0B429;color:#060E1F;padding:13px 24px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px">
      Ver formulario completo
    </a>
    </div>

  </div>

  <div style="text-align:center;padding:14px;font-size:11px;color:#94A3B8">
    Asesoria Visa Global · ${fecha} · Ref: ${refId}
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
  if (action === 'reconstruct')        return reconstructFromDetalle(e);
  if (action === 'test')              return testSistema(e);
  if (action === 'ds160guide') return servirGuiaDS160(e.parameter.ref || '', e.parameter.pin || '');
  if (action === 'carta')     return servirCarta(e.parameter.ref || '', e.parameter.pin || '');

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

// ── Llamar a Gemini (gratis) ──────────────────────────────────────
function llamarClaudeIA(prompt) {
  // Usa Gemini en lugar de Anthropic — gratis, sin consumir creditos pagos
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_KEY}`;
    const resp = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { maxOutputTokens: 4000, temperature: 0.3 }
      }),
      muteHttpExceptions: true
    });

    const raw  = resp.getContentText();
    const json = JSON.parse(raw);
    if (json.error) throw new Error('Gemini error: ' + JSON.stringify(json.error));

    const text  = (json.candidates?.[0]?.content?.parts?.[0]?.text || '').trim();
    const clean = text.replace(/```json|```/g, '').trim();
    const match = clean.match(/\{[\s\S]*\}/);
    if (!match) throw new Error('No JSON en respuesta Gemini');

    return JSON.parse(match[0]);
  } catch(err) {
    console.error('llamarGemini error:', err.toString());
    return {
      probabilidad: 'Pendiente', probabilidad_sin_estrategia: '—',
      paquete: 'Por definir', razon_paquete: '—',
      fuertes: 'Analisis pendiente', debiles: '—',
      estrategia: 'Error: ' + err.toString(),
      documentos_exactos: '—', guion_entrevista: '—',
      proximos_pasos: 'Verificar configuracion de Gemini en Script Properties',
      alerta_principal: 'Error en analisis — revisar logs',
      campos_faltantes: 'PERFIL COMPLETO'
    };
  }
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
// ── TEST SISTEMA ─────────────────────────────────────────────────
function testSistema(e) {
  const resultados = {};
  // 1. Verificar ANTHROPIC_KEY
  resultados.anthropic_key = ANTHROPIC_KEY ? 'CONFIGURADA (' + ANTHROPIC_KEY.substring(0,20) + '...)' : 'NO CONFIGURADA';
  // 2. Verificar Spreadsheet
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    resultados.spreadsheet = 'OK — ' + ss.getName();
    resultados.hojas = ss.getSheets().map(s => s.getName()).join(', ');
  } catch(e2) { resultados.spreadsheet = 'ERROR: ' + e2.toString(); }
  // 3. Enviar email de prueba
  try {
    MailApp.sendEmail({
      to: EMAIL_ROBERTO,
      subject: 'TEST SISTEMA — Visa Global Apps Script funcionando',
      htmlBody: `<div style="font-family:Arial;padding:20px;background:#f4f6f9">
        <h2 style="color:#060E1F">Sistema funcionando correctamente</h2>
        <p>Este email confirma que Apps Script puede enviar emails.</p>
        <p><strong>ANTHROPIC_KEY:</strong> ${resultados.anthropic_key}</p>
        <p><strong>Spreadsheet:</strong> ${resultados.spreadsheet}</p>
        <p><strong>Hojas:</strong> ${resultados.hojas||'—'}</p>
        <p style="color:#64748B;font-size:12px">Visa Global · ${new Date().toISOString()}</p>
      </div>`
    });
    resultados.email = 'ENVIADO a ' + EMAIL_ROBERTO;
  } catch(e3) { resultados.email = 'ERROR: ' + e3.toString(); }

  return ContentService
    .createTextOutput(JSON.stringify(resultados, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ANALIZADOR DE DS-160 ANTERIORES ─────────────────────────────
// Recibe PDFs de formularios anteriores, extrae todo y genera estrategia completa
function analizarDs160Anteriores(payload) {
  try {
    const ref      = payload.ref      || ('DS-' + Date.now().toString().slice(-6));
    const caseName = payload.caseName || '';
    const personas = payload.personas || [];
    const fecha    = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');
    if (!personas.length) return ok({ error: 'Sin personas' });

    // 1. Extraer datos de cada PDF con Claude
    const perfiles = personas.map(p => {
      const extractPrompt = `Eres un experto en formularios DS-160 de visa USA. Extrae TODOS los datos de este formulario DS-160 anterior. Devuelve JSON exacto sin markdown:
{
  "apellidos":"","nombres":"","fecha_nacimiento":"YYYY-MM-DD","sexo":"Masculino o Femenino",
  "estado_civil":"","ciudad_nacimiento":"","provincia_nacimiento":"","pais_nacimiento":"",
  "nacionalidad":"","cedula":"","numero_pasaporte":"","tipo_pasaporte":"Ordinario",
  "fecha_emision_pasaporte":"YYYY-MM-DD","fecha_vencimiento_pasaporte":"YYYY-MM-DD",
  "ciudad_emision_pasaporte":"","direccion":"","ciudad_residencia":"","provincia_residencia":"",
  "telefono":"","email":"","situacion_laboral":"","cargo":"","empleador":"",
  "direccion_empleador":"","ciudad_empleador":"","telefono_empleador":"","fecha_inicio_empleo":"YYYY-MM-DD",
  "salario_mensual":"","funciones":"","empleos_anteriores":[{"empresa":"","cargo":"","ciudad":"","desde":"","hasta":"","razon_salida":""}],
  "nombre_padre":"","apellidos_padre":"","fecha_nac_padre":"","padre_en_usa":"",
  "nombre_madre":"","apellidos_madre":"","fecha_nac_madre":"","madre_en_usa":"",
  "nombre_conyuge":"","apellidos_conyuge":"","fecha_nac_conyuge":"","nacionalidad_conyuge":"",
  "ha_estado_en_usa":"si o no","detalle_visitas_usa":"","ha_tenido_visa_usa":"si o no",
  "numero_visa_anterior":"","fecha_emision_visa_anterior":"","tipo_visa_anterior":"",
  "le_han_negado_visa":"si o no","detalle_negacion":"",
  "paises_visitados_5_anos":"","idiomas":"","organizaciones":"",
  "redes_sociales":[{"plataforma":"","usuario":""}],
  "proposito_viaje_anterior":"","destino_anterior":"","duracion_anterior":"",
  "contacto_usa_nombre":"","contacto_usa_relacion":"","contacto_usa_telefono":"","contacto_usa_direccion":"",
  "alojamiento_usa":"","ciudad_destino_usa":"","estado_destino_usa":"",
  "quienpaga":"","tiene_servicio_militar":"","nivel_educacion":""
}
TEXTO DEL DS-160:
${(p.pdfText || '').substring(0, 8000)}`;

      try {
        const gUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_KEY}`;
        const resp = UrlFetchApp.fetch(gUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          payload: JSON.stringify({ contents:[{parts:[{text: extractPrompt}]}], generationConfig:{maxOutputTokens:2000,temperature:0.2} }),
          muteHttpExceptions: true
        });
        const j = JSON.parse(resp.getContentText());
        const txt = (j.candidates?.[0]?.content?.parts?.[0]?.text || '{}').replace(/```json|```/g,'').trim();
        const m = txt.match(/\{[\s\S]*\}/);
        return { nombre: p.nombre, datos: m ? JSON.parse(m[0]) : {}, hayTexto: (p.pdfText||'').length > 100 };
      } catch(e) {
        return { nombre: p.nombre, datos: {}, hayTexto: false };
      }
    });

    // 2. Análisis experto completo de toda la familia
    const resumenPerfiles = perfiles.map((p, i) => {
      const d = p.datos;
      return `VIAJERO ${i+1}: ${p.nombre}
  Datos personales: ${d.apellidos||'?'}, ${d.nombres||'?'} | Nac: ${d.fecha_nacimiento||'?'} | ${d.sexo||'?'} | ${d.estado_civil||'?'}
  Pasaporte: ${d.numero_pasaporte||'?'} vence ${d.fecha_vencimiento_pasaporte||'?'}
  Empleo: ${d.situacion_laboral||'?'} — ${d.cargo||'?'} en ${d.empleador||'?'} — salario ${d.salario_mensual||'?'} USD/mes
  Historial USA: estuvo antes=${d.ha_estado_en_usa||'?'} | visa anterior=${d.ha_tenido_visa_usa||'?'} | negacion=${d.le_han_negado_visa||'?'}
  Detalle negacion: ${d.detalle_negacion||'ninguno'}
  Destino anterior: ${d.ciudad_destino_usa||'?'}, ${d.estado_destino_usa||'?'} | Proposito anterior: ${d.proposito_viaje_anterior||'?'}
  Familiares en USA: padre=${d.padre_en_usa||'?'} | madre=${d.madre_en_usa||'?'}
  Paises visitados: ${d.paises_visitados_5_anos||'ninguno'}`;
    }).join('\n\n');

    const analysisPrompt = `Eres el mejor asesor de visas B1/B2 para Ecuador del mundo, con 20 años de experiencia. Conoces exactamente como piensan los consules en Guayaquil y Quito.

Analiza estos formularios DS-160 ANTERIORES de esta familia y dame TODO lo que el asesor Roberto necesita para:
1. Llenar los DS-160 NUEVOS correctamente
2. Maximizar la probabilidad de aprobacion
3. Saber exactamente que preguntar/decirle al cliente

PERFILES EXTRAIDOS DE LOS DS-160 ANTERIORES:
${resumenPerfiles}

Responde en JSON exacto sin markdown:
{
  "probabilidad_actual": "% sin cambios",
  "probabilidad_con_estrategia": "% aplicando la estrategia",
  "paquete": "ESENCIAL $97 o PROFESIONAL $197 o VIP $397",
  "razon_paquete": "por que ese paquete",
  "consulado": "GUAYAQUIL o QUITO — razon especifica",
  "analisis_grupo": "evaluacion del grupo como unidad familiar — quien es el perfil fuerte, quien es el debil, como el consul evaluara al grupo",
  "analisis_por_viajero": "analisis de riesgo individual de cada viajero separado por | entre personas",
  "errores_formularios_anteriores": "que errores o debilidades detectas en los DS-160 anteriores que pueden haber causado problemas o que hay que corregir ahora",
  "motivo_ds160_nuevo": "texto EXACTO en espanol para el campo purpose of trip de los nuevos DS-160 — 3-4 oraciones especificas con actividades, fechas aproximadas y motivo real. NO generico.",
  "que_cambiar_vs_anterior": "que debe ser diferente en los nuevos DS-160 respecto a los anteriores — cambios concretos por campo",
  "estrategia_completa": "plan de 10-12 pasos especificos para llevar este caso al 80-100%. Incluir: como presentar el perfil de cada viajero, como manejar puntos debiles, que enfatizar",
  "guia_llenado_ds160": "instrucciones especificas para Roberto de como llenar los campos CRITICOS del DS-160 para esta familia — que poner exactamente en empleo, proposito, lazos con Ecuador, etc.",
  "documentos_exactos": "lista de 8-10 documentos con descripcion exacta de como deben estar redactados",
  "preguntas_para_cliente": "lista numerada de preguntas EXACTAS que Roberto debe hacerle al cliente para completar la informacion que falta. Estas preguntas son lo que Roberto les dira cuando los llame. Formato: 1. [PERSONA] — [pregunta exacta concreta]. Minimo 5 preguntas.",
  "fecha_viaje_ideal": "cuando viajar y por que",
  "fecha_cita_sugerida": "cuando reservar la cita y cuanto tiempo de preparacion necesitan",
  "checklist_pre_cita": "lista de 12-15 items exactos que deben tener listos antes de ir al consulado",
  "guion_entrevista": "5-7 preguntas que el consul hara SEGURO a ESTE perfil y la respuesta ideal. PREGUNTA: [...] | RESPUESTA IDEAL: [...]",
  "alerta_principal": "el riesgo mas critico y como neutralizarlo — si no hay escribir PERFIL LIMPIO",
  "proximos_pasos": "7-8 acciones que Roberto debe hacer esta semana en orden de urgencia"
}`;

    const gUrl2 = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_KEY}`;
    const resp2 = UrlFetchApp.fetch(gUrl2, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify({ contents:[{parts:[{text: analysisPrompt}]}], generationConfig:{maxOutputTokens:6000,temperature:0.3} }),
      muteHttpExceptions: true
    });
    const j2   = JSON.parse(resp2.getContentText());
    const txt2 = (j2.candidates?.[0]?.content?.parts?.[0]?.text || '{}').replace(/```json|```/g,'').trim();
    const m2   = txt2.match(/\{[\s\S]*\}/);
    const analisis = m2 ? JSON.parse(m2[0]) : {};

    // 3. Enviar email a Roberto
    enviarEmailAnalizador(ref, caseName, fecha, perfiles, analisis);

    return ok({ ok: true, ref });
  } catch(err) {
    console.error('analizarDs160Anteriores error:', err.toString());
    return ok({ error: err.toString() });
  }
}

function enviarEmailAnalizador(ref, caseName, fecha, perfiles, a) {
  try {
    // Helper: bloque de sección
    const blk = (color, border, title, content) =>
      `<div style="background:${color};border:1px solid ${border};border-radius:10px;padding:16px;margin-bottom:14px">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#1A2940;opacity:.7;margin-bottom:10px">${title}</div>
        <div style="font-size:12px;color:#1A2940;line-height:1.95">${content}</div>
      </div>`;

    // Helper: fila DS-160
    const fila = (label, valor, falta) => {
      const ok = valor && String(valor).trim() && valor !== 'null' && valor !== 'undefined';
      const bg = ok ? '' : 'background:#FFF5F5;';
      const txt = ok ? `<span style="color:#1A2940;font-weight:500">${String(valor)}</span>`
                     : `<span style="color:#DC2626;font-weight:700">⚠ FALTA${falta?' — '+falta:''}</span>`;
      return `<tr style="${bg}">
        <td style="padding:5px 8px;color:#64748B;font-size:11px;width:38%;vertical-align:top;border-bottom:1px solid #F1F5F9">${label}</td>
        <td style="padding:5px 8px;font-size:11px;vertical-align:top;border-bottom:1px solid #F1F5F9">${txt}</td>
      </tr>`;
    };

    // Ficha DS-160 completa por persona
    const fichas = perfiles.map((p, i) => {
      const d = p.datos;
      const ok  = v => v && String(v).trim() && v !== 'null';
      const llenos  = Object.values(d).filter(v => ok(v)).length;
      const totales = Object.keys(d).length;

      return `
      <div style="border:2px solid #E2E8F0;border-radius:12px;margin-bottom:20px;overflow:hidden">
        <div style="background:#1E293B;padding:12px 18px;display:flex;align-items:center;justify-content:space-between">
          <div style="color:white;font-weight:700;font-size:15px">${i+1}. ${p.nombre}</div>
          <div style="background:rgba(255,255,255,.15);border-radius:6px;padding:4px 12px;font-size:11px;color:rgba(255,255,255,.8)">${llenos} / ${totales} campos extraidos</div>
        </div>

        <table style="width:100%;border-collapse:collapse;font-size:11px">
          <tr><td colspan="2" style="padding:8px 18px 4px;background:#F8FAFC;font-size:10px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em">INFORMACION PERSONAL</td></tr>
          ${fila('Apellidos (como en pasaporte)', d.apellidos, 'apellidos exactos del pasaporte')}
          ${fila('Nombres (como en pasaporte)', d.nombres, 'nombres exactos del pasaporte')}
          ${fila('Fecha de nacimiento', d.fecha_nacimiento, 'DD/MM/AAAA')}
          ${fila('Sexo', d.sexo, 'Masculino o Femenino')}
          ${fila('Estado civil', d.estado_civil, 'Soltero/Casado/Union de hecho/Divorciado/Viudo')}
          ${fila('Cedula de identidad', d.cedula, 'numero de cedula ecuatoriana')}
          ${fila('Ciudad de nacimiento', d.ciudad_nacimiento, 'ciudad donde nacio')}
          ${fila('Provincia de nacimiento', d.provincia_nacimiento, '')}
          ${fila('Pais de nacimiento', d.pais_nacimiento || 'Ecuador', '')}
          ${fila('Nacionalidad', d.nacionalidad || 'Ecuador', '')}

          <tr><td colspan="2" style="padding:8px 18px 4px;background:#F8FAFC;font-size:10px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em">PASAPORTE</td></tr>
          ${fila('Numero de pasaporte', d.numero_pasaporte, 'numero exacto del pasaporte vigente')}
          ${fila('Tipo de pasaporte', d.tipo_pasaporte || 'Ordinario', '')}
          ${fila('Fecha de emision', d.fecha_emision_pasaporte, 'DD/MM/AAAA')}
          ${fila('Fecha de vencimiento', d.fecha_vencimiento_pasaporte, 'DD/MM/AAAA — debe tener 6+ meses')}
          ${fila('Ciudad de emision', d.ciudad_emision_pasaporte, 'ciudad donde fue emitido')}

          <tr><td colspan="2" style="padding:8px 18px 4px;background:#F8FAFC;font-size:10px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em">CONTACTO Y RESIDENCIA</td></tr>
          ${fila('Direccion de residencia', d.direccion, 'calle y numero')}
          ${fila('Ciudad de residencia', d.ciudad_residencia, 'ciudad donde vive')}
          ${fila('Provincia de residencia', d.provincia_residencia, '')}
          ${fila('Telefono principal', d.telefono, '+593 ... con codigo de pais')}
          ${fila('Email', d.email, 'correo electronico')}

          <tr><td colspan="2" style="padding:8px 18px 4px;background:#F8FAFC;font-size:10px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em">FAMILIA</td></tr>
          ${fila('Apellidos del padre', d.apellidos_padre, 'apellidos del padre')}
          ${fila('Nombres del padre', d.nombre_padre, 'nombres del padre')}
          ${fila('Fecha de nac. del padre', d.fecha_nac_padre, 'si se conoce')}
          ${fila('Padre en USA', d.padre_en_usa || 'No', '')}
          ${fila('Apellidos de la madre', d.apellidos_madre, 'apellidos de la madre')}
          ${fila('Nombres de la madre', d.nombre_madre, 'nombres de la madre')}
          ${fila('Fecha de nac. de la madre', d.fecha_nac_madre, 'si se conoce')}
          ${fila('Madre en USA', d.madre_en_usa || 'No', '')}
          ${ok(d.nombre_conyuge) || ok(d.apellidos_conyuge) ? `
          ${fila('Apellidos del conyuge', d.apellidos_conyuge, '')}
          ${fila('Nombres del conyuge', d.nombre_conyuge, '')}
          ${fila('Fecha de nac. del conyuge', d.fecha_nac_conyuge, '')}
          ${fila('Nacionalidad del conyuge', d.nacionalidad_conyuge || 'Ecuador', '')}` : ''}

          <tr><td colspan="2" style="padding:8px 18px 4px;background:#F8FAFC;font-size:10px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em">TRABAJO Y EDUCACION</td></tr>
          ${fila('Situacion laboral', d.situacion_laboral, 'Empleado/Independiente/Estudiante/Jubilado/etc')}
          ${fila('Cargo / Titulo profesional', d.cargo, 'cargo exacto')}
          ${fila('Nombre del empleador', d.empleador, 'nombre completo de la empresa')}
          ${fila('Direccion del empleador', d.direccion_empleador, 'calle y numero')}
          ${fila('Ciudad del empleador', d.ciudad_empleador, '')}
          ${fila('Telefono del empleador', d.telefono_empleador, 'con codigo de area')}
          ${fila('Fecha de inicio en empleo actual', d.fecha_inicio_empleo, 'DD/MM/AAAA')}
          ${fila('Salario mensual (USD)', d.salario_mensual, 'monto en dolares')}
          ${fila('Descripcion de funciones', d.funciones, 'que hace en su trabajo')}
          ${fila('Nivel de educacion', d.nivel_educacion, 'maximo nivel alcanzado')}
          ${d.empleos_anteriores && d.empleos_anteriores.length ?
            d.empleos_anteriores.map((e,j) => `
            ${fila(`Empleo anterior ${j+1} — Empresa`, e.empresa||e.name, '')}
            ${fila(`Empleo anterior ${j+1} — Cargo`, e.cargo||e.title, '')}
            ${fila(`Empleo anterior ${j+1} — Periodo`, `${e.desde||e.from||'?'} a ${e.hasta||e.to||'?'}`, '')}
            ${fila(`Empleo anterior ${j+1} — Razon salida`, e.razon_salida||e.reason, '')}
            `).join('') : ''}

          <tr><td colspan="2" style="padding:8px 18px 4px;background:#F8FAFC;font-size:10px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em">HISTORIAL DE VIAJES Y VISAS</td></tr>
          ${fila('Ha estado en USA antes', d.ha_estado_en_usa || 'No', '')}
          ${fila('Detalle visitas anteriores a USA', d.detalle_visitas_usa, 'cuando, cuanto tiempo, proposito')}
          ${fila('Ha tenido visa USA antes', d.ha_tenido_visa_usa || 'No', '')}
          ${fila('Numero de visa anterior', d.numero_visa_anterior, 'si aplica')}
          ${fila('Fecha emision visa anterior', d.fecha_emision_visa_anterior, 'si aplica')}
          ${fila('Tipo de visa anterior', d.tipo_visa_anterior, 'B1/B2, F1, etc')}
          ${fila('Le han negado visa o entrada a USA', d.le_han_negado_visa || 'No', '')}
          ${fila('Detalle de la negacion', d.detalle_negacion, 'CRITICO — cuando, motivo exacto, que ha cambiado')}
          ${fila('Paises visitados ultimos 5 anos', d.paises_visitados_5_anos, 'paises y fechas aproximadas')}

          <tr><td colspan="2" style="padding:8px 18px 4px;background:#F8FAFC;font-size:10px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em">VIAJE ANTERIOR (del DS-160 anterior)</td></tr>
          ${fila('Proposito declarado antes', d.proposito_viaje_anterior, '')}
          ${fila('Destino anterior', d.destino_anterior, '')}
          ${fila('Duracion anterior', d.duracion_anterior, '')}
          ${fila('Contacto USA (anterior)', d.contacto_usa_nombre, 'nombre del contacto')}
          ${fila('Relacion contacto USA', d.contacto_usa_relacion, '')}
          ${fila('Telefono contacto USA', d.contacto_usa_telefono, '')}
          ${fila('Alojamiento USA', d.alojamiento_usa, 'hotel o direccion')}
          ${fila('Ciudad destino USA', d.ciudad_destino_usa, '')}
          ${fila('Estado destino USA', d.estado_destino_usa, '')}

          <tr><td colspan="2" style="padding:8px 18px 4px;background:#F8FAFC;font-size:10px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em">OTROS</td></tr>
          ${fila('Idiomas que habla', d.idiomas || 'Espanol', '')}
          ${fila('Organizaciones a las que pertenece', d.organizaciones, 'si aplica')}
          ${fila('Servicio militar', d.tiene_servicio_militar || 'No', '')}
          ${d.redes_sociales && d.redes_sociales.length ? fila('Redes sociales', d.redes_sociales.map(r=>`${r.plataforma}: ${r.usuario}`).join(' | '), '') : ''}
        </table>
      </div>`;
    }).join('');

    const html = `<div style="font-family:Calibri,Arial,sans-serif;max-width:800px;margin:0 auto;background:#F4F6F9;padding:20px">

  <!-- CABECERA -->
  <div style="background:#060E1F;padding:22px 28px;border-radius:12px 12px 0 0;border-bottom:3px solid #F0B429">
    <div style="font-size:10px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:.1em;margin-bottom:6px">ANALISIS COMPLETO DS-160 — VISA GLOBAL</div>
    <h1 style="font-size:22px;font-weight:700;color:white;margin:0 0 6px">${caseName || ref}</h1>
    <div style="font-size:13px;color:#F0B429">${perfiles.length} viajero${perfiles.length>1?'s':''} · ${fecha} · Ref: ${ref}</div>
    <div style="margin-top:14px;display:flex;gap:12px;flex-wrap:wrap">
      <div style="background:rgba(74,222,128,.2);border:1px solid rgba(74,222,128,.4);border-radius:8px;padding:8px 16px;text-align:center">
        <div style="font-size:22px;font-weight:700;color:#4ade80">${a.probabilidad_con_estrategia||'—'}</div>
        <div style="font-size:10px;color:rgba(255,255,255,.6);text-transform:uppercase">Con estrategia</div>
      </div>
      <div style="background:rgba(248,113,113,.15);border:1px solid rgba(248,113,113,.3);border-radius:8px;padding:8px 16px;text-align:center">
        <div style="font-size:22px;font-weight:700;color:#f87171">${a.probabilidad_actual||'—'}</div>
        <div style="font-size:10px;color:rgba(255,255,255,.6);text-transform:uppercase">Sin preparacion</div>
      </div>
      <div style="background:rgba(240,180,41,.15);border:1px solid rgba(240,180,41,.3);border-radius:8px;padding:8px 16px;text-align:center">
        <div style="font-size:14px;font-weight:700;color:#F0B429">${a.paquete||'—'}</div>
        <div style="font-size:10px;color:rgba(255,255,255,.6);text-transform:uppercase">Paquete</div>
      </div>
    </div>
  </div>

  <div style="background:white;border:1px solid #E2E8F0;border-top:none;padding:24px 28px">

    ${a.alerta_principal && !a.alerta_principal.includes('LIMPIO') ? `
    <div style="background:#FEF2F2;border:2px solid #DC2626;border-radius:10px;padding:16px;margin-bottom:20px">
      <div style="font-size:11px;font-weight:700;color:#991B1B;text-transform:uppercase;letter-spacing:.07em;margin-bottom:8px">ALERTA CRITICA — ATENDER PRIMERO</div>
      <div style="font-size:13px;color:#7F1D1D;line-height:1.7">${a.alerta_principal}</div>
    </div>` : `<div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:10px;padding:10px 16px;margin-bottom:20px;font-size:12px;color:#166534;font-weight:600">Sin alertas criticas — perfil viable</div>`}

    <!-- PREGUNTAS PARA LLAMAR AL CLIENTE -->
    <div style="background:#FFF7ED;border:2px solid #F0B429;border-radius:10px;padding:18px;margin-bottom:20px">
      <div style="font-size:12px;font-weight:700;color:#92400E;text-transform:uppercase;letter-spacing:.07em;margin-bottom:12px">PREGUNTAS EXACTAS PARA LLAMAR AL CLIENTE AHORA</div>
      <div style="font-size:13px;color:#78350F;line-height:2.1">${(a.preguntas_para_cliente||'—').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong style="color:#92400E">$1</strong>')}</div>
    </div>

    <!-- FICHAS DS-160 COMPLETAS -->
    <div style="font-size:12px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em;margin-bottom:12px">FICHAS COMPLETAS DS-160 POR VIAJERO — LISTOS PARA LLENAR</div>
    <div style="font-size:11px;color:#64748B;margin-bottom:14px">Los campos en rojo (⚠ FALTA) son los que debes preguntar al cliente. Los demas ya puedes copiarlos directamente al DS-160.</div>
    ${fichas}

    <!-- TEXTO EXACTO PURPOSE OF TRIP -->
    <div style="background:#F0FDF4;border:2px solid #22c55e;border-radius:10px;padding:18px;margin-bottom:14px">
      <div style="font-size:11px;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:.07em;margin-bottom:10px">TEXTO EXACTO — COPIAR EN "PURPOSE OF TRIP" DEL DS-160 NUEVO</div>
      <div style="font-size:14px;color:#166534;font-style:italic;line-height:1.8;border-left:3px solid #22c55e;padding-left:12px">"${a.motivo_ds160_nuevo||'—'}"</div>
      <div style="font-size:11px;color:#166534;margin-top:8px">Usar este texto para todos los adultos del grupo.</div>
    </div>

    <!-- ERRORES FORMULARIOS ANTERIORES -->
    ${blk('#FEF2F2','#FCA5A5','ERRORES Y RIESGOS EN LOS DS-160 ANTERIORES — QUE CORREGIR',(a.errores_formularios_anteriores||'—').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong>$1</strong>'))}

    <!-- QUE CAMBIAR -->
    ${blk('#FFF7ED','#FCD34D','QUE CAMBIAR EN LOS NUEVOS DS-160 VS LOS ANTERIORES',(a.que_cambiar_vs_anterior||'—').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong style="color:#92400E">$1</strong>'))}

    <!-- GUIA DE LLENADO -->
    ${blk('#F8F9FF','#C7D2FE','GUIA DE LLENADO — INSTRUCCIONES PARA CAMPOS CRITICOS DEL DS-160',(a.guia_llenado_ds160||'—').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong style="color:#3730A3">$1</strong>'))}

    <!-- ESTRATEGIA -->
    ${blk('#EFF6FF','#93C5FD','ESTRATEGIA COMPLETA PARA MAXIMIZAR APROBACION',(a.estrategia_completa||'—').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong style="color:#1E40AF">$1</strong>'))}

    <!-- ANALISIS POR PERSONA -->
    ${blk('#F8FAFC','#E2E8F0','ANALISIS DE RIESGO POR VIAJERO',(a.analisis_por_viajero||'—').replace(/\|/g,'<br><br>').replace(/\n/g,'<br>'))}

    <!-- LOGISTICA -->
    <div style="background:#EFF6FF;border:1px solid #93C5FD;border-radius:10px;padding:16px;margin-bottom:14px">
      <div style="font-size:11px;font-weight:700;color:#1E40AF;text-transform:uppercase;letter-spacing:.07em;margin-bottom:10px">LOGISTICA — CONSULADO Y FECHAS</div>
      <table style="width:100%;font-size:12px;border-collapse:collapse">
        <tr><td style="padding:6px 0;color:#64748B;width:38%;vertical-align:top">Consulado recomendado</td><td style="font-weight:600;color:#1A2940">${a.consulado||'—'}</td></tr>
        <tr><td style="padding:6px 0;color:#64748B;vertical-align:top">Fecha de viaje ideal</td><td style="color:#1A2940">${a.fecha_viaje_ideal||'—'}</td></tr>
        <tr><td style="padding:6px 0;color:#64748B;vertical-align:top">Cuando reservar la cita</td><td style="color:#1A2940">${a.fecha_cita_sugerida||'—'}</td></tr>
      </table>
    </div>

    <!-- DOCUMENTOS -->
    <div style="background:#FEF9EE;border:1px solid #FCD34D;border-radius:10px;padding:16px;margin-bottom:14px">
      <div style="font-size:11px;font-weight:700;color:#92400E;text-transform:uppercase;letter-spacing:.07em;margin-bottom:10px">DOCUMENTOS EXACTOS A PREPARAR</div>
      <div style="font-size:12px;color:#1A2940;line-height:1.9">${(a.documentos_exactos||'—').split(';').map((s,i)=>`<div style="padding:6px 0;border-bottom:1px solid #FEF3C7"><strong style="color:#92400E">${i+1}.</strong> ${s.trim()}</div>`).join('')}</div>
    </div>

    <!-- CHECKLIST PRE-CITA -->
    ${blk('#F8F9FF','#C7D2FE','CHECKLIST — LO QUE DEBEN LLEVAR AL CONSULADO',(a.checklist_pre_cita||'—').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong>$1</strong>'))}

    <!-- GUION ENTREVISTA -->
    <div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:10px;padding:16px;margin-bottom:14px">
      <div style="font-size:11px;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:.07em;margin-bottom:10px">GUION DE ENTREVISTA</div>
      <div style="font-size:12px;color:#1A2940;line-height:2">${(a.guion_entrevista||'—').replace(/PREGUNTA:/g,'<br><strong style="color:#166534">PREGUNTA:</strong>').replace(/RESPUESTA IDEAL:/g,'<strong style="color:#1E40AF">RESPUESTA IDEAL:</strong>')}</div>
    </div>

    <!-- PROXIMOS PASOS -->
    ${blk('#EFF6FF','#93C5FD','PROXIMOS PASOS ESTA SEMANA — EN ORDEN DE URGENCIA',(a.proximos_pasos||'—').replace(/\n/g,'<br>').replace(/(\d+\.\s)/g,'<strong style="color:#1E40AF">$1</strong>'))}

  </div>
  <div style="text-align:center;padding:14px;font-size:11px;color:#94A3B8">Asesoria Visa Global · ${fecha} · Ref: ${ref}</div>
</div>`;

    MailApp.sendEmail({ to: EMAIL_ROBERTO, subject: `ANALISIS DS-160 — ${caseName||ref} — ${perfiles.length} viajero${perfiles.length>1?'s':''} [${ref}]`, htmlBody: html });
  } catch(e) {
    console.error('enviarEmailAnalizador error:', e.toString());
  }
}

function extractarDS160(payload) {
  try {
    const extractId   = payload.extractId;
    const travelerIdx = payload.travelerIdx || 0;
    const pdfText     = payload.pdfText || '';

    const prompt = `Eres un experto en formularios DS-160 de visa USA. Analiza este texto extraido de un DS-160 completado anteriormente y extrae TODOS los datos que puedas identificar. Devuelve SOLO un JSON valido sin markdown ni texto extra. Usa null SOLO si el dato definitivamente no aparece en el texto. No inventes datos.

JSON a completar (null = no encontrado):
{
  "surnames":"apellidos MAYUSCULAS como en pasaporte",
  "givenNames":"nombres MAYUSCULAS como en pasaporte",
  "dob":"YYYY-MM-DD fecha de nacimiento",
  "sex":"Masculino o Femenino",
  "maritalStatus":"Soltero/a | Casado/a | Union de hecho | Divorciado/a | Viudo/a",
  "hasOtherNames":"si o no",
  "otherNames":"otros nombres usados si aplica",
  "cityOfBirth":"ciudad de nacimiento",
  "stateOfBirth":"provincia de nacimiento",
  "countryOfBirth":"pais de nacimiento",
  "nationality":"nacionalidad principal",
  "hasSecondNationality":"si o no",
  "secondNationality":"segunda nacionalidad si aplica",
  "nationalId":"cedula de identidad 10 digitos",
  "passportNumber":"numero de pasaporte",
  "passportType":"Ordinario",
  "passportCountry":"pais que emitio el pasaporte",
  "passportCity":"ciudad donde fue emitido",
  "passportIssueDate":"YYYY-MM-DD fecha de emision",
  "passportExpiry":"YYYY-MM-DD fecha de vencimiento",
  "hadLostPassport":"si o no",
  "lostPassportNumber":"numero del pasaporte perdido o robado si aplica",
  "lostPassportCountry":"pais que emitio el pasaporte perdido si aplica",
  "homeStreet":"direccion de residencia calle y numero",
  "homeCity":"ciudad de residencia",
  "homeProvince":"provincia de residencia",
  "homeCountry":"pais de residencia",
  "primaryPhone":"telefono principal con codigo de pais",
  "secondaryPhone":"telefono secundario si hay",
  "email":"correo electronico",
  "fatherSurname":"apellidos del padre",
  "fatherGivenName":"nombres del padre",
  "fatherDob":"YYYY-MM-DD fecha nacimiento padre",
  "fatherInUS":"No | Si — Ciudadano americano | Si — Residente permanente (green card) | Si — Visa temporal | No lo se",
  "motherSurname":"apellidos de la madre",
  "motherGivenName":"nombres de la madre",
  "motherDob":"YYYY-MM-DD fecha nacimiento madre",
  "motherInUS":"No | Si — Ciudadana americana | Si — Residente permanente (green card) | Si — Visa temporal | No lo se",
  "hasRelativesInUS":"si o no",
  "relativesInUSDetails":"descripcion de familiares en USA aparte de padres",
  "isMarried":"si o no",
  "spouseSurname":"apellidos del conyuge si aplica",
  "spouseGivenName":"nombres del conyuge si aplica",
  "spouseDob":"YYYY-MM-DD fecha nacimiento conyuge",
  "spouseNationality":"nacionalidad del conyuge",
  "spouseCountryBirth":"pais de nacimiento del conyuge",
  "employmentStatus":"Empleado (relacion de dependencia) | Independiente / Empresa propia | Estudiante | Jubilado/a | Ama/Amo de casa | Sin empleo actualmente",
  "currentOccupation":"cargo o titulo profesional actual",
  "currentEmployerName":"nombre de la empresa actual",
  "currentEmployerStreet":"direccion de la empresa actual",
  "currentEmployerCity":"ciudad de la empresa actual",
  "currentEmployerPhone":"telefono de la empresa",
  "currentStartDate":"YYYY-MM-DD fecha de inicio en el empleo actual",
  "monthlySalary":"salario mensual en USD",
  "jobDuties":"descripcion de las funciones laborales",
  "hasPreviousEmployers":"si o no tuvo empleos anteriores en ultimos 5 anos",
  "previousEmployers":[{"name":"empresa","title":"cargo","city":"ciudad","from":"YYYY-MM-DD","to":"YYYY-MM-DD","reason":"razon de salida"}],
  "schoolName":"nombre de institucion educativa si es estudiante",
  "schoolCity":"ciudad de la institucion",
  "schoolCourse":"carrera o programa de estudios",
  "educationSummary":"resumen de educacion desde los 16 anos colegios y universidades con fechas",
  "hasBeenInUS":"si o no ha estado en USA antes",
  "usVisitDetails":"cuando fue y cuanto tiempo se quedo en USA",
  "hasHadUSVisa":"si o no ha tenido visa americana",
  "previousVisaNumber":"numero de la visa americana anterior",
  "previousVisaIssueDate":"YYYY-MM-DD fecha de emision de la visa anterior",
  "previousVisaType":"B1/B2 Turismo | F-1 Estudiante | H-1B Trabajo | Otro",
  "hasBeenRefused":"si o no le han negado visa o entrada a USA",
  "refusalDetails":"cuando y por que razon fue rechazado",
  "countriesVisited5Years":"paises visitados en los ultimos 5 anos excepto Ecuador y USA con fechas aproximadas",
  "languages":"idiomas que habla",
  "organizations":"organizaciones o grupos a los que pertenece",
  "hasMilitaryService":"si o no ha prestado servicio militar"
}

TEXTO DEL DS-160:
${pdfText.substring(0, 7000)}`;

    const gUrl3 = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_KEY}`;
    const resp = UrlFetchApp.fetch(gUrl3, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { maxOutputTokens: 2500, temperature: 0.2 }
      }),
      muteHttpExceptions: true,
    });

    const json  = JSON.parse(resp.getContentText());
    const text  = (json.candidates?.[0]?.content?.parts?.[0]?.text || '{}').trim();
    const clean = text.replace(/```json|```/g, '').trim();
    const match = clean.match(/\{[\s\S]*\}/);
    const extracted = match ? JSON.parse(match[0]) : {};

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

// ── Reconstruir estado del formulario desde Intake DS-160 Detalle ─
// Usado cuando el draft fue borrado al enviar pero los datos están en Sheets
function reconstructFromDetalle(e) {
  try {
    const ref = (e.parameter.ref || '').trim();
    if (!ref) return ContentService.createTextOutput(JSON.stringify({ok:false,msg:'no_ref'})).setMimeType(ContentService.MimeType.JSON);

    const ss = SpreadsheetApp.openById(SS_ID);

    // ── Intento 1: Borrador guardado (formulario en progreso) ────────
    const shBorr = ss.getSheetByName('Borradores');
    if (shBorr) {
      const bRows = shBorr.getDataRange().getValues();
      // Columnas: Ref | Datos JSON | Actualizado | Paso
      for (let i = 1; i < bRows.length; i++) {
        if (String(bRows[i][0]).trim() !== ref) continue;
        const jsonStr = String(bRows[i][1]).trim();
        if (!jsonStr || jsonStr === '{}') continue;
        try {
          const parsed = JSON.parse(jsonStr);
          const savedS = parsed.S || parsed;
          // Solo usar si tiene viajeros con nombre
          if (savedS.travelers && savedS.travelers.length > 0 &&
              savedS.travelers.some(t => t.name && t.name.trim())) {
            const state = Object.assign({}, savedS, {
              phase: 'review',
              submitted: savedS.submitted || false
            });
            return ContentService
              .createTextOutput(JSON.stringify({ ok: true, S: state, source: 'borrador' }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        } catch(parseErr) {}
      }
    }

    // ── Intento 2: Formulario enviado (Intake DS-160 Detalle) ────────
    const shDet = ss.getSheetByName('Intake DS-160 Detalle');
    if (!shDet) return ContentService.createTextOutput(JSON.stringify({ok:false,msg:'sin_datos'})).setMimeType(ContentService.MimeType.JSON);

    const rows = shDet.getDataRange().getValues();
    const refRows = rows.filter((r, i) => i > 0 && String(r[0]).trim() === ref);
    if (!refRows.length) return ContentService.createTextOutput(JSON.stringify({ok:false,msg:'not_found'})).setMimeType(ContentService.MimeType.JSON);

    const shared = {};
    const travelerMap = {};
    const travelerOrder = [];

    refRows.forEach(r => {
      const viajero = String(r[2]).trim();
      const rol     = String(r[3]).trim();
      const campo   = String(r[4]).trim();
      const valor   = String(r[5]).trim();
      if (viajero === 'COMPARTIDO') {
        shared[campo] = valor;
      } else {
        if (!travelerMap[viajero]) {
          travelerMap[viajero] = { datos: {}, rol: rol };
          travelerOrder.push(viajero);
        }
        if (campo !== 'BANDERAS_SEGURIDAD') travelerMap[viajero].datos[campo] = valor;
      }
    });

    const travelers = [], perTraveler = [], securityFlags = {};
    travelerOrder.forEach((nombre, i) => {
      const info = travelerMap[nombre];
      travelers.push({ name: nombre, role: info.rol, dob: info.datos.dob || '' });
      perTraveler.push(info.datos);
      const flagRow = refRows.find(r => String(r[2]).trim() === nombre && String(r[4]).trim() === 'BANDERAS_SEGURIDAD');
      if (flagRow) {
        const flags = String(flagRow[5]).split(',').map(s => s.trim()).filter(Boolean);
        if (flags.length) securityFlags[i] = flags;
      }
    });

    const state = {
      phase: 'review', submitted: true,
      sharedStep: 0, travelerIdx: 0, travelerStep: 0,
      travelers, shared, perTraveler, securityFlags
    };

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, S: state, source: 'detalle' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    console.error('reconstructFromDetalle error:', err.toString());
    return ContentService.createTextOutput(JSON.stringify({ok:false,msg:err.toString()})).setMimeType(ContentService.MimeType.JSON);
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

// ═══════════════════════════════════════════════════════════════════
// AUTOMATIZACIÓN DS-160 — Generación automática por caso
// ═══════════════════════════════════════════════════════════════════

// Mapea datos de intake → campos exactos DS-160
function generarDS160PreFill(personas, shared, analisis) {
  const guia = {
    ref: shared.ref || '',
    fecha_generado: new Date().toISOString(),
    datos_coordinados: {
      fecha_llegada: shared.intendedArrival || 'PENDIENTE',
      duracion: (shared.lengthOfStayDays || 14) + ' DAYS',
      hotel_nombre: shared.usStayName || 'PENDIENTE',
      hotel_address: (shared.usStayAddress || '') + ', ' + (shared.usStayCity || '') + ', ' + (shared.usStayState || '') + ' ' + (shared.usStayZip || ''),
      contacto_usa_org: shared.usStayName || 'PENDIENTE',
      rechazo_previo: shared.hayRechazo === 'si' ? 'YES — ' + (shared.rechazoDetalles || 'Indicar fecha y consulado') : 'NO',
      consulado_recomendado: (analisis && analisis.consulado) || 'PENDIENTE'
    },
    viajeros: personas.map((p, idx) => {
      const d = p.datos || {};
      const nombres = (p.nombre || '').split(',');
      const apellido = nombres[0] ? nombres[0].trim().toUpperCase() : '';
      const nombre   = nombres[1] ? nombres[1].trim().toUpperCase() : p.nombre.toUpperCase();
      const esAdulto = p.rol && !p.rol.includes('Menor');
      const pagador  = idx === 0 ? 'SELF' : ('OTHER PERSON — ' + (personas[0].nombre || '').toUpperCase() + ' — ' + (personas[0].datos.primaryPhone || ''));

      return {
        viajero_num: idx + 1,
        rol: p.rol || 'Adulto',
        campos: {
          // Información personal
          surname: apellido || 'REVISAR',
          given_names: nombre || 'REVISAR',
          full_name_native: 'NO APLICA',
          other_names: 'NO',
          sex: d.sex || (p.rol && p.rol.includes('Menor') ? 'MASCULINO' : 'PENDIENTE'),
          marital_status: d.maritalStatus || 'PENDIENTE',
          date_of_birth: d.dateOfBirth || 'PENDIENTE',
          place_of_birth: (d.birthCity || 'PENDIENTE') + ', ECUADOR',
          nationality: d.nationality || 'ECUADOR',
          national_id: d.nationalId || d.passportNumber || 'PENDIENTE',
          us_social_security: 'NO APLICA',
          us_tax_id: 'NO APLICA',
          // Dirección
          home_address: (d.homeAddress || d.homeCity || 'PENDIENTE') + ', ECUADOR',
          same_mailing: 'YES',
          primary_phone: d.primaryPhone || 'PENDIENTE',
          work_phone: d.workPhone || (esAdulto ? 'PENDIENTE' : 'NO APLICA'),
          email: d.email || 'PENDIENTE',
          social_media_1_platform: d.socialMedia1 || (esAdulto ? 'Facebook' : 'Ninguno'),
          social_media_1_handle: d.socialHandle1 || 'PENDIENTE — verificar nombre real del perfil',
          // Pasaporte
          passport_type: 'REGULAR',
          passport_number: d.passportNumber || 'PENDIENTE',
          book_number: d.bookNumber || 'PENDIENTE',
          passport_issued_by: 'ECUADOR',
          passport_city: d.passportCity || 'PENDIENTE',
          passport_state: d.passportState || 'PENDIENTE',
          passport_issue_date: d.passportIssueDate || 'PENDIENTE',
          passport_expiry: d.passportExpiry || 'PENDIENTE',
          lost_stolen: 'NO',
          // Viaje
          purpose: 'B1/B2 — BUSINESS OR TOURISM (TEMPORARY VISITOR)',
          specific_plans: 'YES',
          intended_arrival: shared.intendedArrival || 'PENDIENTE',
          length_of_stay: (shared.lengthOfStayDays || 14) + ' DAYS',
          us_address: (shared.usStayName || '') + ', ' + (shared.usStayAddress || '') + ', ' + (shared.usStayCity || '') + ', ' + (shared.usStayState || ''),
          who_pays: idx === 0 ? 'SELF' : pagador,
          // Familia
          father_surnames: d.fatherSurname || 'PENDIENTE',
          father_given: d.fatherGiven || 'PENDIENTE',
          father_dob: d.fatherDOB || 'DO NOT KNOW',
          father_in_us: d.fatherInUS || 'NO',
          mother_surnames: d.motherSurname || 'PENDIENTE',
          mother_given: d.motherGiven || 'PENDIENTE',
          mother_dob: d.motherDOB || 'DO NOT KNOW',
          mother_in_us: d.motherInUS || 'NO',
          relatives_us: d.hasRelativesInUS === 'si' ? 'YES — ' + (d.relativesInUSDetails || 'ESPECIFICAR') : 'NO',
          // Empleo
          occupation: d.employmentStatus === 'empleado' ? 'EMPLOYED' : (d.employmentStatus === 'gobierno' ? 'GOVERNMENT' : (d.employmentStatus === 'independiente' ? 'SELF EMPLOYED' : (esAdulto ? 'PENDIENTE' : 'STUDENT'))),
          employer_name: d.currentEmployerName || (esAdulto ? 'PENDIENTE' : d.schoolName || 'PENDIENTE'),
          employer_address: d.currentEmployerAddress || 'PENDIENTE',
          monthly_salary: d.monthlySalary || (esAdulto ? 'PENDIENTE' : 'DOES NOT APPLY'),
          duties: d.jobDescription || (esAdulto ? 'PENDIENTE — describir cargo en inglés mayúsculas, 2-3 oraciones' : 'ESTUDIANTE — ' + (d.schoolName || 'PENDIENTE')),
          previously_employed: d.previousEmployer ? 'YES — ' + d.previousEmployer : 'NO',
          education_secondary_above: d.educationLevel ? 'YES — ' + d.educationLevel : 'PENDIENTE',
          // Historial
          been_in_us: d.hasBeenInUS === 'si' ? 'YES' : 'NO',
          us_visa_issued: d.hasHadUSVisa === 'si' ? 'YES' : 'NO',
          visa_refused: (d.hasBeenRefused === 'si' || shared.hayRechazo === 'si') ? 'YES — ' + (d.refusalDetails || shared.rechazoDetalles || 'INDICAR FECHA Y CONSULADO') : 'NO',
          immigrant_petition: 'NO',
          // Seguridad
          communicable_disease: 'NO',
          mental_disorder: 'NO',
          drug_abuser: 'NO',
          arrested: 'NO',
          controlled_substances: 'NO',
          prostitution: 'NO',
          money_laundering: 'NO',
          human_trafficking: 'NO',
          terrorist: 'NO',
          genocide: 'NO',
          torturer: 'NO',
          child_soldier: 'NO',
          religious_freedom: 'NO',
          population_controls: 'NO',
          transplant: 'NO'
        }
      };
    })
  };
  return guia;
}

// Guarda la guía DS-160 en la hoja "Guias DS-160" de Sheets
function guardarGuiaDS160(refId, guiaJSON, analisis) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = getOrCreateSheet(ss, 'Guias DS-160',
      ['Ref ID','Fecha','Num Viajeros','DS160_JSON','Probabilidad','Paquete','Consulado','Estado']);

    const rows = sh.getDataRange().getValues();
    const refCol = 0;
    const existingIdx = rows.findIndex((r, i) => i > 0 && String(r[refCol]) === refId);

    const rowData = [
      refId,
      Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm'),
      guiaJSON.viajeros ? guiaJSON.viajeros.length : 0,
      JSON.stringify(guiaJSON),
      (analisis && analisis.probabilidad) || '—',
      (analisis && analisis.paquete) || '—',
      (analisis && analisis.consulado) || '—',
      'Generado'
    ];

    if (existingIdx > 0) {
      sh.getRange(existingIdx + 1, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sh.appendRow(rowData);
    }
    formatearCabecera(sh, 1);
    return true;
  } catch(e) {
    console.error('guardarGuiaDS160 error:', e.toString());
    return false;
  }
}

// Genera HTML de la guía DS-160 para un caso
function generarHTMLGuiaDS160(guia, analisis) {
  const colores = { ok: '#F0FDF4', pendiente: '#EFF6FF' };

  function fila(campo, valor, tipo) {
    const bg = tipo === 'pendiente' ? colores.pendiente : colores.ok;
    const prefijo = tipo === 'pendiente' ? '? ' : '✓ ';
    const color = tipo === 'pendiente' ? '#1E40AF' : '#166534';
    return `<tr>
      <td style="padding:6px 10px;font-size:11px;font-weight:600;color:#555;border-bottom:1px solid #eee;width:35%">${campo}</td>
      <td style="padding:6px 10px;font-size:11px;font-weight:700;color:${color};background:${bg};border-bottom:1px solid #eee">${prefijo}${valor}</td>
    </tr>`;
  }

  function tabla(campos) {
    return '<table style="width:100%;border-collapse:collapse;margin-bottom:14px">' +
      Object.entries(campos).map(([k, v]) => {
        const isPend = String(v).includes('PENDIENTE') || String(v).includes('DO NOT KNOW') || String(v).includes('REVISAR');
        return fila(k.replace(/_/g,' ').toUpperCase(), v, isPend ? 'pendiente' : 'ok');
      }).join('') + '</table>';
  }

  const viajerosSections = (guia.viajeros || []).map(v => {
    const c = v.campos;
    return `
    <div style="margin-bottom:24px;border:1px solid #e8e0d4;border-radius:8px;overflow:hidden">
      <div style="background:#060E1F;padding:12px 20px;color:white">
        <span style="font-size:10px;font-weight:700;text-transform:uppercase;color:#F0B429">Viajero ${v.viajero_num} — ${v.rol}</span>
      </div>
      <div style="padding:16px 20px">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#999;margin:10px 0 5px">INFORMACIÓN PERSONAL</div>
        ${tabla({surname: c.surname, 'given names': c.given_names, 'full name native': c.full_name_native, sex: c.sex, 'marital status': c.marital_status, 'date of birth': c.date_of_birth, 'place of birth': c.place_of_birth, nationality: c.nationality, 'national ID': c.national_id})}
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#999;margin:10px 0 5px">DIRECCIÓN Y CONTACTO</div>
        ${tabla({'home address': c.home_address, 'primary phone': c.primary_phone, 'work phone': c.work_phone, email: c.email, 'social media 1 — platform': c.social_media_1_platform, 'social media 1 — handle': c.social_media_1_handle})}
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#999;margin:10px 0 5px">PASAPORTE</div>
        ${tabla({type: c.passport_type, number: c.passport_number, 'book number': c.book_number, 'issued by': c.passport_issued_by, city: c.passport_city, 'issue date': c.passport_issue_date, expiry: c.passport_expiry, 'lost or stolen': c.lost_stolen})}
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#999;margin:10px 0 5px">INFORMACIÓN DE VIAJE</div>
        ${tabla({purpose: c.purpose, 'intended arrival': c.intended_arrival, 'length of stay': c.length_of_stay, 'us address': c.us_address, 'who pays': c.who_pays})}
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#999;margin:10px 0 5px">FAMILIA</div>
        ${tabla({'father surnames': c.father_surnames, 'father given names': c.father_given, 'father dob': c.father_dob, 'father in us': c.father_in_us, 'mother surnames': c.mother_surnames, 'mother given names': c.mother_given, 'mother dob': c.mother_dob, 'relatives in us': c.relatives_us})}
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#999;margin:10px 0 5px">EMPLEO / EDUCACIÓN</div>
        ${tabla({'primary occupation': c.occupation, 'employer name': c.employer_name, 'employer address': c.employer_address, 'monthly salary': c.monthly_salary, duties: c.duties, 'previously employed': c.previously_employed, education: c.education_secondary_above})}
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#999;margin:10px 0 5px">HISTORIAL</div>
        ${tabla({'been in us': c.been_in_us, 'us visa issued': c.us_visa_issued, 'visa refused': c.visa_refused, 'immigrant petition': c.immigrant_petition})}
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#999;margin:10px 0 5px">PREGUNTAS DE SEGURIDAD — TODAS:</div>
        <p style="font-size:11px;color:#166534;font-weight:700;margin:0">✓ NO — en todas las preguntas de seguridad</p>
      </div>
    </div>`;
  }).join('');

  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>Guia DS-160 — ${guia.ref}</title>
<style>body{font-family:Segoe UI,Arial,sans-serif;background:#F0EDE8;margin:0;padding:20px;font-size:12px}
.doc{max-width:860px;margin:0 auto;background:white;border-radius:8px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.1)}
</style></head><body><div class="doc">
<div style="background:#060E1F;padding:24px 32px;color:white">
  <div style="color:#F0B429;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;margin-bottom:6px">Guia DS-160 — Uso Interno — Asesoría Visa Global</div>
  <h1 style="margin:0 0 4px;font-size:18px">Campos exactos para los nuevos DS-160</h1>
  <p style="margin:0;color:rgba(255,255,255,.5);font-size:11px">Expediente ${guia.ref} · Generado ${guia.fecha_generado} · ${(guia.viajeros||[]).length} viajeros</p>
</div>
<div style="padding:24px 32px">
  <div style="display:flex;gap:16px;margin-bottom:20px;flex-wrap:wrap">
    <div style="padding:6px 14px;border-radius:20px;font-size:11px;font-weight:600;background:#F0FDF4;border:1px solid #86EFAC;color:#166534">✓ Verde — copiar exactamente</div>
    <div style="padding:6px 14px;border-radius:20px;font-size:11px;font-weight:600;background:#EFF6FF;border:1px solid #93C5FD;color:#1E40AF">? Azul — confirmar con cliente primero</div>
  </div>
  <div style="background:#F8F5F0;border-left:3px solid #F0B429;padding:10px 14px;margin-bottom:20px;font-size:11px">
    <strong>Datos coordinados — idénticos en los 4 formularios:</strong><br>
    Llegada: ${guia.datos_coordinados.fecha_llegada} | Duración: ${guia.datos_coordinados.duracion} |
    Hotel: ${guia.datos_coordinados.hotel_nombre} | Rechazo previo: ${guia.datos_coordinados.rechazo_previo}
  </div>
  ${viajerosSections}
</div>
<div style="background:#060E1F;padding:16px 32px;color:rgba(255,255,255,.4);font-size:11px">Asesoría Visa Global · Roberto Acosta · nanotiendaec@gmail.com</div>
</div></body></html>`;
}

// Endpoint público para ver la guía DS-160 de un caso — llamar con ?action=ds160guide&ref=VG-XXXXX&pin=visa2026
function servirGuiaDS160(ref, pin) {
  if (pin !== (ADMIN_PIN || 'visa2026')) {
    return HtmlService.createHtmlOutput('<p>Acceso denegado</p>');
  }
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName('Guias DS-160');
    if (!sh) return HtmlService.createHtmlOutput('<p>No hay guias generadas aún</p>');
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === ref) {
        const guia = JSON.parse(rows[i][3]);
        const html = generarHTMLGuiaDS160(guia, { probabilidad: rows[i][4], paquete: rows[i][5], consulado: rows[i][6] });
        return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    }
    return HtmlService.createHtmlOutput('<p>Caso ' + ref + ' no encontrado</p>');
  } catch(e) {
    return HtmlService.createHtmlOutput('<p>Error: ' + e.toString() + '</p>');
  }
}

// ═══════════════════════════════════════════════════════════════════
// AUTO-CARTA: Genera la carta del cliente automáticamente
// Acceso: ?action=carta&ref=VG-XXXXX&pin=visa2026
// ═══════════════════════════════════════════════════════════════════

// ── Analiza DS-160 anteriores enviados desde admin (textos ya extraidos por bot) ──
function analizarDs160AnterioresAuto(payload) {
  try {
    const ref         = payload.ref || '';
    const textosPdfs  = payload.textos_pdfs || '';
    const numArchivos = payload.num_archivos || 1;
    const fecha       = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');

    if (!textosPdfs) return ok({ error: 'Sin texto de PDFs' });

    // Leer datos actuales del caso desde Sheets
    const ss  = SpreadsheetApp.openById(SS_ID);
    const crm = ss.getSheetByName('CASOS CRM');
    let datosActuales = 'No disponibles';
    if (crm) {
      const rows = crm.getDataRange().getValues();
      const idx  = rows.findIndex((r, i) => i > 0 && String(r[0]) === ref);
      if (idx > 0) datosActuales = JSON.stringify(rows[idx]);
    }

    // Analisis con Gemini
    const prompt = `Eres el mejor asesor de visas B1/B2 USA para Ecuador. Analiza estos DS-160 anteriores que fueron rechazados y genera un analisis completo de correcciones.

TEXTOS DE DS-160 ANTERIORES (${numArchivos} formularios):
${textosPdfs.substring(0, 12000)}

DATOS ACTUALES DEL CASO EN CRM:
${datosActuales}

Genera en HTML limpio (sin backticks, sin markdown) usando estas secciones exactas:

<h3>Lo que declararon en los DS-160 anteriores</h3>
[tabla: Viajero | Ocupacion declarada | Fecha llegada | Duracion | Quien paga]

<h3>Inconsistencias detectadas entre formularios</h3>
[lista de inconsistencias criticas entre los formularios del grupo]

<h3>Errores que causaron el rechazo</h3>
[lista numerada ordenada por gravedad, con explicacion de por que cada error es problema]

<h3>Correcciones obligatorias en los nuevos DS-160</h3>
[tabla: Campo | Valor anterior (incorrecto) | Valor correcto a usar]

<h3>Datos a confirmar con el cliente antes de llenar</h3>
[lista de preguntas especificas que el asesor debe hacer]`;

    let analisisHtml = '';
    try {
      const gUrl  = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_KEY}`;
      const gResp = UrlFetchApp.fetch(gUrl, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { maxOutputTokens: 3000, temperature: 0.2 }
        }),
        muteHttpExceptions: true
      });
      const gData = JSON.parse(gResp.getContentText());
      analisisHtml = (gData.candidates || [])[0]?.content?.parts?.[0]?.text || '';
      analisisHtml = analisisHtml.replace(/```html?\s*/g, '').replace(/```\s*/g, '').trim();
    } catch(e) {
      analisisHtml = '<p>Error en analisis IA: ' + e.toString() + '</p>';
    }

    // Guardar analisis en hoja "DS160 Anteriores"
    const shAnt = getOrCreateSheet(ss, 'DS160 Anteriores',
      ['Ref ID', 'Fecha', 'Num PDFs', 'Analisis HTML']);
    const rows  = shAnt.getDataRange().getValues();
    const idx   = rows.findIndex((r, i) => i > 0 && String(r[0]) === ref);
    const row   = [ref, fecha, numArchivos, analisisHtml];
    if (idx > 0) { shAnt.getRange(idx + 1, 1, 1, row.length).setValues([row]); }
    else         { shAnt.appendRow(row); }
    formatearCabecera(shAnt, 1);

    // Notificar a Roberto por email
    try {
      MailApp.sendEmail({
        to: EMAIL_ROBERTO,
        subject: `DS-160 anteriores analizados — ${ref}`,
        htmlBody: `<div style="font-family:sans-serif;max-width:700px;margin:0 auto">
          <div style="background:#060E1F;color:white;padding:20px 24px;border-radius:8px 8px 0 0">
            <div style="font-size:11px;color:rgba(255,255,255,.5);margin-bottom:4px">Asesoria Visa Global — Analisis Automatico</div>
            <h2 style="margin:0;font-size:18px">DS-160 anteriores analizados — ${ref}</h2>
            <div style="font-size:12px;color:rgba(255,255,255,.5);margin-top:4px">${numArchivos} PDF(s) procesados el ${fecha}</div>
          </div>
          <div style="background:white;border:1px solid #e2e8f0;padding:24px;border-radius:0 0 8px 8px">
            ${analisisHtml}
          </div>
        </div>`
      });
    } catch(e) { console.error('Email DS160 error:', e.toString()); }

    return ok({ ok: true, ref, analisis: 'generado', fecha });
  } catch(e) {
    return ok({ error: e.toString() });
  }
}

function generarCartaHTML(refId, personas, shared, analisis) {
  const fecha          = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy');
  const nombrePrincipal = personas[0] ? (personas[0].nombre || 'Cliente').split(' ')[0] : 'Cliente';
  const numViaj        = personas.length;
  const prob           = analisis.probabilidad || '—';
  const probNum        = parseFloat(prob) || 0;
  const colorProb      = probNum >= 70 ? '#16a34a' : probNum >= 50 ? '#D97706' : '#DC2626';
  const consulado      = (analisis.consulado || '—').split('—')[0].trim();
  const paquete        = analisis.paquete || '—';

  // Viajeros chips
  const viajList = personas.map(p =>
    `<span style="display:inline-block;background:#F0FDF4;border:1px solid #86EFAC;color:#166534;font-size:11px;font-weight:600;padding:3px 10px;border-radius:20px;margin:2px">${p.nombre} (${p.rol})</span>`
  ).join(' ');

  // Documentos — lista numerada
  const docs = (analisis.documentos_exactos || '').split(';').map(d => d.trim()).filter(Boolean);
  const docsHTML = docs.map((d, i) =>
    `<div style="display:flex;align-items:flex-start;gap:8px;padding:7px 0;border-bottom:1px solid #f0ece4">
       <div style="min-width:21px;height:21px;border-radius:50%;background:#F0B429;color:#060E1F;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0">${i+1}</div>
       <div style="font-size:12px;color:#444;line-height:1.5">${d}</div>
     </div>`
  ).join('');

  // Estrategia — pasos numerados
  const pasosEstrategia = (analisis.estrategia || '').split(/\d+\./).filter(s => s.trim()).map((s, i) =>
    `<div style="display:flex;gap:8px;margin-bottom:7px">
       <div style="min-width:20px;height:20px;border-radius:50%;background:#060E1F;color:#F0B429;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0">${i+1}</div>
       <div style="font-size:12px;color:#2d2d2d;line-height:1.6">${s.trim()}</div>
     </div>`
  ).join('');

  // Próximos pasos
  const pasosProximos = (analisis.proximos_pasos || '').split(/\d+\./).filter(s => s.trim()).map((s, i) =>
    `<div style="font-size:12px;color:#2d2d2d;line-height:1.6;margin-bottom:5px"><strong>${i+1}.</strong> ${s.trim()}</div>`
  ).join('');

  // Alerta crítica
  const alertaHTML = analisis.alerta_principal && !analisis.alerta_principal.includes('LIMPIO') && !analisis.alerta_principal.includes('limpio')
    ? `<div style="background:#FEF2F2;border-left:3px solid #DC2626;padding:10px 14px;margin-bottom:14px;font-size:12px;color:#7F1D1D"><strong>Punto critico a resolver antes de la cita:</strong> ${analisis.alerta_principal}</div>`
    : `<div style="background:#F0FDF4;border-left:3px solid #16a34a;padding:10px 14px;margin-bottom:14px;font-size:12px;color:#166534">Perfil sin alertas criticas — caso manejable con la preparacion adecuada.</div>`;

  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Informe — ${refId} — ${nombrePrincipal}</title>
<style>
@page{margin:11mm 13mm}
body{font-family:'Segoe UI',Arial,sans-serif;background:#f0ece4;margin:0;padding:20px 16px;color:#1a1a1a}
.doc{max-width:760px;margin:0 auto;background:white;box-shadow:0 6px 32px rgba(0,0,0,.12);overflow:hidden}
p{margin:0 0 10px;line-height:1.65;font-size:13px;color:#2d2d2d}
h3{color:#060E1F;font-size:13px;font-weight:700;border-bottom:2px solid #F0B429;padding-bottom:4px;margin:18px 0 9px}
strong{color:#060E1F}
@media print{body{background:white;padding:0}}
</style>
</head>
<body>
<div class="doc">

<!-- HEADER -->
<div style="background:#060E1F;padding:22px 32px">
  <div style="color:#F0B429;font-size:16px;font-weight:700;margin-bottom:2px">Asesoria Visa Global</div>
  <div style="color:rgba(255,255,255,.4);font-size:9px;text-transform:uppercase;letter-spacing:.12em">Informe de Expediente &middot; Visa USA B1/B2 &middot; Confidencial</div>
  <div style="color:white;font-size:18px;font-weight:700;margin:10px 0 3px">${personas[0] ? personas[0].nombre : 'Cliente'}</div>
  <div style="color:rgba(255,255,255,.5);font-size:11px">Expediente ${refId} &middot; ${fecha} &middot; ${numViaj} viajero${numViaj>1?'s':''}</div>
</div>

<!-- STRIP -->
<div style="display:flex;flex-wrap:wrap;border-bottom:1px solid #e8e0d4">
  <div style="flex:1;min-width:100px;padding:10px 14px;border-right:1px solid #e8e0d4;text-align:center">
    <div style="font-size:18px;font-weight:700;color:${colorProb};margin-bottom:1px">${prob}</div>
    <div style="font-size:9px;text-transform:uppercase;letter-spacing:.07em;color:#999">Aprobacion estimada</div>
  </div>
  <div style="flex:1;min-width:100px;padding:10px 14px;border-right:1px solid #e8e0d4;text-align:center">
    <div style="font-size:13px;font-weight:700;color:#1E40AF;margin-bottom:1px">${consulado}</div>
    <div style="font-size:9px;text-transform:uppercase;letter-spacing:.07em;color:#999">Consulado</div>
  </div>
  <div style="flex:1;min-width:100px;padding:10px 14px;border-right:1px solid #e8e0d4;text-align:center">
    <div style="font-size:12px;font-weight:700;color:#92400E;margin-bottom:1px">${analisis.fecha_viaje_ideal ? analisis.fecha_viaje_ideal.split(' ').slice(0,3).join(' ') : '—'}</div>
    <div style="font-size:9px;text-transform:uppercase;letter-spacing:.07em;color:#999">Fecha viaje ideal</div>
  </div>
  <div style="flex:1;min-width:100px;padding:10px 14px;text-align:center">
    <div style="font-size:13px;font-weight:700;color:#1a1a1a;margin-bottom:1px">${numViaj} viajero${numViaj>1?'s':''}</div>
    <div style="font-size:9px;text-transform:uppercase;letter-spacing:.07em;color:#999">Ref: ${refId}</div>
  </div>
</div>

<!-- BODY -->
<div style="padding:24px 32px">

<p>Estimado/a <strong>${nombrePrincipal}</strong>,</p>
<p>He revisado su expediente en detalle. A continuacion encontrara el plan de trabajo personalizado, los documentos requeridos y la estrategia para maximizar las probabilidades de aprobacion de su visa USA B1/B2.</p>

<div style="background:#FFF9F0;border-left:3px solid #F0B429;padding:10px 14px;margin-bottom:14px;font-size:12px">
  <strong>Grupo que viaja:</strong><br>${viajList}
</div>

${alertaHTML}

<h3>Fortalezas del perfil</h3>
<p>${(analisis.fuertes||'Sin datos').split(';').map(f=>`<strong>+</strong> ${f.trim()}`).join('<br>')}</p>

${analisis.debiles ? `<h3>Puntos a reforzar antes de la cita</h3>
<p>${analisis.debiles.split(';').map(d=>`<strong>&minus;</strong> ${d.trim()}`).join('<br>')}</p>` : ''}

<h3>Estrategia del caso — plan de trabajo</h3>
${pasosEstrategia || `<p>${analisis.estrategia||'—'}</p>`}

<h3>Texto exacto para el DS-160 — campo "Purpose of Trip"</h3>
<div style="background:#FFF7ED;border:2px solid #F0B429;border-radius:8px;padding:14px;margin-bottom:10px;font-size:13px;color:#78350F;font-style:italic;line-height:1.7">
  &ldquo;${analisis.motivo_ds160||'—'}&rdquo;
</div>
<p style="font-size:11px;color:#666">Copiar este texto exactamente en el DS-160 de cada adulto del grupo.</p>

<h3>Documentos requeridos para el expediente</h3>
${docsHTML || `<p>${analisis.documentos_exactos||'—'}</p>`}

${analisis.checklist_pre_cita ? `<h3>Checklist antes de la cita consular</h3>
<p style="font-size:12px;color:#2d2d2d;line-height:1.8">${analisis.checklist_pre_cita.split(';').map((c,i)=>`${i+1}. ${c.trim()}`).join('<br>')}</p>` : ''}

${analisis.tiempo_preparacion ? `<h3>Tiempo de preparacion estimado</h3>
<p style="font-size:12px;color:#2d2d2d;line-height:1.8">${analisis.tiempo_preparacion}</p>` : ''}

${pasosProximos ? `<h3>Proximos pasos — esta semana</h3>${pasosProximos}` : ''}

<p style="margin-top:20px">Quedo a su disposicion para cualquier consulta. Escribame directamente por WhatsApp con su numero de expediente <strong>${refId}</strong> para respuesta inmediata.</p>
<p>Atentamente,<br><strong>Roberto Acosta</strong></p>

</div>

<!-- FOOTER -->
<div style="background:#060E1F;padding:18px 32px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px">
  <div>
    <div style="color:#F0B429;font-weight:700;font-size:13px">Roberto Acosta</div>
    <div style="color:rgba(255,255,255,.5);font-size:11px">Asesor en Visas B1/B2 USA</div>
  </div>
  <div style="color:rgba(255,255,255,.5);font-size:11px;text-align:right">
    WhatsApp: <span style="color:#F0B429">+593 98 784 6751</span><br>
    info@asesoriadevisadosglobal.com<br>
    asesoriadevisadosglobal.com
  </div>
</div>

</div>
</body>
</html>`;
}

function guardarCarta(refId, cartaHTML) {
  try {
    const ss  = SpreadsheetApp.openById(SS_ID);
    const sh  = getOrCreateSheet(ss, 'Cartas', ['Ref ID','Fecha','HTML','Estado']);
    const rows = sh.getDataRange().getValues();
    const idx  = rows.findIndex((r, i) => i > 0 && String(r[0]) === refId);
    const fecha = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');
    const row  = [refId, fecha, cartaHTML, 'Generada'];
    if (idx > 0) { sh.getRange(idx + 1, 1, 1, row.length).setValues([row]); }
    else         { sh.appendRow(row); }
    formatearCabecera(sh, 1);
    return true;
  } catch(e) { console.error('guardarCarta:', e.toString()); return false; }
}

function servirCarta(ref, pin) {
  if (pin !== (ADMIN_PIN || 'visa2026'))
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:30px;color:#991B1B">Acceso denegado</p>');
  try {
    const ss  = SpreadsheetApp.openById(SS_ID);
    const sh  = ss.getSheetByName('Cartas');
    if (!sh) return HtmlService.createHtmlOutput('<p style="padding:20px">No hay cartas generadas aun. Espera a que un cliente envie el formulario.</p>');
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === ref) {
        return HtmlService.createHtmlOutput(rows[i][2])
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    }
    return HtmlService.createHtmlOutput('<p style="padding:20px">Caso ' + ref + ' sin carta generada todavia.</p>');
  } catch(e) {
    return HtmlService.createHtmlOutput('<p style="padding:20px;color:red">Error: ' + e.toString() + '</p>');
  }
}
