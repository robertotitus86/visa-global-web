// ═══════════════════════════════════════════════════════════════════════
// ADDON — Diagnóstico $50 + Chat Widget + Payphone
// Agregar al final de appscript_portal.js ANTES de la función ok()
// ═══════════════════════════════════════════════════════════════════════

// ── Añadir al doGet, antes de la línea "const refId = e.parameter.id;" ──
// if (action === 'payphone_prepare') return payphonePrepare(e);
// if (action === 'payphone_verify')  return payphoneVerify(e);
// ────────────────────────────────────────────────────────────────────────

// ── Añadir al doPost, junto a los otros if de payload.action ────────────
// if (payload.action === 'run_diagnostic') return runDiagnostico(payload);
// if (payload.action === 'chat_message')   return chatMessage(payload);
// ────────────────────────────────────────────────────────────────────────

const PAYPHONE_TOKEN    = PropertiesService.getScriptProperties().getProperty('PAYPHONE_TOKEN');
const PAYPHONE_STORE_ID = PropertiesService.getScriptProperties().getProperty('PAYPHONE_STORE_ID');
const SITE_URL          = 'https://www.asesoriadevisadosglobal.com';

// ════════════════════════════════════════════════════════════════════════
// PAYPHONE — Crear transacción
// ════════════════════════════════════════════════════════════════════════
function payphonePrepare(e) {
  try {
    if (!PAYPHONE_TOKEN) return ok({ error: 'PAYPHONE_TOKEN no configurado en Script Properties' });

    const ref    = e.parameter.ref    || 'DIAG-000000';
    const nombre = e.parameter.nombre || 'Cliente';
    const email  = e.parameter.email  || '';

    const payload = {
      amount:            5000,  // $50.00 en centavos
      amountWithTax:     0,
      tax:               0,
      service:           0,
      tip:               0,
      currency:          'USD',
      clientTransactionId: ref,
      storeId:           PAYPHONE_STORE_ID || null,
      responseUrl:       SITE_URL + '/diagnostico.html',
      cancellationUrl:   SITE_URL + '/diagnostico.html?cancelled=true',
      reference:         'Diagnostico Visa Global — ' + nombre,
      lang:              'es'
    };

    const resp = UrlFetchApp.fetch('https://pay.payphone.com/api/button/Prepare', {
      method:      'post',
      contentType: 'application/json',
      headers:     { Authorization: 'Bearer ' + PAYPHONE_TOKEN },
      payload:     JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const data = JSON.parse(resp.getContentText());
    if (data.payWithCard) {
      // Guardar intento en Sheets
      guardarIntentoPayphone(ref, nombre, email, 'pendiente');
      return ok({ url: data.payWithCard });
    }
    return ok({ error: 'Payphone no devolvio URL: ' + JSON.stringify(data) });
  } catch(err) {
    return ok({ error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════════════
// PAYPHONE — Verificar pago
// ════════════════════════════════════════════════════════════════════════
function payphoneVerify(e) {
  try {
    if (!PAYPHONE_TOKEN) return ok({ approved: false, error: 'Token no configurado' });

    const id = parseInt(e.parameter.id);
    const ct = e.parameter.clientTransactionId;

    const resp = UrlFetchApp.fetch('https://pay.payphone.com/api/button/V2/Confirm', {
      method:      'post',
      contentType: 'application/json',
      headers:     { Authorization: 'Bearer ' + PAYPHONE_TOKEN },
      payload:     JSON.stringify({ id, clientTransactionId: ct }),
      muteHttpExceptions: true
    });

    const data = JSON.parse(resp.getContentText());
    const aprobado = data.transactionStatus === 'Approved';

    if (aprobado) guardarIntentoPayphone(ct, '', '', 'pagado', id);
    return ok({ approved: aprobado, status: data.transactionStatus });
  } catch(err) {
    return ok({ approved: false, error: err.toString() });
  }
}

function guardarIntentoPayphone(ref, nombre, email, estado, ppId) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let sh = ss.getSheetByName('Pagos Diagnostico');
    if (!sh) {
      sh = ss.insertSheet('Pagos Diagnostico');
      sh.appendRow(['Fecha','Ref','Nombre','Email','Estado','PP_ID','Monto']);
      sh.getRange(1,1,1,7).setBackground('#060E1F').setFontColor('#F0B429').setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const existing = sh.getDataRange().getValues();
    const row = existing.findIndex(r => r[1] === ref);
    const fecha = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');
    if (row > 0) {
      sh.getRange(row + 1, 5).setValue(estado);
      if (ppId) sh.getRange(row + 1, 6).setValue(ppId);
    } else {
      sh.appendRow([fecha, ref, nombre || '', email || '', estado, ppId || '', '$50.00']);
    }
  } catch(e) { console.error('guardarIntentoPayphone:', e); }
}

// ════════════════════════════════════════════════════════════════════════
// DIAGNÓSTICO — Analizar perfil con Gemini y guardar
// ════════════════════════════════════════════════════════════════════════
function runDiagnostico(payload) {
  try {
    const client  = payload.client  || {};
    const answers = payload.answers || {};

    const prompt = buildDiagnosticoPrompt(client, answers);
    const raw    = callGeminiChat(prompt);
    const report = parseDiagnosticoJSON(raw);

    // Guardar en Sheets
    guardarDiagnostico(client, answers, report);

    // Email al cliente
    enviarEmailDiagnostico(client, answers, report);

    // Alerta a Roberto
    alertarRobertoDiagnostico(client, answers, report);

    return ok({ ok: true, report });
  } catch(err) {
    console.error('runDiagnostico:', err);
    return ok({ ok: false, error: err.toString() });
  }
}

function buildDiagnosticoPrompt(client, a) {
  return `Eres un experto en visas con 20 años de experiencia en el consulado americano en Ecuador.
Analiza el siguiente perfil y genera un diagnostico completo en JSON. Responde SOLO JSON sin markdown.

PERFIL DEL SOLICITANTE:
- Nombre: ${client.name || 'No indicado'}
- Destino: ${a.destino || '-'}
- Urgencia del viaje: ${a.urgencia || '-'}
- Edad: ${a.edad || '-'}
- Estado civil: ${a.civil || '-'}
- Hijos menores en Ecuador: ${a.hijos || '-'}
- Situacion laboral: ${a.empleo || '-'}
- Antiguedad laboral: ${a.antiguedad || '-'}
- Ingresos mensuales: ${a.ingresos || '-'}
- Bienes propios: ${a.bienes || '-'}
- Historial de visa: ${a.historial || '-'}
- Visas vigentes de otros paises: ${a.visa_otra || '-'}
- Viajes al exterior recientes: ${a.viajes || '-'}
- Familiares en USA/Europa: ${a.familiares || '-'}
- Historial bancario: ${a.banco || '-'}
- Motivo del viaje: ${a.motivo || '-'}

CONTEXTO IMPORTANTE:
- Ecuador tiene 42% de tasa de rechazo de visa USA en 2025 (record historico).
- El consulado busca ARRAIGO: trabajo, bienes, familia, raices en Ecuador.
- Un rechazo previo pesa mucho en expediente futuro.
- La especialidad del asesor es casos de rechazo previo y consulado español.

Genera el JSON con esta estructura exacta:
{
  "fortalezas": ["punto 1", "punto 2", "punto 3"],
  "riesgos": ["riesgo 1", "riesgo 2", "riesgo 3"],
  "documentos": ["documento 1", "documento 2", "documento 3", "documento 4", "documento 5"],
  "analisis": "parrafo de 3-4 oraciones explicando el caso de forma directa y honesta",
  "riesgo_label": "RIESGO ALTO|RIESGO MEDIO|RIESGO BAJO",
  "riesgo_nivel": "alto|medio|bajo",
  "paquete": "esencial|profesional|vip",
  "paquete_razon": "una oracion explicando por que ese paquete"
}

Reglas:
- Si hay rechazo previo: paquete = "vip" obligatoriamente
- Si hay familiares sin documentos en USA: riesgo_nivel = "alto"
- Si ingresos < $500 y sin bienes: riesgo = "alto"
- Sé honesto, directo y util. No suavices los problemas reales.
- Los documentos deben ser especificos al destino (${a.destino || 'USA'}).`;
}

function parseDiagnosticoJSON(raw) {
  try {
    const clean = raw.replace(/```json|```/g, '').trim();
    const match = clean.match(/\{[\s\S]*\}/);
    if (!match) throw new Error('Sin JSON');
    return JSON.parse(match[0]);
  } catch(e) {
    return {
      fortalezas: ['Perfil recibido y procesado'],
      riesgos:    ['No se pudo analizar completamente — Roberto revisara manualmente'],
      documentos: ['Pasaporte vigente', 'Estados de cuenta (3 meses)', 'Carta de empleo', 'Comprobante de bienes', 'Fotos recientes'],
      analisis:   'El sistema proceso tus respuestas. Roberto revisara tu expediente y te contactara en las proximas horas.',
      riesgo_label: 'EN REVISION',
      riesgo_nivel: 'medio',
      paquete:    'profesional',
      paquete_razon: 'Recomendacion pendiente de revision manual.'
    };
  }
}

function guardarDiagnostico(client, answers, report) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let sh = ss.getSheetByName('Diagnosticos');
    if (!sh) {
      sh = ss.insertSheet('Diagnosticos');
      sh.appendRow(['Fecha','Ref','Nombre','Email','WhatsApp','Destino','Riesgo','Paquete','Fortalezas','Riesgos','Documentos','Analisis']);
      sh.getRange(1,1,1,12).setBackground('#060E1F').setFontColor('#F0B429').setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const fecha = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');
    sh.appendRow([
      fecha,
      client.ref   || '',
      client.name  || '',
      client.email || '',
      client.phone || '',
      answers.destino || '',
      report.riesgo_label || '',
      report.paquete || '',
      (report.fortalezas || []).join(' | '),
      (report.riesgos    || []).join(' | '),
      (report.documentos || []).join(' | '),
      report.analisis || ''
    ]);
  } catch(e) { console.error('guardarDiagnostico:', e); }
}

function enviarEmailDiagnostico(client, answers, report) {
  try {
    if (!client.email) return;
    const pkgNombres = { esencial:'Esencial ($197)', profesional:'Profesional ($250)', vip:'VIP Rechazo ($320)' };
    const pNombre = pkgNombres[report.paquete] || 'Profesional ($250)';
    const rColor  = report.riesgo_nivel === 'alto' ? '#dc2626' : report.riesgo_nivel === 'bajo' ? '#059669' : '#d97706';

    const html = `
<div style="font-family:'DM Sans',Arial,sans-serif;max-width:600px;margin:0 auto;color:#111827">
  <div style="background:#060E1F;padding:2rem;border-radius:0.75rem 0.75rem 0 0;text-align:center">
    <h1 style="font-family:Georgia,serif;color:#fff;font-size:1.4rem;margin:0 0 0.25rem">Expediente de Diagnostico</h1>
    <p style="color:rgba(255,255,255,0.6);font-size:0.85rem;margin:0">Asesoria Visa Global · ${Utilities.formatDate(new Date(),'America/Guayaquil','dd/MM/yyyy')}</p>
  </div>
  <div style="background:#fff;padding:1.75rem;border:1px solid #f3f4f6;border-top:none">
    <p style="font-size:0.95rem;color:#374151">Hola <strong>${client.name}</strong>, aqui esta el analisis de tu perfil para <strong>${answers.destino || 'visa solicitada'}</strong>.</p>
    <div style="background:#f9fafb;border-radius:0.5rem;padding:1rem;margin:1.25rem 0;text-align:center">
      <div style="background:${rColor};color:#fff;display:inline-block;padding:0.25rem 1rem;border-radius:999px;font-size:0.8rem;font-weight:700;letter-spacing:0.05em">${report.riesgo_label || 'EN ANALISIS'}</div>
    </div>
    <h2 style="font-size:0.8rem;text-transform:uppercase;letter-spacing:0.1em;color:#9ca3af;margin:1.5rem 0 0.5rem">Fortalezas</h2>
    <ul style="margin:0;padding-left:1.25rem;color:#374151;font-size:0.9rem">${(report.fortalezas||[]).map(f=>`<li style="margin-bottom:0.3rem">${f}</li>`).join('')}</ul>
    <h2 style="font-size:0.8rem;text-transform:uppercase;letter-spacing:0.1em;color:#9ca3af;margin:1.5rem 0 0.5rem">Factores de riesgo</h2>
    <ul style="margin:0;padding-left:1.25rem;color:#374151;font-size:0.9rem">${(report.riesgos||[]).map(r=>`<li style="margin-bottom:0.3rem">${r}</li>`).join('')}</ul>
    <h2 style="font-size:0.8rem;text-transform:uppercase;letter-spacing:0.1em;color:#9ca3af;margin:1.5rem 0 0.5rem">Documentos prioritarios</h2>
    <ul style="margin:0;padding-left:1.25rem;color:#374151;font-size:0.9rem">${(report.documentos||[]).map(d=>`<li style="margin-bottom:0.3rem">${d}</li>`).join('')}</ul>
    <div style="background:#f9fafb;border-left:3px solid #F0B429;padding:1rem 1.25rem;margin:1.5rem 0;border-radius:0 0.5rem 0.5rem 0">
      <p style="margin:0;font-size:0.9rem;color:#374151">${report.analisis || ''}</p>
    </div>
    <div style="background:#060E1F;border-radius:0.75rem;padding:1.5rem;text-align:center;margin-top:1.5rem">
      <p style="color:rgba(255,255,255,0.7);font-size:0.85rem;margin:0 0 0.5rem">Paquete recomendado para tu caso</p>
      <p style="color:#F0B429;font-family:Georgia,serif;font-size:1.4rem;font-weight:700;margin:0 0 0.5rem">${pNombre}</p>
      <p style="color:rgba(255,255,255,0.6);font-size:0.8rem;margin:0 0 1.25rem">${report.paquete_razon || ''}</p>
      <a href="https://wa.me/593994442512?text=${encodeURIComponent('Hola Roberto, recibi mi diagnostico (ref: ' + (client.ref||'') + ') y quiero contratar el ' + pNombre)}" style="display:inline-block;background:#25D366;color:#fff;padding:0.65rem 1.75rem;border-radius:0.5rem;text-decoration:none;font-weight:700;font-size:0.9rem">Contratar asesoria por WhatsApp</a>
    </div>
    <p style="font-size:0.78rem;color:#9ca3af;text-align:center;margin-top:1.25rem">Ref: ${client.ref || ''} · Roberto Acosta · asesoriadevisadosglobal.com</p>
  </div>
</div>`;

    MailApp.sendEmail({ to: client.email, subject: 'Tu expediente de diagnostico — Asesoria Visa Global', htmlBody: html });
  } catch(e) { console.error('enviarEmailDiagnostico:', e); }
}

function alertarRobertoDiagnostico(client, answers, report) {
  try {
    const emoji = report.riesgo_nivel === 'bajo' ? '🟢' : report.riesgo_nivel === 'alto' ? '🔴' : '🟡';
    MailApp.sendEmail({
      to: EMAIL_ROBERTO,
      subject: emoji + ' Nuevo diagnostico: ' + (client.name||'Sin nombre') + ' — ' + (report.paquete||'').toUpperCase(),
      body:
        'NUEVO DIAGNOSTICO PAGADO\n\n' +
        'Nombre:   ' + (client.name||'') + '\n' +
        'Email:    ' + (client.email||'') + '\n' +
        'WhatsApp: ' + (client.phone||'') + '\n' +
        'Ref:      ' + (client.ref||'') + '\n' +
        'Destino:  ' + (answers.destino||'') + '\n' +
        'Riesgo:   ' + (report.riesgo_label||'') + '\n' +
        'Paquete:  ' + (report.paquete||'').toUpperCase() + '\n\n' +
        'Analisis: ' + (report.analisis||'') + '\n\n' +
        'Riesgos: ' + (report.riesgos||[]).join(' | ')
    });
  } catch(e) { console.error('alertarRobertoDiagnostico:', e); }
}

// ════════════════════════════════════════════════════════════════════════
// CHAT WIDGET — Respuestas IA con Gemini
// ════════════════════════════════════════════════════════════════════════
function chatMessage(payload) {
  try {
    const message = payload.message || '';
    const hist    = payload.history || [];

    const systemContext = `Eres el asistente virtual de Asesoria Visa Global, empresa de Roberto Acosta (espanol con conocimiento interno del consulado).
Tu rol: atender la PRIMERA CONSULTA GRATUITA de prospectos ecuatorianos que quieren visa USA, España, Schengen o UK.

SERVICIOS Y PRECIOS:
- Diagnostico de perfil: $50 (analisis IA del caso + expediente PDF + recomendacion)
- Paquete Esencial: $197 / persona (primer intento, perfil estandar)
- Paquete Profesional: $250 / persona (perfil complejo, servicio completo)
- Paquete VIP Rechazo: $320 / persona (para casos con rechazo previo)
- Familia (mismo caso): $250 / persona adicional desde el 2do miembro

VENTAJA COMPETITIVA: Roberto es español. Conoce el consulado desde adentro. Especialista en casos de rechazo.
CONTEXTO: Ecuador tiene 42% de tasa de rechazo de visa USA (2025, record historico).

REGLAS DE CONVERSACION:
1. Responde en espanol, de forma calida y directa. Sin emojis.
2. Da respuestas cortas (2-4 oraciones maximo).
3. Despues de 3-4 intercambios, ofrece el Diagnostico de $50 como siguiente paso.
4. Para ofrecer el diagnostico usa: <a href="/diagnostico.html" class="vg-cta-btn">Obtener diagnostico — $50</a>
5. Si pregunta por precio directamente, da los precios arriba.
6. Si el caso es urgente o complejo, ofrece hablar con Roberto directamente.
7. NUNCA inventes informacion sobre requisitos especificos de visa — di que Roberto lo revisara.
8. Si detectas rechazo previo, mencionalo como especialidad de Roberto.
9. quick_replies: sugiere 2-3 respuestas rapidas relevantes al contexto.

FORMATO DE RESPUESTA — JSON exacto:
{"reply": "texto del mensaje (puede incluir HTML)", "quick_replies": ["opcion 1", "opcion 2"]}`;

    const geminiHistory = hist.slice(-8).map(h => ({
      role:  h.role,
      parts: h.parts
    }));

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + GEMINI_KEY;
    const body = {
      systemInstruction: { parts: [{ text: systemContext }] },
      contents: [...geminiHistory, { role: 'user', parts: [{ text: message }] }],
      generationConfig: { temperature: 0.5, maxOutputTokens: 400 }
    };

    const resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(body), muteHttpExceptions: true
    });
    const gemData = JSON.parse(resp.getContentText());
    const rawText = gemData.candidates[0].content.parts[0].text;

    let parsed;
    try {
      const match = rawText.replace(/```json|```/g,'').trim().match(/\{[\s\S]*\}/);
      parsed = match ? JSON.parse(match[0]) : { reply: rawText, quick_replies: [] };
    } catch(_) {
      parsed = { reply: rawText, quick_replies: [] };
    }

    // Guardar conversacion
    guardarChat(payload.session || '', message, parsed.reply);

    return ok({ reply: parsed.reply, quick_replies: parsed.quick_replies || [] });
  } catch(err) {
    return ok({ reply: 'Disculpa, tuve un problema tecnico. Puedes escribirnos directamente al WhatsApp: <a href="https://wa.me/593994442512" target="_blank">+593 99 444 2512</a>', quick_replies: [] });
  }
}

function guardarChat(session, userMsg, botReply) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let sh = ss.getSheetByName('Chat Logs');
    if (!sh) {
      sh = ss.insertSheet('Chat Logs');
      sh.appendRow(['Fecha','Session','Usuario','Bot']);
      sh.getRange(1,1,1,4).setBackground('#060E1F').setFontColor('#F0B429').setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const fecha = Utilities.formatDate(new Date(), 'America/Guayaquil', 'dd/MM/yyyy HH:mm');
    sh.appendRow([fecha, session, userMsg, botReply]);
  } catch(e) {}
}

function callGeminiChat(prompt) {
  const url  = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + GEMINI_KEY;
  const resp = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2, maxOutputTokens: 1500 } }),
    muteHttpExceptions: true
  });
  const data = JSON.parse(resp.getContentText());
  return data.candidates[0].content.parts[0].text;
}
