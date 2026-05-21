/**
 * NODO CODE — n8n
 * Pega este código en el nodo "Code" de tu workflow de n8n.
 * El webhook recibe el JSON del screening.html y este código:
 *   1. Valida y normaliza el payload
 *   2. Recalcula el score (verificación servidor)
 *   3. Determina la ruta (A / B / C)
 *   4. Prepara los mensajes de WhatsApp y email personalizados
 *   5. Retorna el objeto enriquecido para los nodos siguientes
 */

const data = $input.first().json;

// ── 1. SCORING (recálculo servidor — no confiar solo en el cliente) ──────────

const LABORAL = {
  empleado_publico:        35,
  empleado_privado_plus2:  30,
  negocio_propio_plus3:    28,
  empleado_privado_menos2: 15,
  negocio_propio_menos1:   10,
  desempleado:              0,
};

const BIEN = {
  casa:    15,
  terreno: 12,
  vehiculo: 5,
  ninguno:  0,
};

const INGRESOS = {
  mas_2000:  10,
  '1000_2000': 7,
  '500_999':   3,
  menos_500:   0,
};

const CREDITO = {
  si_hipotecario: 10,
  si_consumo:      5,
  no:              0,
};

const CIVIL = {
  casado:          10,
  union_hecho:      8,
  divorciado_hijos: 6,
  soltero:          0,
};

const HIJOS = {
  dos_o_mas: 10,
  uno:        7,
  ninguno:    0,
};

const VIAJES = {
  sello_entrada:    20,
  visa_sin_viajar:  10,
  ninguno:           0,
};

const RECHAZOS = {
  nunca:       10,
  mas_3_anos:   5,
  menos_1_ano: -5,
  dos_o_mas:  -15,
};

const FAMILIARES_INDOG = {
  no:   0,
  si: -10,
};

// Calcular sub-scores
const scoreLaboral    = LABORAL[data.situacion_laboral]         ?? 0;
const scoreFinanciero = (BIEN[data.bien_principal]              ?? 0)
                      + (INGRESOS[data.ingresos_rango]          ?? 0)
                      + (CREDITO[data.credito_activo]           ?? 0);
const scoreFamiliar   = (CIVIL[data.estado_civil]               ?? 0)
                      + (HIJOS[data.hijos_ecuador]              ?? 0);
const scoreHistorial  = (VIAJES[data.viajes_previos]            ?? 0)
                      + (RECHAZOS[data.historial_rechazos]      ?? 0)
                      + (FAMILIARES_INDOG[data.familiares_indocumentados] ?? 0);

const scoreRaw   = scoreLaboral + scoreFinanciero + scoreFamiliar + scoreHistorial;
const scoreTotal = Math.max(0, Math.min(100, Math.round((scoreRaw / 100) * 100)));

// ── 2. RUTEO ────────────────────────────────────────────────────────────────

let ruta, riesgo214b, prioridad;

if (scoreTotal >= 70) {
  ruta       = 'C';
  riesgo214b = 'BAJO';
  prioridad  = 'HOT';          // cierre directo
} else if (scoreTotal >= 40) {
  ruta       = 'B';
  riesgo214b = 'MEDIO';
  prioridad  = 'WARM';         // vender consulta estratégica
} else {
  ruta       = 'A';
  riesgo214b = 'ALTO';
  prioridad  = 'NURTURE';      // nutrir y educar
}

// ── 3. MENSAJES PERSONALIZADOS ──────────────────────────────────────────────

const nombre   = data.nombre || 'cliente';
const visaTipo = data.visa_tipo === 'USA_B1B2' ? 'visa USA (B1/B2)' : 'visa Schengen';

const MENSAJES = {
  A: {
    whatsapp: `Hola ${nombre}, gracias por completar tu evaluación de perfil en Asesoría Visa Global.\n\n` +
              `Tu score de viabilidad es *${scoreTotal}/100*.\n\n` +
              `⚠️ Tu perfil tiene un riesgo elevado de rechazo bajo la Sección 214(b). ` +
              `Aplicar ahora implicaría gastar la tasa consular ($185) con alta probabilidad de rechazo.\n\n` +
              `La buena noticia: este riesgo _sí se puede reducir_. ` +
              `Te envío una guía gratuita con los pasos concretos para fortalecer tu arraigo antes de aplicar.\n\n` +
              `Cuando tu perfil esté listo, te avisamos. Asesoría Visa Global.`,
    email_subject: `Tu evaluación de perfil — ${riesgo214b} riesgo 214(b)`,
    email_body: `Hola ${nombre},\n\nRevisamos tu perfil y tu score de viabilidad es ${scoreTotal}/100.\n\n` +
                `Existe un riesgo alto de rechazo en tu solicitud de ${visaTipo} en este momento.\n\n` +
                `Esto no significa que no puedas obtener la visa — significa que necesitas ` +
                `fortalecer tu perfil antes de aplicar. Te adjuntamos nuestra guía gratuita.\n\n` +
                `Cuando estés listo para una evaluación completa, escríbenos: +593 98 784 6751\n\n` +
                `Primera consulta siempre gratis.\n— Asesoría Visa Global`
  },
  B: {
    whatsapp: `Hola ${nombre}, gracias por tu evaluación en Asesoría Visa Global.\n\n` +
              `Tu score de viabilidad es *${scoreTotal}/100*.\n\n` +
              `📋 Tu caso *tiene viabilidad*, pero el oficial consular buscará inconsistencias en tu expediente. ` +
              `Los casos como el tuyo no se aprueban solos — requieren una *estrategia sólida y bien estructurada*.\n\n` +
              `Podemos trabajar en eso. ¿Cuándo te va bien para una llamada de evaluación? ` +
              `La primera consulta es completamente gratis.\n\n` +
              `Asesoría Visa Global — +593 98 784 6751`,
    email_subject: `Tu perfil tiene viabilidad — siguiente paso`,
    email_body: `Hola ${nombre},\n\nTu score de viabilidad es ${scoreTotal}/100 para tu ${visaTipo}.\n\n` +
                `Tu caso tiene puntos fuertes pero también vulnerabilidades que un oficial consular ` +
                `puede usar en tu contra. Con la estrategia correcta, la probabilidad de aprobación ` +
                `sube significativamente.\n\n` +
                `Agenda tu consulta gratuita aquí: https://asesoriadevisadosglobal.com/screening.html\n\n` +
                `— Asesoría Visa Global`
  },
  C: {
    whatsapp: `Hola ${nombre} 🎯 Tu evaluación en Asesoría Visa Global muestra un *perfil de arraigo sólido*.\n\n` +
              `Tu score de viabilidad es *${scoreTotal}/100*.\n\n` +
              `✅ Cumples con los criterios que busca el consulado. Los tiempos de espera para citas ` +
              `están aumentando esta semana. Te conviene iniciar tu trámite *ahora* para asegurar tu fecha.\n\n` +
              `¿Empezamos hoy? Escríbenos a este mismo número.\n— Asesoría Visa Global`,
    email_subject: `¡Tu perfil está listo para aplicar! — Siguiente paso`,
    email_body: `Hola ${nombre},\n\nExcelente noticia: tu score de viabilidad es ${scoreTotal}/100.\n\n` +
                `Tu perfil cumple con los criterios de arraigo que evalúan los consulados para ${visaTipo}. ` +
                `Podemos iniciar tu trámite esta semana.\n\n` +
                `Responde este email o escríbenos por WhatsApp: +593 98 784 6751\n\n` +
                `Primera consulta gratis. Empezamos cuando tú quieras.\n— Asesoría Visa Global`
  }
};

const mensajes = MENSAJES[ruta];

// ── 4. OUTPUT ENRIQUECIDO ───────────────────────────────────────────────────

return [{
  json: {
    // — datos originales —
    ...data,
    // — scoring recalculado servidor —
    score_laboral_srv:    scoreLaboral,
    score_financiero_srv: scoreFinanciero,
    score_familiar_srv:   scoreFamiliar,
    score_historial_srv:  scoreHistorial,
    score_total_srv:      scoreTotal,
    // — decisión de ruteo —
    ruta,
    riesgo_214b:    riesgo214b,
    prioridad,
    // — mensajes listos para usar —
    msg_whatsapp:   mensajes.whatsapp,
    msg_email_subj: mensajes.email_subject,
    msg_email_body: mensajes.email_body,
    // — metadata —
    procesado_en: new Date().toISOString(),
  }
}];
