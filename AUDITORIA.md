# AUDITORÍA — Asesoría Visa Global
Fecha: 2026-09-11

## 1. Mapa del sitio y stack
- **Hosting:** GitHub Pages (repo `visa-global-web`, rama `main`), dominio propio vía `CNAME`.
- **Frontend:** HTML estático + CSS (`design.css`) + JS inline, sin framework ni build step.
- **Backend:** Google Apps Script (`appscript_portal.js`, ~184 KB — archivo grande, probable candidato a dividir en el futuro, no ahora).
- **Pagos:** PayPhone vía `pagar.html` (redirect simple) → `pago-exitoso.html` / `pago-cancelado.html`.
- **WhatsApp:** enlaces directos `wa.me/593994442512` con texto prellenado por sección (no Business API en el sitio — el bot de WhatsApp es un servicio aparte en Render/`visa-bot`).
- **CRM:** `admin.html` (kanban local, PIN), respaldado por Google Sheets.
- **Páginas activas relevantes:** `index.html`, `diagnostico.html`, `visa-usa-ecuador.html`, `visa-espana-ecuatorianos.html`, `visa-canada-ecuatorianos.html`, `visa-rechazada-que-hacer.html`, `migracion-circular-ecuador.html`, `asesoria-visas-quito.html`, `asesoria-visas-guayaquil.html`, `portal.html`, `intake.html`/`intake-ds160.html`, `simulador.html`, `screening.html`, `privacy.html`, `terms.html`.
- Hay bastantes archivos de casos individuales (brochures, informes) en la raíz — no afectan SEO negativamente porque no están en el sitemap, pero ensucian el repo.

## 2. SEO técnico
- **Title/meta description:** presentes y bien escritos en `index.html`.
- **Schema.org (JSON-LD):** ya implementado — 6 bloques en `index.html` (probablemente Organization/Service/FAQPage). Buena base.
- **robots.txt:** correcto — permite todo, bloquea `/admin.html` y `/pagar.html`, referencia el sitemap.
- **sitemap.xml:** existe pero **incompleto** — le faltan `asesoria-visas-quito.html`, `asesoria-visas-guayaquil.html` (❌ no, revisar: sí están listadas), `simulador.html`, `screening.html`. `lastmod` desactualizado (todo en junio, sin reflejar cambios de agosto).
- **Encoding roto detectado en `diagnostico.html`:** caracteres como `Â·`, `âœ“`, `â€”` en vez de tildes/comillas — típico de guardar UTF-8 mal interpretado. Afecta percepción de profesionalismo justo en la página de pago del diagnóstico.
- **Velocidad/imágenes:** hay decenas de imágenes pesadas (varias >1MB, algunas de 7-14MB como `brochure-disney.jpg`, `brochure-kennedy.jpg`) sueltas en la raíz del repo — no confirmé si `index.html` las carga, pero si el sitio referencia imágenes sin optimizar en cualquier página, penaliza velocidad y por tanto SEO/conversión móvil.

## 3. Medición — mejor de lo esperado, con un hueco real
- **GA4 instalado:** sí (`G-YBSM0X9L5E`), en `index.html`.
- **Meta Pixel instalado:** sí (`1713317820112908`), con `PageView` disparando.
- **Conversions API:** no verificado (requeriría revisar Apps Script/backend; no hay señal en el HTML).
- **Hueco importante:** los ~15 enlaces `wa.me` del sitio (CTA principal en TODO el funnel) **no disparan ningún evento** (`fbq('track','Contact')` o `gtag('event', ...)`). Hoy, Meta/GA4 no saben cuándo alguien hace clic para hablar por WhatsApp — es la conversión más importante del negocio y no se está midiendo. Esto es barato de arreglar y de alto impacto para cualquier campaña futura.

## 4. Embudo del diagnóstico IA ($50, nunca vendido)
- Estructura completa: `diagnostico.html` → 15 preguntas → análisis → PDF con paquete recomendado → pago vía PayPhone.
- Técnicamente está bien armado (por eso no achacaría el fracaso al funnel en sí), pero:
  - Nunca ha tenido tráfico pagado ni orgánico dirigido específicamente a esta página (confirmado: Instagram con 55 seguidores tras 3 meses).
  - El encoding roto mencionado arriba está justo aquí, en la página de venta.
  - Es un producto de "IA" en un nicho donde la **confianza** es la barrera #1 (según tu propio brief) — pedir $50 antes de hablar con un humano puede ser más fricción que ayuda para gente que ya desconfía de asesorías de visa.

## 5. Bloque de confianza — no existe
- Búsqueda de "RUC", "factura electrónica", "dirección" en `index.html`, `privacy.html`, `terms.html`: **cero resultados**.
- No hay contrato de servicio descargable.
- No hay página "cómo reconocer una estafa de visas".
- Sí existe ya `visa-rechazada-que-hacer.html` — cubre parte de lo que pide la FASE 2, pero está enfocado solo como página informativa, no como landing de conversión con el posicionamiento nuevo de "segunda oportunidad" que quieres construir.

---

## LAS 5 ACCIONES DE MAYOR IMPACTO (por esfuerzo vs resultado)

1. **Arreglar el encoding roto en `diagnostico.html`** — 15 minutos, cero riesgo, mejora inmediata de percepción profesional en la única página que vende algo directamente.
2. **Agregar tracking de eventos a los clics de WhatsApp** (`fbq('track','Contact')` + `gtag('event','whatsapp_click')` en los ~15 enlaces `wa.me`) — sin esto, cualquier campaña de ads que lancemos vuela a ciegas. Es la base de todo lo que pide la FASE 3.
3. **Agregar el bloque de confianza verificable** (foto/nombre de Roberto — ya está parcialmente, RUC, factura SRI, dirección, contrato descargable) — bajo esfuerzo, ataca directamente la barrera #1 de compra que tú mismo identificaste.
4. **Reemplazar/despriorizar el diagnóstico IA de $50 por la oferta "Revisión de caso con Roberto"** (~$25-30, videollamada 20 min, abonable al paquete) como producto de entrada — más alineado con "la confianza es la barrera principal": hablar con una persona real baja más la fricción que un test de IA.
5. **Actualizar `sitemap.xml`** con fechas reales y páginas faltantes, y decidir si `visa-rechazada-que-hacer.html` se convierte en la landing `/negaron-visa` con las dos rutas (EE.UU. 214(b) / Schengen) o si se crea nueva y se redirige la vieja.

Faltan por decidir de tu parte (marco como `[[DECIDIR]]` en el código cuando toque):
- RUC / razón social exacta para facturación.
- Dirección física a mostrar (o si se mantiene "atendemos remoto, sin oficina física").
- Precio final de la "Revisión de caso con Roberto".
- Monto de recompensa por referido.

¿Apruebas empezar por los puntos 1 y 2 (encoding + tracking de WhatsApp), que son los de menor riesgo y esfuerzo, antes de tocar precios/ofertas?
