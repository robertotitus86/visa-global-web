# INSTRUCCIONES — visa-global-web (CRM + Simuladores + CDN)

## QUÉ ES ESTE PROYECTO
Sitio web de Asesoría Visa Global desplegado en GitHub Pages.
URL: https://www.asesoriadevisadosglobal.com
GitHub: robertotitus86/visa-global-web → push a `main` → deploy automático

## ARCHIVOS CLAVE

| Archivo | Función |
|---------|---------|
| `admin.html` | CRM — gestión de casos de clientes |
| `portal.html` | Portal del cliente (acceso con Ref ID) |
| `simulador.html` | Simulador de elegibilidad para visa USA |
| `diagnostico.html` | Diagnóstico inicial del caso |
| `intake.html` | Formulario de entrada de nuevos clientes |
| `screening.html` | Screening previo al onboarding |
| `ds160.html` | Guía DS-160 interactiva |
| `index.html` | Landing page principal |

## APP SCRIPTS (Google Sheets backend)

| Archivo | Función |
|---------|---------|
| `appscript_followups.js` | Seguimiento automático de casos |
| `appscript_ds160.js` | Procesamiento DS-160 |
| `appscript_portal.js` | Lógica del portal cliente |
| `appscript_diagnostico_addon.js` | Addon diagnóstico |

## CASOS ACTIVOS (simuladores individuales)
- `familia-seas-guaman.html` + `plan-seas-guaman.pdf` + `checklist-seas-guaman.pdf`
- `familia-rodriguez-masache.html` + `plan-rodriguez-masache.pdf` + `checklist-rodriguez-masache.pdf`

## USO COMO CDN (GitHub Pages para contenido Meta)
La carpeta `archivo/` sirve como CDN público para imágenes y videos de FB/IG.
- Siempre nombres ÚNICOS con fecha: `ig20260617_N.jpg`, `reel20260617_N.mp4`
- NUNCA reutilizar nombres genéricos (cig1.png, ig1.jpg, etc.) — GitHub Pages cachea y sirve versión vieja
- Verificar disponibilidad con HEAD request (esperar 200) antes de usar en API de Meta
- Push a main → esperar ~15s → disponible en https://www.asesoriadevisadosglobal.com/archivo/

## BUGS CRÍTICOS CONOCIDOS

### localStorage se destruye con sync — CRÍTICO
El botón "Sincronizar" en admin.html NUNCA debe sobreescribir localStorage con array vacío del servidor.
SIEMPRE hacer merge (no sobreescribir):
```javascript
const serverRefs = new Set(fromServer.map(c => c['Ref ID']));
const onlyLocal = CASES.filter(c => !serverRefs.has(c['Ref ID']));
CASES = [...fromServer, ...onlyLocal];
```

### Casos deben estar en Google Sheets SIEMPRE
Nunca confiar solo en localStorage — se borra con cache o cambia de dispositivo.
Al crear un caso: llamar SIEMPRE al Apps Script `?action=newCase`.

### Orden save/validate
SIEMPRE: validateAndContinue() ANTES de save(). Nunca al revés o se pierden los _pendientes.

### Datos de PDF — guardar inmediatamente
Después de extraer datos de PDF con IA: guardar a S.perTraveler INMEDIATAMENTE, no esperar clic en Continuar.

## CREDENCIALES / CONFIGURACIÓN
- GitHub repo: robertotitus86/visa-global-web
- Deploy: GitHub Pages automático en push a main
- Google Sheets backend: configurado en los appscript_*.js
- Pagos: PayPal (ver configuración en portal.html)
- Dominio: asesoriadevisadosglobal.com (CNAME en repo)
