# Graph Report - .  (2026-07-27)

## Corpus Check
- Large corpus: 299 files · ~6,735,370 words. Semantic extraction will be expensive (many Claude tokens). Consider running on a subfolder.

## Summary
- 193 nodes · 317 edges · 13 communities
- Extraction: 88% EXTRACTED · 11% INFERRED · 1% AMBIGUOUS · INFERRED: 35 edges (avg confidence: 0.81)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- Community 0
- Community 1
- Community 2
- Community 3
- Community 4
- Community 5
- Community 6
- Community 7
- Community 8
- Community 9
- Community 10
- Community 11

## God Nodes (most connected - your core abstractions)
1. `ok()` - 20 edges
2. `doGet()` - 17 edges
3. `doPost()` - 13 edges
4. `manejarIntakeFamiliar()` - 13 edges
5. `admin.html (CRM)` - 12 edges
6. `visa-global-web Project Instructions` - 10 edges
7. `manejarPortalLegacy()` - 9 edges
8. `getOrCreateSheet()` - 9 edges
9. `formatearCabecera()` - 9 edges
10. `runDiagnostico()` - 9 edges

## Surprising Connections (you probably didn't know these)
- `Guía Proceso Visa USA (PDF)` --conceptually_related_to--> `ds160.html (Interactive DS-160 Guide)`  [AMBIGUOUS]
  guia-proceso-visa-usa.pdf → ds160.html
- `Guía Proceso Visa USA (PDF)` --conceptually_related_to--> `simulador.html (Eligibility Simulator)`  [AMBIGUOUS]
  guia-proceso-visa-usa.pdf → simulador.html
- `ds160.html (Interactive DS-160 Guide)` --semantically_similar_to--> `intake-ds160.html (DS-160 Intake Form)`  [INFERRED] [semantically similar]
  ds160.html → intake-ds160.html
- `intake.html (New Client Intake Form)` --semantically_similar_to--> `goStep`  [INFERRED] [semantically similar]
  intake.html → screening.html
- `visa-global-web Project Instructions` --references--> `screening.html (Pre-onboarding Screening)`  [EXTRACTED]
  CLAUDE.md → screening.html

## Import Cycles
- None detected.

## Hyperedges (group relationships)
- **Client Intake-to-CRM Pipeline** — intake_html, screening_gostep, diagnostico_enviardiagnostico, admin_creacaso, visa_global_appscript_backend [INFERRED 0.75]
- **DS-160 Data Collection & Report Flow** — intake_ds160_html, intake_ds160_handlepdfupload, intake_ds160_applyextracteddata, intake_ds160_submitform, intake_ds160_regenerarinforme [INFERRED 0.85]
- **localStorage/Cloud Sync Merge Pattern** — admin_synccases, intake_ds160_savetocloud, portal_savedraft, admin_localstorage_merge_bug [INFERRED 0.75]

## Communities (13 total, 0 thin omitted)

### Community 0 - "Community 0"
Cohesion: 0.09
Nodes (47): adminGetCases(), adminUpdateCase(), alertarRobertoDiagnostico(), analizarDs160Anteriores(), ANTHROPIC_KEY, BOT_SECRET, botHistory(), buildDiagnosticoPrompt() (+39 more)

### Community 1 - "Community 1"
Cohesion: 0.08
Nodes (31): checkPin, crearCaso, deleteCase, admin.html (CRM), loadLocal, localStorage merge on sync (critical bug fix), openDetail, renderCard (+23 more)

### Community 2 - "Community 2"
Cohesion: 0.15
Nodes (21): adminNewCase(), analizarDs160AnterioresAuto(), analizarPerfil(), analizarPerfilFamiliar(), analizarRechazo(), construirPerfil(), construirPromptFamiliar(), formatearCabecera() (+13 more)

### Community 3 - "Community 3"
Cohesion: 0.11
Nodes (18): enviarDiagnostico, diagnostico.html (Initial Diagnostic), redirectPayphone, renderReport, verificarRetornoPayphone, confetti, enviar, intake.html (New Client Intake Form) (+10 more)

### Community 4 - "Community 4"
Cohesion: 0.22
Nodes (14): alertarRobertoDiagnostico(), buildDiagnosticoPrompt(), callGeminiChat(), chatMessage(), enviarEmailDiagnostico(), guardarChat(), guardarDiagnostico(), guardarIntentoPayphone() (+6 more)

### Community 5 - "Community 5"
Cohesion: 0.17
Nodes (11): BIEN, CIVIL, CREDITO, FAMILIARES_INDOG, HIJOS, INGRESOS, LABORAL, MENSAJES (+3 more)

### Community 6 - "Community 6"
Cohesion: 0.22
Nodes (11): validateAndContinue before save ordering rule, Informes de expedientes README, intake-ds160.html (DS-160 Intake Form), regenerarInforme, save, saveToCloud, startAutoSave, submitForm (+3 more)

### Community 7 - "Community 7"
Cohesion: 0.29
Nodes (9): analizarConIA(), ANTHROPIC_KEY, COL_RESUMEN, doPost(), formatearCabecera(), getOrCreateSheet(), notificarRoberto(), ok() (+1 more)

### Community 8 - "Community 8"
Cohesion: 0.40
Nodes (8): addMsg(), clearQuickReplies(), greet(), hideTyping(), scrollBottom(), sendMessage(), setQuickReplies(), showTyping()

### Community 9 - "Community 9"
Cohesion: 0.36
Nodes (5): cancelarFollowups(), COLS, doPost(), getOrCreateFollowupSheet(), guardarFollowup()

### Community 10 - "Community 10"
Cohesion: 0.67
Nodes (4): applyExtractedData, extractPDFText, handlePDFUpload, Save PDF-extracted data to S.perTraveler immediately

### Community 11 - "Community 11"
Cohesion: 0.67
Nodes (3): handlePdfDrop, procesarPDFs, subirPDFs

## Ambiguous Edges - Review These
- `Guía Proceso Visa USA (PDF)` → `ds160.html (Interactive DS-160 Guide)`  [AMBIGUOUS]
  guia-proceso-visa-usa.pdf · relation: conceptually_related_to
- `Guía Proceso Visa USA (PDF)` → `simulador.html (Eligibility Simulator)`  [AMBIGUOUS]
  guia-proceso-visa-usa.pdf · relation: conceptually_related_to

## Knowledge Gaps
- **51 isolated node(s):** `PAYPHONE_TOKEN`, `PAYPHONE_STORE_ID`, `ANTHROPIC_KEY`, `SHEET_ID`, `COL_RESUMEN` (+46 more)
  These have ≤1 connection - possible missing edges or undocumented components.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **What is the exact relationship between `Guía Proceso Visa USA (PDF)` and `ds160.html (Interactive DS-160 Guide)`?**
  _Edge tagged AMBIGUOUS (relation: conceptually_related_to) - confidence is low._
- **What is the exact relationship between `Guía Proceso Visa USA (PDF)` and `simulador.html (Eligibility Simulator)`?**
  _Edge tagged AMBIGUOUS (relation: conceptually_related_to) - confidence is low._
- **Why does `visa-global-web Project Instructions` connect `Community 1` to `Community 3`?**
  _High betweenness centrality (0.061) - this node is a cross-community bridge._
- **Why does `ds160.html (Interactive DS-160 Guide)` connect `Community 1` to `Community 6`?**
  _High betweenness centrality (0.028) - this node is a cross-community bridge._
- **Are the 3 inferred relationships involving `admin.html (CRM)` (e.g. with `Google Apps Script Backend (appscript_*.js)` and `Gold-on-Dark Color Strategy (70/30 rule)`) actually correct?**
  _`admin.html (CRM)` has 3 INFERRED edges - model-reasoned connections that need verification._
- **What connects `PAYPHONE_TOKEN`, `PAYPHONE_STORE_ID`, `ANTHROPIC_KEY` to the rest of the system?**
  _51 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `Community 0` be split into smaller, more focused modules?**
  _Cohesion score 0.08776595744680851 - nodes in this community are weakly interconnected._