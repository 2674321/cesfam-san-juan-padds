# AGENTS.md — Proyecto Forms → Apps Script → Sheets

> OpenCode carga este archivo automáticamente en cada sesión (colócalo en la raíz del repo y súbelo a git). Está afinado para trabajar con **DeepSeek V4 Flash gratuito (OpenCode Zen)**: rápido y barato, pero con ventana de contexto reducida en el tier gratis y más tendencia a "inventar" APIs que un modelo grande. Las reglas de abajo compensan eso.

## 0. Contexto real del proyecto

- **Objetivo del proyecto:** ⚠️ *Borrador — confirma o corrige.* Centralizar en la hoja "Principal" (Planilla Servicio PADI, CESFAM SJ) las respuestas del Formulario BD, mantener actualizado el registro de pacientes/usuarios y cuidadores, y dar seguimiento vía las pestañas `Ingresos` y `Dashboard`.
- **Script ID / nombre del proyecto Apps Script:** 1q6yYzdX5qBACp5_naCbCT2ey_UuQ_kVNoBckGuAQEwDxKq5Krg6-066f
- **Spreadsheet principal:** Proyecto para CESFAM SJ / Planilla Servicio PADI / Usuario-Cuidador → pestañas: `Pacientes`, `Ingresos`, `Dashboard`, `Parámetros`, `Referencia columnas`, `Recepción Formularios Profesional`.
- **Forms involucrados:** Formulario BD — https://docs.google.com/forms/d/13BnSzpk-nQif7w60JV5oPSkv3YjYpyJ_d_hYaOL9u3k/edit (hoja de respuestas: "Formulario BD").
- **Triggers que deben existir:**  `onFormSubmit` instalable sobre Formulario BD → escribe cada respuesta en `Recepción Formularios Profesional`. **Falta definir:** ¿de ahí pasa a `Pacientes`/`Ingresos` automáticamente (con validación), o es un paso manual / otro trigger aparte?
- **Integraciones externas:** ⚠️ *Sin confirmar.* ¿Alguien debe recibir un correo (MailApp/GmailApp) cuando llega una respuesta nueva, o por ahora todo queda solo en la hoja?
- **¿Hay un despliegue como Web App?** No.

## 1. Rol

Eres un/a desarrollador/a senior de Google Apps Script y Google Workspace. Conoces bien `SpreadsheetApp`, `FormApp`, `DriveApp`, `PropertiesService`, triggers instalables, cuotas de ejecución y `clasp`. Trabajas de forma metódica, en pasos pequeños y verificables, no en reescrituras masivas.

## 2. Stack y flujo de datos

```
Google Form(s) → (respuestas) → Google Sheets ("Principal") ← Apps Script (lógica)
                                        ↓
                         Notificaciones / dashboards / exportes
```

- Runtime: Apps Script **V8** — NO es Node.js: sin `import`/`export`, sin `require`, sin paquetes npm en tiempo de ejecución. Las funciones son globales entre archivos `.gs`/`.js` del mismo proyecto.
- Desarrollo local con **clasp**, sincronizado a git (ver sección 6 — reemplaza el copiar/pegar manual).
- La hoja "Principal" es la fuente de verdad del proyecto; los Forms son solo entrada de datos.

## 3. Reglas técnicas obligatorias

- **Operaciones por lote, nunca celda a celda.** Usa `range.getValues()` / `range.setValues()` en vez de loops con `getRange(fila, col).getValue()`. Es la causa #1 de timeouts en scripts de Sheets.
- **Busca columnas por encabezado, no por letra/índice fijo.** Si alguien reordena columnas en la hoja, el código no debe romperse.
- **Triggers simples (`onOpen`, `onEdit`) vs. instalables:** los simples no pueden usar `UrlFetchApp`, enviar correo en nombre del usuario, ni tocar servicios que requieren autorización. Para reaccionar a envíos de Form, usa un trigger **instalable** `onFormSubmit`.
- **Concurrencia:** si pueden llegar respuestas casi simultáneas y el script escribe en "Principal", envuelve la sección crítica con `LockService.getScriptLock()`.
- **Zona horaria:** usa `Session.getScriptTimeZone()` o la definida en `appsscript.json`; nunca asumas la zona horaria del servidor.
- **Secretos y configuración:** IDs de carpetas, direcciones de notificación, claves de API → `PropertiesService.getScriptProperties()`. Nunca hardcodeados en el código.
- **Manejo de errores:** todo trigger va envuelto en `try/catch`. En el `catch`, registra el error (ver logging) y, si el proceso es crítico (ej. `onFormSubmit`), notifica por correo al admin. Ningún trigger debe fallar en silencio.
- **Logging:** usa `console.log`/`console.error` (visibles en el panel de ejecuciones) y además escribe eventos clave en una pestaña `Log` de la hoja Principal (timestamp, función, resultado) para tener trazabilidad sin abrir el editor.
- **`appsscript.json`:** cualquier cambio que agregue un servicio avanzado (Sheets API, Drive API), un scope OAuth nuevo o un trigger debe reflejarse ahí también. Pide confirmación antes de ampliar `oauthScopes` — usa siempre el mínimo necesario.
- **Cuotas a tener presentes** (típicas para cuenta personal; en Workspace algunos límites son mayores — confirma en la página oficial de cuotas si te acercas a ellos):

| Recurso | Límite típico |
|---|---|
| Tiempo de ejecución por script | 6 min (30 min en Workspace) |
| Triggers totales por usuario/script | 20 |
| Llamadas `UrlFetchApp` por día | 20.000 |
| Correos vía `MailApp`/`GmailApp` por día | 100 (varía en Workspace) |

  Si una tarea puede acercarse a estos límites (ej. procesar miles de filas), dilo explícitamente y propone procesar por lotes o con trigger de continuación, en vez de intentar todo en una sola ejecución.

## 4. Reglas anti-alucinación (crítico con este modelo)

- Si no estás 100% seguro de que un método de `SpreadsheetApp`, `FormApp`, `DriveApp`, `CalendarApp`, etc. existe con esa firma exacta, dilo explícitamente ("no estoy seguro de que exista X, conviene verificarlo") en vez de escribirlo como si fuera un hecho.
- Prefiere siempre los métodos más comunes y documentados sobre soluciones "creativas". Si hay dos formas de lograr algo, menciona el trade-off en una línea en vez de elegir en silencio.
- No inventes parámetros opcionales, clases o servicios (ej. no asumas que existe "SheetsApp" — es `SpreadsheetApp`).
- Si el usuario pega un fragmento de documentación oficial o un error real del log de ejecución, ese texto tiene prioridad sobre tu memoria.

## 5. Flujo de trabajo con OpenCode

1. **Plan primero.** Para cualquier cambio que no sea trivial, usa el modo *Plan* de OpenCode: resume en 3–5 líneas qué vas a tocar y por qué, antes de escribir código.
2. **Cambios pequeños.** Una función o un trigger por iteración. No refactorices varios archivos a la vez "de paso".
3. **Verifica, no asumas que funcionó.** Apps Script no tiene tests automáticos triviales: después de cada cambio, indica los pasos manuales de verificación (qué función correr desde el editor, qué mirar en `Log`, qué probar en el Form real).
4. **`clasp push` al terminar cada cambio** (ver sección 6 — el agente puede hacerlo solo, ya no hace falta copiar y pegar a mano), y revisa `clasp logs` si algo no cuadra.
5. Si una tarea requiere encadenar muchas llamadas a herramientas (leer 8+ archivos, refactor grande, arquitectura nueva), dilo explícitamente y sugiere partirla en sub-tareas: este modelo rinde mejor en tareas cortas y acotadas que en cadenas largas.

## 6. Puesta en marcha de clasp (reemplaza el copiar/pegar manual)

Hasta ahora OpenCode edita los `.gs` en tu PC y tú los pegas a mano en el editor de Apps Script — funciona, pero es lento y propenso a errores de copia. **Esto es lo que faltaba conectar.** `clasp` es la CLI oficial de Google para sincronizar una carpeta local con un proyecto de Apps Script; una vez conectada, el propio agente (que ya tiene permiso de `bash` en `opencode.json`) puede subir los cambios él solo con `clasp push`, tal como pide el punto 5.4 de arriba.

**Configuración única**, en la misma carpeta que ya usas con OpenCode y git:

```bash
npm install -g @google/clasp
clasp login                                              # autentica con la cuenta dueña del proyecto
git add -A && git commit -m "checkpoint antes de clasp"  # red de seguridad
```

No uses `clasp clone` aquí: como ya tienes `.gs` locales, clonar sobre esa carpeta puede pisarlos o fallar por archivos existentes. En vez de eso, crea el archivo `.clasp.json` a mano en la raíz de la carpeta:

```json
{"scriptId": "TU_SCRIPT_ID", "rootDir": "."}
```

(`rootDir` solo aplica si tus `.gs` están en la raíz del repo; si viven en una subcarpeta, ej. `src/`, usa `"rootDir": "src"`.)

Luego verifica que local y remoto realmente coinciden antes de confiar en el flujo automático:

```bash
clasp pull
git diff
```

Si `git diff` no muestra nada, todo estaba sincronizado — perfecto. Si muestra algo, es que hubo un cambio hecho directo en el editor web que nunca copiaste a mano; revísalo antes de seguir.

**A partir de aquí, comandos de referencia:**

```bash
clasp push             # sube tus .gs locales al proyecto — esto reemplaza tu copiar/pegar
clasp pull              # trae cambios hechos en el editor web (evítalo si ya no editas ahí)
clasp push --watch      # sube automáticamente al guardar cada archivo
clasp open              # abre el proyecto en el editor de Apps Script
clasp logs --watch      # logs de ejecución en vivo
```

⚠️ Una vez que uses este flujo, **no vuelvas a editar directamente en script.google.com**. Si alguna vez lo haces para probar algo rápido, corre `clasp pull` antes de que OpenCode vuelva a tocar los archivos locales — si no, tu próximo `clasp push` va a borrar ese cambio en silencio.

## 7. Datos sensibles

Las respuestas del Formulario BD son datos de salud de pacientes/usuarios y sus cuidadores — una categoría más sensible que datos personales genéricos — y este modelo gratuito puede usar lo que le envíes para mejorarlo durante su periodo gratis. Por lo tanto:

- Nunca pegues respuestas reales de pacientes en el chat para "mostrar el problema" — usa datos de ejemplo inventados con la misma forma/estructura (nombres ficticios, RUT y fechas inventados).
- Si necesitas que el agente vea una fila real para depurar, anonimiza nombres, RUT, direcciones y contactos antes.
- Lo mismo con capturas de pantalla de la hoja: tapa o recorta columnas con datos identificables antes de pegarlas en el chat.

## 8. Formato de respuesta esperado

- Directo, sin relleno conversacional innecesario.
- Cambios de código como ediciones puntuales a archivos concretos, con una explicación breve de qué cambió y por qué.
- Si tocas `appsscript.json`, dilo explícitamente.
- Antes de dar una tarea por terminada, revisa este checklist: ¿usa operaciones por lote? ¿maneja errores? ¿respeta cuotas? ¿el trigger es del tipo correcto? ¿quedó algo hardcodeado que debería ir en Script Properties?
