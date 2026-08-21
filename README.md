# CESFAM San Juan — Módulo de Urgencia (línea V4)

Sistema de gestión de **pacientes de urgencia para el CESFAM San Juan**, construido sobre
**Google Apps Script** con Google Sheets como capa de datos. Esta rama del proyecto corresponde
a la línea **V4**: fichas PADI, agenda médica, alertas y módulo de recepción.

> 📌 Este repositorio es la **versión activa** del sistema.
> El histórico PADI original: [proyecto-cesfam-san-juan](https://github.com/2674321/proyecto-cesfam-san-juan) ·
> Snapshot pre-V4 archivado: [cesfam-san-juan-v3-backup](https://github.com/2674321/cesfam-san-juan-v3-backup)

## Funcionalidades

- 🧑‍⚕️ **Fichas de pacientes** (formato PADI) con historial por paciente
- 📅 **Agenda** de atenciones
- 🔔 **Alertas** clínicas y administrativas
- 🚪 **Recepción**: ingreso y triaje de pacientes
- 💾 **Backup** integrado de datos
- 🎨 Sistema de diseño unificado (`02_DesignSystem.js`)

## Estructura

| Archivo | Módulo |
|---------|--------|
| `00_Constantes.js` | Constantes globales |
| `01_Utilidades.js` | Utilidades compartidas |
| `02_DesignSystem.js` | Sistema de diseño (UI) |
| `03_Pacientes.js` | Fichas PADI de pacientes |
| `04_Eventos.js` | Eventos / atenciones |
| `05_Alertas.js` | Alertas |
| `06_Formato.js` | Formato de planillas |
| `07_Agenda.js` | Agenda médica |
| `08_Recepcion.js` | Recepción e ingresos |
| `09_Config.js` | Configuración |
| `10_Pruebas.js` | Pruebas |
| `11_Ingresos.js` | Ingresos |
| `12_Backup.js` | Respaldo de datos |
| `13_V4.js` | Núcleo de la línea V4 |

## Desarrollo

El proyecto se sincroniza con Apps Script mediante [clasp](https://github.com/google/clasp):

```bash
clasp login          # primera vez
clasp push           # subir cambios locales → Apps Script
clasp pull           # bajar cambios desde Apps Script
```

## Estado

- ✅ En producción — versión activa (línea V4)
- Etiqueta actual: `v4.0`

---

Desarrollado por [@2674321](https://github.com/2674321) · Coquimbo, Chile
