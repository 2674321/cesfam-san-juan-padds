// ════════════════════════════════════════════════════════
// 12_Backup.gs │ backups automáticos de la planilla a Google Drive
//
// Configurable desde el menú 🔧 Herramientas → 💾 Backups:
//   · ⚙️ Configurar backup automático  → días entre copias + hora + copias a
//     mantener (crea el disparador diario que revisa el calendario)
//   · 📦 Hacer backup ahora             → copia inmediata
// Las copias van a una carpeta "PADDS Backups" en Drive y las antiguas se
// eliminan automáticamente (se conservan solo las últimas N).

var _BACKUP_DEF = { dias: 7, keep: 10, hora: 6, carpeta: 'PADDS Backups' }

function _backupProps(p) {
  return {
    dias: Number(p.getProperty('BACKUP_DIAS')) || _BACKUP_DEF.dias,
    keep: Number(p.getProperty('BACKUP_KEEP')) || _BACKUP_DEF.keep,
    hora: Number(p.getProperty('BACKUP_HORA')) || _BACKUP_DEF.hora,
    ultimo: Number(p.getProperty('BACKUP_ULTIMO')) || 0,
    ssId: String(p.getProperty('BACKUP_SS_ID') || ''),
  }
}

// ⚙️ Configura el backup automático (días + hora + copias a mantener).
function configurarBackupAutomatico() {
  var ui = SpreadsheetApp.getUi()
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  var r1 = ui.prompt('💾 Backup automático',
    '¿Cada cuántos días hacer el respaldo?\n(ej: 7 = semanal, 1 = diario)',
    ui.ButtonSet.OK_CANCEL)
  if (r1.getSelectedButton() !== ui.Button.OK) return
  var dias = parseInt(r1.getResponseText(), 10)
  if (isNaN(dias) || dias < 1) dias = 7

  var r2 = ui.prompt('💾 Backup automático',
    '¿Hora del día en que se ejecuta? (0 a 23, ej: 6 = 06:00)',
    ui.ButtonSet.OK_CANCEL)
  if (r2.getSelectedButton() !== ui.Button.OK) return
  var hora = parseInt(r2.getResponseText(), 10)
  if (isNaN(hora) || hora < 0 || hora > 23) hora = _BACKUP_DEF.hora

  var r3 = ui.prompt('💾 Backup automático',
    '¿Cuántas copias mantener?\n(las más antiguas se eliminan)',
    ui.ButtonSet.OK_CANCEL)
  if (r3.getSelectedButton() !== ui.Button.OK) return
  var keep = parseInt(r3.getResponseText(), 10)
  if (isNaN(keep) || keep < 1) keep = _BACKUP_DEF.keep

  var p = PropertiesService.getScriptProperties()
  p.setProperty('BACKUP_DIAS', String(dias))
  p.setProperty('BACKUP_HORA', String(hora))
  p.setProperty('BACKUP_KEEP', String(keep))
  p.setProperty('BACKUP_SS_ID', ss.getId())

  var triggers = ScriptApp.getProjectTriggers()
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'verificarBackup') {
      ScriptApp.deleteTrigger(triggers[i])
    }
  }
  ScriptApp.newTrigger('verificarBackup').timeBased().everyDays(1)
    .atHour(hora).inTimezone(Session.getScriptTimeZone()).create()

  var resBk = null
  try { resBk = _ejecutarBackup(ss, false) } catch (eBk) {
    console.error('backup de prueba: ' + eBk.message)
  }

  ui.alert('💾 Backup automático',
    'Configurado:\n· Respaldo cada ' + dias + ' día(s) a las ' +
    ('0' + hora).slice(-2) + ':00\n· Se mantienen las últimas ' + keep + ' copias\n\n' +
    (resBk
      ? '✅ Copia de prueba creada:\n📄 ' + resBk.nombre + '\n🔗 ' + resBk.url
      : '⚠️ No se pudo hacer la copia de prueba. Revisa la consola.'),
    ui.ButtonSet.OK)
}

// Disparador diario: hace el backup solo si ya pasaron los días configurados.
function verificarBackup() {
  var p = PropertiesService.getScriptProperties()
  var st = _backupProps(p)
  if (!st.ssId) return
  if (Date.now() - st.ultimo >= st.dias * 86400000) {
    try {
      var ss = SpreadsheetApp.openById(st.ssId)
      _ejecutarBackup(ss, false)
    } catch (e) {
      console.error('verificarBackup: ' + e.message)
    }
  }
}

// 📦 Hacer backup ahora (desde el menú).
function ejecutarBackup() {
  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('📦 Hacer backup ahora',
    '¿Crear una copia de seguridad de la planilla en Google Drive?\n(se guarda en la carpeta "PADDS Backups")',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return
  _ejecutarBackup(SpreadsheetApp.getActiveSpreadsheet(), true)
}

function _ejecutarBackup(ss, conAlerta) {
  var p = PropertiesService.getScriptProperties()
  if (!p.getProperty('BACKUP_SS_ID')) p.setProperty('BACKUP_SS_ID', ss.getId())
  var st = _backupProps(p)

  var folders = DriveApp.getFoldersByName(_BACKUP_DEF.carpeta)
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(_BACKUP_DEF.carpeta)

  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm')
  var copiaId = ss.copy('PADDS Backup ' + fecha).getId()
  var copiaFile = DriveApp.getFileById(copiaId)
  folder.addFile(copiaFile)
  try { DriveApp.getRootFolder().removeFile(copiaFile) } catch (eRoot) {}

  var files = folder.getFiles()
  var lista = []
  while (files.hasNext()) {
    var f = files.next()
    if (f.getName().indexOf('PADDS Backup') === 0) {
      lista.push({ id: f.getId(), fecha: f.getDateCreated().getTime(), nombre: f.getName() })
    }
  }
  lista.sort(function(a, b) { return b.fecha - a.fecha })
  var eliminadas = 0
  for (var i = st.keep; i < lista.length; i++) {
    try {
      DriveApp.getFileById(lista[i].id).setTrashed(true)
      eliminadas++
    } catch (eTrash) {}
  }

  p.setProperty('BACKUP_ULTIMO', String(Date.now()))
  var url = copiaFile.getUrl()

  if (conAlerta) {
    try { ss.toast('Backup creado: ' + copiaFile.getName() + ' 🔗 ' + url, 'PADDS', 6) } catch (eT) {}
    SpreadsheetApp.getUi().alert('📦 Backup',
      '✅ Backup creado:\n📄 ' + copiaFile.getName() +
      '\n🔗 ' + url +
      '\n📁 Carpeta: ' + _BACKUP_DEF.carpeta +
      '\n\nCopias guardadas: ' + lista.length +
      (eliminadas ? '\nEliminadas (antiguas): ' + eliminadas : ''),
      SpreadsheetApp.getUi().ButtonSet.OK)
  } else {
    try { ss.toast('Backup automático creado: ' + copiaFile.getName() + ' 🔗 ' + url, 'PADDS', 6) } catch (eT2) {}
  }
  return { nombre: copiaFile.getName(), url: url, total: lista.length, eliminadas: eliminadas }
}
