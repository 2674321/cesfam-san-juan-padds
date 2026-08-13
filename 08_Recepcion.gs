
// ════════════════════════════════════════════════════════
// ÍNDICE ▏08_Recepcion.gs │ formulario profesional, aprobación de formularios, estadísticas

// ─────────────────────────────────────────────────────────

function _oscurecer(hex, amount) {
  return _ajustarHex(hex, -amount, -amount, -amount)
}

var _ESTADO_CSS = {
  'Pendiente': { badge: '#E65100', tint: '#FFF3E0', fg: '#FFFFFF' },
  'Aprobado':  { badge: '#2E7D32', tint: '#E8F5E9', fg: '#FFFFFF' },
  'Rechazado': { badge: '#C62828', tint: '#FFEBEE', fg: '#FFFFFF' },
}

var _FORM_SEC_TINTS = {
  1:'#F5F5F5', 2:'#FFF8F0', 3:'#F5F5F5', 4:'#FFF8F0', 5:'#F5F5F5',
  6:'#E8EAF6', 7:'#F0F4FF', 8:'#E8EAF6', 9:'#F0F4FF',
  10:'#E8F5E9',11:'#F0F9F6',12:'#E8F5E9',13:'#F0F9F6',14:'#E8F5E9',15:'#F0F9F6',16:'#E8F5E9',
  17:'#F3E5F5',
}

var _ALIGN = {
  1: 'center', 2: 'center', 3: 'center', 4: 'left', 5: 'center',
  6: 'center', 7: 'center', 8: 'left', 9: 'left', 10: 'left', 11: 'left',
  12: 'left', 13: 'center', 14: 'left', 15: 'left', 16: 'center', 17: 'left',
}

function _estiloFila(sh, row, lc) {
  var even = row % 2 === 0
  var rng = sh.getRange(row, 1, 1, lc)
  rng
    .setFontFamily('Calibri').setFontSize(10).setFontColor('#212121')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID)
  var bgs = []
  for (var c = 1; c <= lc; c++)
    bgs.push(even ? _oscurecer(_FORM_SEC_TINTS[c] || '#ffffff', 8) : (_FORM_SEC_TINTS[c] || '#ffffff'))
  rng.setBackgrounds([bgs])
  for (var c = 1; c <= lc; c++)
    sh.getRange(row, c).setHorizontalAlignment(_ALIGN[c] || 'center')
  sh.setRowHeight(row, 26)
  sh.getRange(row, 1).setNumberFormat('dd/mm/yyyy hh:mm')
  sh.getRange(row, 5).setNumberFormat('dd/mm/yyyy hh:mm')
  sh.getRange(row, 6).setNumberFormat('dd/mm/yyyy')
  sh.getRange(row, 16).setNumberFormat('dd/mm/yyyy')
  sh.getRange(row, 17).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
}

function _colorEstado(sh, row, estado, lc) {
  var css = _ESTADO_CSS[estado]
  if (!css) return
  sh.getRange(row, 3).setBackground(css.badge).setFontColor(css.fg).setFontWeight('bold')
}

function _miniResumenForm(sh) {
  var lc = sh.getLastColumn() || 17
  var SEP = getFormulaSep()

  var f2 = '="⏳ Pendientes: "&COUNTIF($C$5:$C' + SEP + '"Pendiente")' +
    '&"   ✅ Aprobados: "&COUNTIF($C$5:$C' + SEP + '"Aprobado")' +
    '&"   ❌ Rechazados: "&COUNTIF($C$5:$C' + SEP + '"Rechazado")' +
    '&"   📥 Total formularios: "&COUNTA($C$5:$C)'
  try { sh.getRange(2, 1, 1, lc).breakApart() } catch(eB2) {}
  sh.getRange(2, 1, 1, lc).merge()
  sh.getRange(2, 1).setFormula(f2)
    .setFontFamily('Calibri').setFontSize(11).setFontWeight('bold')
    .setFontColor('#0D47A1').setBackground('#E3F2FD')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#1565C0', SpreadsheetApp.BorderStyle.SOLID)
  sh.setRowHeight(2, 28)
}

function _obtenerEstadisticas() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_FORM)
  if (!sh) return null
  var lr = sh.getLastRow()
  if (lr < 5) return { total: 0, pend: 0, aprob: 0, rech: 0 }
  var data = sh.getRange(5, 1, lr - 4, 3).getValues()
  var stats = { total: 0, pend: 0, aprob: 0, rech: 0 }
  for (var r = 0; r < data.length; r++) {
    var est = String(data[r][2] || '').trim()
    if (!est) continue
    stats.total++
    if (est === 'Pendiente') stats.pend++
    else if (est === 'Aprobado') stats.aprob++
    else if (est === 'Rechazado') stats.rech++
  }
  return stats
}

function mostrarSidebarEstadisticas() {
  var stats = _obtenerEstadisticas()
  if (!stats) { SpreadsheetApp.getUi().alert('Primero crea la hoja de recepción (📝 Formulario → "📋 Crear hoja de recepción").'); return }
  var html = '<html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:0;padding:16px;color:#202124;font-size:14px;background:#fff}' +
    'h2{margin:0 0 4px 0;font-size:16px;font-weight:500;color:#1a237e}' +
    '.sub{font-size:12px;color:#5f6368;margin-bottom:16px}' +
    '.card{border:1px solid #e0e0e0;border-radius:8px;padding:12px;margin-bottom:12px}' +
    '.total-num{font-size:36px;font-weight:300;color:#1a237e}' +
    '.total-label{font-size:12px;color:#5f6368}' +
    '.row{display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #f5f5f5}' +
    '.row:last-child{border-bottom:none}' +
    '.count{font-size:18px;font-weight:500}' +
    '.tag{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:8px}' +
    '.tag-pend{background:#E65100} .tag-aprob{background:#2E7D32} .tag-rech{background:#C62828}' +
    '.label-name{color:#3c4043;display:flex;align-items:center}' +
    '.footer{font-size:11px;color:#9aa0a6;text-align:center;padding-top:12px;border-top:1px solid #e8eaed;margin-top:12px}' +
    '</style></head><body>' +
    '<h2>Recepción formularios PADDS</h2>' +
    '<div class="sub" id="updated">Actualizado: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss') + '</div>' +
    '<div class="card"><div class="total-num" id="totalNum">' + stats.total + '</div><div class="total-label">Formularios recibidos</div></div>' +
    '<div class="card"><div style="font-size:12px;color:#5f6368;margin-bottom:8px">Estado</div>' +
    '<div class="row"><span class="label-name"><span class="tag tag-pend"></span>Pendiente</span><span class="count" id="pendCount">' + stats.pend + '</span></div>' +
    '<div class="row"><span class="label-name"><span class="tag tag-aprob"></span>Aprobado</span><span class="count" id="aprobCount">' + stats.aprob + '</span></div>' +
    '<div class="row"><span class="label-name"><span class="tag tag-rech"></span>Rechazado</span><span class="count" id="rechCount">' + stats.rech + '</span></div></div>' +
    '<div class="footer" id="footer">Los datos se actualizan automáticamente</div>' +
    '<script>' +
    'function refreshStats(){google.script.run.withSuccessHandler(function(s){' +
    'document.getElementById("totalNum").textContent=s.total||0;' +
    'document.getElementById("pendCount").textContent=s.pend||0;' +
    'document.getElementById("aprobCount").textContent=s.aprob||0;' +
    'document.getElementById("rechCount").textContent=s.rech||0;' +
    'var d=new Date();var h=d.getHours().toString().padStart(2,"0");var m=d.getMinutes().toString().padStart(2,"0");var s2=d.getSeconds().toString().padStart(2,"0");' +
    'document.getElementById("updated").textContent="Actualizado: "+h+":"+m+":"+s2;' +
    '})._obtenerEstadisticas();}' +
    'setInterval(refreshStats,5000);' +
    '</script></body></html>'
  var output = HtmlService.createHtmlOutput(html).setTitle('Recepción Formulario').setWidth(300)
  SpreadsheetApp.getUi().showSidebar(output)
}

function _configurarTriggerFormulario() {
  var triggers = ScriptApp.getProjectTriggers()
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(triggers[i])
    }
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create()
}

function configurarTriggerFormulario() {
  try {
    _configurarTriggerFormulario()
    SpreadsheetApp.getUi().alert('Recepción automática activada. Los formularios se recibirán solos al ser enviados.')
  } catch(e) {
    SpreadsheetApp.getUi().alert('Error al configurar: ' + e.message + '\nAsegúrate de tener permisos y que el formulario este vinculado a este spreadsheet.')
  }
}

function crearFormularioBase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var legacy = ss.getSheetByName('Formulario Usuario / Profesional')
  if (legacy) legacy.setName(HOJA_FORM)
  var sh = ss.getSheetByName(HOJA_FORM)
  var existe = !!sh
  if (!existe) {
    ss.toast('Creando hoja de recepción…', 'PADDS', 1)
    sh = ss.insertSheet(HOJA_FORM)
  } else {
    ss.toast('Aplicando mejoras sin borrar datos…', 'PADDS', 1)
  }
  sh.setTabColor('#9B8CA8')

  sh.setFrozenColumns(0)
  sh.setFrozenRows(0)

  var lc = 17

  var WHITE = '#ffffff', NAVY = '#0D47A1', BLUE = '#1565C0'
  var LIGHT_BLUE = '#E3F2FD', WARN = '#FFF8E1', BORDER = '#BBDEFB'
  var BS = SpreadsheetApp.BorderStyle.SOLID
  var BS_MED = SpreadsheetApp.BorderStyle.SOLID_MEDIUM

  sh.setRowHeight(1, 52)
  sh.getRange(1, 1, 1, lc).merge()
  sh.getRange(1, 1)
    .setValue('📋  Recepcion formularios PADDS')
    .setFontFamily('Calibri').setFontSize(20).setFontWeight('bold')
    .setFontColor(WHITE).setBackground(NAVY)
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, WHITE, BS_MED)

  sh.setRowHeight(2, 28)
  sh.getRange(2, 1).setNote(
    'Los formularios llegan automáticamente. Revise los datos, corrija si es necesario, ' +
    'luego use el menú 📝 Formulario: "✅ Aprobar y pasar a Pacientes" (se copia a Pacientes ' +
    'y la fila se elimina) o "❌ Rechazar solicitudes pendientes".')

  sh.setRowHeight(3, 34)
  var SEC_COLORS = ['#1A237E', '#1565C0', '#2E7D5B', '#6A1B9A']
  for (var s = 0; s < FORM_SECCIONES.length; s++) {
    var sec = FORM_SECCIONES[s]
    var fin = Math.min(sec.fin, lc)
    sh.getRange(3, sec.ini, 1, fin - sec.ini + 1).merge()
      .setBackground(SEC_COLORS[s]).setFontColor(WHITE)
      .setFontFamily('Calibri').setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, WHITE, BS_MED)
  }

  var FORM_TOOLTIPS = [
    'Fecha y hora de recepción del formulario (se llena automáticamente).',
    'Tipo de formulario recibido (ej: "Registro Atención Diaria").',
    'Estado de revisión: Pendiente → Aprobado → Rechazado. Cambie el valor usando el menú.',
    'Correo del profesional que aprobó el formulario (se llena automáticamente).',
    'Fecha y hora de aprobación (se llena automáticamente).',
    'Fecha de la atención registrada. Se usa para actualizar la columna de servicio en Pacientes.',
    'RUN del usuario (sin puntos, con guion). Usado para buscar al paciente en Pacientes.',
    'Nombre del usuario.',
    'Apellido del usuario.',
    'Profesional que realizó la prestación.',
    'Prestación o servicio realizado. Determina a qué columna de Pacientes se copia la fecha.',
    'Clasificación del estado nutricional (solo aplica cuando la prestación es de Nutrióloga).',
    'Indica si la visita fue perdida (SI/NO). Si es SI, se registra en Observaciones.',
    'Resultado del índice Barthel (dependencia funcional).',
    'Resultado del test Zarit (sobrecarga del cuidador).',
    'Fecha sugerida para el próximo control.',
    'Observaciones de la atención. Se acumulan en Pacientes en formato [fecha] con los detalles.',
  ]

  sh.setRowHeight(4, 30)
  var HEADERS = [
    'FECHA RECEPCIÓN', 'FORMULARIO', 'ESTADO', 'APROBADO POR', 'FECHA APROB.',
    'FECHA ATENCIÓN', 'RUT USUARIO', 'NOMBRE', 'APELLIDO', 'PROFESIONAL',
    'PRESTACIÓN', 'ESTADO NUTRICIONAL', 'VISITA PERDIDA', 'RESULTADO BARTHEL',
    'RESULTADO ZARIT', 'PRÓXIMO CONTROL', 'OBSERVACIONES',
  ]
  for (var c = 0; c < lc; c++) {
    var secIdx = -1
    for (var si = 0; si < FORM_SECCIONES.length; si++) {
      if (c + 1 >= FORM_SECCIONES[si].ini && c + 1 <= FORM_SECCIONES[si].fin) { secIdx = si; break }
    }
    var hdrBg = secIdx >= 0 ? _lightenHex(SEC_COLORS[secIdx], 60, 60, 60) : '#ECEFF1'
    sh.getRange(4, c + 1)
      .setValue(HEADERS[c])
      .setBackground(hdrBg).setFontColor(WHITE)
      .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, WHITE, BS_MED)
      .setNote([
        HEADERS[c],
        FORM_TOOLTIPS[c] || ''
      ].join('\n'))
  }

  var NEW_ANCHOS = [150, 120, 100, 180, 140, 110, 100, 170, 150, 160, 180, 150, 90, 140, 130, 110, 280]
  for (var i = 0; i < Math.min(NEW_ANCHOS.length, lc); i++)
    sh.setColumnWidth(i + 1, NEW_ANCHOS[i])

  _miniResumenForm(sh)

  var lrForm = sh.getLastRow()
  var maxRows = Math.max(300, lrForm + 5)
  var statusCol = sh.getRange(5, 3, maxRows - 4, 1)
  var rules = []
  var estados = ['Rechazado', 'Aprobado', 'Pendiente']
  for (var ei = 0; ei < estados.length; ei++) {
    var css = _ESTADO_CSS[estados[ei]]
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(estados[ei])
      .setBackground(css.badge).setFontColor(css.fg).setBold(true)
      .setRanges([statusCol]).build())
  }
  // Reglas para filas completas según ESTADO: tinte de fondo + lectura rápida
  var dataRng = sh.getRange(5, 1, maxRows - 4, lc)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C5="Pendiente"')
    .setBackground('#FFF3E0')
    .setRanges([dataRng]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C5="Aprobado"')
    .setBackground('#E8F5E9').setFontColor('#888888')
    .setRanges([dataRng]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C5="Rechazado"')
    .setBackground('#FAFAFA').setFontColor('#aaaaaa').setStrikethrough(true)
    .setRanges([dataRng]).build())

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($G5<>"",COUNTIF($G$5:$G,$G5)>1)')
    .setBackground('#FFF8E1').setFontColor('#E65100').setBold(true)
    .setRanges([sh.getRange(5, 7, maxRows - 4, 1)]).build())
  sh.setConditionalFormatRules(rules)

  var hasta = existe ? Math.min(lrForm, 500) : 25
  for (var fr = 5; fr <= Math.max(hasta, 25); fr++) _estiloFila(sh, fr, lc)

  sh.getRange(5, 3, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(FORM_ESTADOS, true).setAllowInvalid(true).build())

  // API de Sheets en reglas de validación (bug conocido), y el RUN se normaliza y

  sh.getRange(4, 7).setNote('Formato RUN: 12345678-5 (sin puntos, con guión). Se normaliza y valida automáticamente al procesar la fila.')

  sh.getRange(5, 13, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['SI', 'NO'], true)
      .setAllowInvalid(true).setHelpText('SI = visita perdida (se registra en Observaciones)').build())

  sh.getRange(5, 16, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build())
  sh.getRange(5, 1, maxRows - 4, 1).setNumberFormat('dd/mm/yyyy hh:mm')
  sh.getRange(5, 5, maxRows - 4, 1).setNumberFormat('dd/mm/yyyy hh:mm')
  sh.getRange(5, 6, maxRows - 4, 1).setNumberFormat('dd/mm/yyyy')
  sh.getRange(5, 16, maxRows - 4, 1).setNumberFormat('dd/mm/yyyy')
  sh.getRange(5, 17, maxRows - 4, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
  sh.setRowHeights(5, maxRows - 4, 26)

  var _prestList = _PRESTACIONES_LIST.map(function(n) {
    return n.toLowerCase().replace(/\b\w/g, function(ch) { return ch.toUpperCase() })
  }).sort()
  sh.getRange(5, 11, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(_prestList, true)
      .setAllowInvalid(true).setHelpText('Elija la prestación realizada (se copia a la columna correspondiente de Pacientes)').build())

  sh.getRange(5, 12, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['NORMAL', 'SOBREPESO', 'OBESIDAD', 'BAJO PESO', 'N/A'], true)
      .setAllowInvalid(true).build())

  sh.setFrozenColumns(0)
  _unmergeQueCruzaFila(sh, 4)
  try { sh.setFrozenRows(4) } catch(eFz) {}
  var _frmF = sh.getFilter()
  if (_frmF) _frmF.remove()
  sh.getRange(4, 1, 1, lc).createFilter()

  try { ss.setSpreadsheetTimeZone('America/Santiago') } catch(e) {}

  try { _configurarTriggerFormulario() } catch(e) {}

  // Recortar solo filas completamente vacías al final (nunca se tocan datos)
  var _fmr = sh.getMaxRows()
  var _minimo = Math.max(lrForm + 10, maxRows + 5)
  if (_fmr > _minimo) {
    var _sobra = sh.getRange(_minimo + 1, 1, _fmr - _minimo, lc).getValues()
    var _todoVacio = true
    for (var _v = 0; _v < _sobra.length && _todoVacio; _v++) {
      for (var _vc = 0; _vc < lc; _vc++) {
        if (_sobra[_v][_vc] !== '' && _sobra[_v][_vc] !== null && _sobra[_v][_vc] !== undefined) { _todoVacio = false; break }
      }
    }
    if (_todoVacio) sh.deleteRows(_minimo + 1, _fmr - _minimo)
  }

  try { configurarResaltadoFila() } catch(e) {}

  try { sh.setFrozenRows(4) } catch(eFz) {}

  ss.toast(existe
    ? 'Hoja de recepción actualizada: formato, validaciones y resumen aplicados sin borrar ningún formulario'
    : 'Hoja de recepción creada con 17 columnas (recepción automática activada)', 'Recepción', 4)
}

function onFormSubmit(e) {
  try {
    if (!e || !e.namedValues) return
    var ss = e.source
    if (!ss) return
    var stagingSh = ss.getSheetByName(HOJA_FORM)
    if (!stagingSh) return

    var lc = 17
    var newRow = []
    for (var c = 0; c < lc; c++) newRow.push('')

    newRow[0] = e.namedValues['Marca temporal'] ? e.namedValues['Marca temporal'][0] : new Date()
    newRow[2] = 'Pendiente'
    newRow[1] = 'Registro Atención Diaria'

    var trimmedKeys = {}
    for (var _ef in e.namedValues) {
      if (e.namedValues.hasOwnProperty(_ef)) trimmedKeys[_ef.trim()] = e.namedValues[_ef]
    }

    for (var field in _FORM_FIELD_MAP) {
      if (_FORM_FIELD_MAP.hasOwnProperty(field) && trimmedKeys[field] && trimmedKeys[field][0]) {
        newRow[_FORM_FIELD_MAP[field] - 1] = String(trimmedKeys[field][0]).trim()
      }
    }

    var lr = stagingSh.getLastRow()
    var nr = Math.max(lr + 1, 5)
    if (nr > stagingSh.getMaxRows()) stagingSh.insertRowsAfter(stagingSh.getMaxRows(), nr - stagingSh.getMaxRows() + 10)
    stagingSh.getRange(nr, 1, 1, lc).setValues([newRow])
    _estiloFila(stagingSh, nr, lc)
    _colorEstado(stagingSh, nr, 'Pendiente', lc)
    _miniResumenForm(stagingSh)
  } catch(err) {
    console.log('onFormSubmit: ' + err.message)
  }
}

function aprobarFormularios() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Procesando formularios…', 'PADDS', 1)
  var sh = ss.getSheetByName(HOJA_FORM)
  if (!sh) { SpreadsheetApp.getUi().alert('Primero crea la hoja de recepción (menú Formulario).'); return }

  var lr = sh.getLastRow()
  var lc = sh.getLastColumn()
  if (lr < 5) { SpreadsheetApp.getUi().alert('No hay formularios para procesar.'); return }

  var lock = LockService.getScriptLock()
  try { if (!lock.tryLock(10000)) { ss.toast('No se pudo adquirir el bloqueo, intenta de nuevo', 'Formulario', 3); return } } catch(ee) { return }

  try {
    var data = sh.getRange(5, 1, lr - 4, lc).getValues()
    var email = ''
    try { email = Session.getActiveUser().getEmail() } catch(e) {}
    var now = new Date()

    var aprobRows = []

    var rechRows = []

    for (var r = data.length - 1; r >= 0; r--) {
      var rowNum = r + 5
      var estado = String(data[r][2] || '').trim()
      if (estado === 'Rechazado') {
        rechRows.push(rowNum)
      } else if (estado === 'Aprobado') {
        if (!data[r][6] || String(data[r][6]).trim() === '') continue
        if (!data[r][7] || String(data[r][7]).trim() === '') continue
        aprobRows.push({ row: rowNum, data: data[r] })
      }

    }

    // (las filas no se eliminan aún: primero se transfieren a Pacientes)
    for (var i = 0; i < aprobRows.length; i++) {
      var rn = aprobRows[i].row
      var curAprob = String(data[rn - 6][3] || '').trim()
      var curFecha = data[rn - 6][4]
      if (!curAprob) sh.getRange(rn, 4).setValue(email)
      if (!curFecha) {
        sh.getRange(rn, 5).setValue(now).setNumberFormat('dd/mm/yyyy hh:mm')
      }
    }

    var formRows = aprobRows.map(function(a) { return a.data })
    var transferidos = formRows.length ? _batchCopiarFormularios(formRows) : 0
    if (transferidos > 0) recalcularTodo()

    // Eliminar SOLO filas (nunca hojas): Aprobados (ya copiados a Pacientes)

    var aEliminar = rechRows.slice()
    for (var i = 0; i < aprobRows.length; i++) aEliminar.push(aprobRows[i].row)
    aEliminar.sort(function(a, b) { return b - a })
    var eliminados = 0
    for (var i = 0; i < aEliminar.length; i++) {
      try { sh.deleteRow(aEliminar[i]); eliminados++ } catch (eDel) {}
    }

    var msg = []
    if (transferidos) msg.push(transferidos + ' transferidos a Pacientes')
    if (aprobRows.length) msg.push(aprobRows.length + ' aprobados eliminados de recepción')
    if (rechRows.length) msg.push(rechRows.length + ' rechazados eliminados')
    ss.toast(msg.length ? msg.join(', ') : 'No hay Aprobados ni Rechazados para procesar. Pendientes se saltan.', 'PADDS', 4)
    _miniResumenForm(sh)
  } finally {
    lock.releaseLock()
  }
}

function _batchCopiarFormularios(formRows) {
  if (!formRows.length) return 0
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var pac = ss.getSheetByName(HOJA_PAC)
  if (!pac) return 0

  var lr = pac.getLastRow()
  var lc = pac.getLastColumn()

  var runMap = {}
  var pacData = []
  var maxId = 0
  if (lr >= 4) {
    pacData = pac.getRange(4, 1, lr - 3, lc).getValues()
    for (var i = 0; i < pacData.length; i++) {
      var rut = String(pacData[i][COL.RUN - 1] || '').replace(/\./g, '').replace(/\s/g, '').toUpperCase()
      if (rut) runMap[rut] = { row: i + 4, data: pacData[i] }
      var id = Number(pacData[i][COL.ID - 1])
      if (id > maxId) maxId = id
    }
  }

  var newRows = []
  var updated = 0

  for (var r = 0; r < formRows.length; r++) {
    var formRut = String(formRows[r][6] || '').replace(/\./g, '').replace(/\s/g, '').toUpperCase()
    if (!formRut) continue

    var fechaAtencion = formRows[r][5]
    var prestacion = String(formRows[r][10] || '').trim()
    var prestNorm = prestacion.toUpperCase()
      .replace(/[ÁÀÄÂ]/g,'A').replace(/[ÉÈËÊ]/g,'E').replace(/[ÍÌÏÎ]/g,'I')
      .replace(/[ÓÒÖÔ]/g,'O').replace(/[ÚÙÜÛ]/g,'U')
    var servicioCol = _SERVICIO_COL_MAP[prestNorm] || _SERVICIO_SI_MAP[prestNorm]
    var observaciones = String(formRows[r][16] || '').trim()
    var profesional = String(formRows[r][9] || '').trim()
    var visitaPerdida = String(formRows[r][12] || '').trim()
    var proxControlRaw = formRows[r][15]
    var proxControl = (proxControlRaw instanceof Date && !isNaN(proxControlRaw.getTime()))
      ? proxControlRaw : _parseDate(proxControlRaw)

    var fechaStr = (fechaAtencion instanceof Date && !isNaN(fechaAtencion.getTime()))
      ? Utilities.formatDate(fechaAtencion, Session.getScriptTimeZone(), 'dd/MM/yyyy')
      : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy')

    if (runMap[formRut]) {
      // ─── Existing patient: update service date and cumulative logs ────────
      var pacRow = runMap[formRut].row

      if (servicioCol && servicioCol <= lc) {
        var sv = _SERVICIO_SI_MAP[prestNorm] ? 'SI' : fechaAtencion
        if (sv) pac.getRange(pacRow, servicioCol, 1, 1).setValues([[sv]])
      }

      var obsParts = []
      if (prestacion) obsParts.push('Prestación: ' + prestacion)
      if (observaciones) obsParts.push('Obs: ' + observaciones)
      if (visitaPerdida === 'SI') obsParts.push('Visita perdida')
      if (proxControl) obsParts.push('Próximo control: ' + Utilities.formatDate(proxControl, Session.getScriptTimeZone(), 'dd/MM/yyyy'))
      if (prestNorm.indexOf('CONTROL CUIDADOR') >= 0) obsParts.push('Control cuidador')
      if (obsParts.length) {
        var obsLine = '[' + fechaStr + '] ' + obsParts.join(' | ')
        var obsOld = String(runMap[formRut].data[COL.OBSERVACIONES - 1] || '').trim()
        var obsLines = obsOld ? obsOld.split('\n') : []
        var obsDupe = false
        for (var oli = 0; oli < obsLines.length; oli++) {
          if (obsLines[oli] === obsLine) { obsDupe = true; break }
        }
        if (!obsDupe) {
          var obsCell = pac.getRange(pacRow, COL.OBSERVACIONES)
          obsCell.setValues([[obsOld ? obsOld + '\n' + obsLine : obsLine]])
          obsCell.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
        }
      }

      if (prestacion) {
        var mcOld = String(runMap[formRut].data[COL.CONTROLES_MISCELANEOS - 1] || '').trim()
        var mcLine = '[' + fechaStr + '] ' + prestacion
        var mcLines = mcOld ? mcOld.split('\n') : []
        var mcDupe = false
        for (var mli = 0; mli < mcLines.length; mli++) {
          if (mcLines[mli] === mcLine) { mcDupe = true; break }
        }
        if (!mcDupe) {
          var mcCell = pac.getRange(pacRow, COL.CONTROLES_MISCELANEOS)
          mcCell.setValues([[mcOld ? mcOld + '\n' + mcLine : mcLine]])
          mcCell.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
        }
      }

      if (profesional) {
        pac.getRange(pacRow, COL.EDITOR, 1, 1).setValues([[profesional]])
      }

      for (var fi = 0; fi < FORM_A_PAC.length; fi++) {
        var fpCol = FORM_A_PAC[fi][0] - 1, ppCol = FORM_A_PAC[fi][1] - 1
        if (fpCol < formRows[r].length && ppCol >= 0 && ppCol < lc
            && fpCol !== 5 && ppCol !== COL.OBSERVACIONES - 1) {
          var fv = formRows[r][fpCol]
          if (fv !== undefined && fv !== null && String(fv).trim() !== '') {
            var fvs = String(fv).trim()
            var valMap = _FORM_VAL_MAP[ppCol + 1]
            if (valMap && valMap[fvs]) fvs = valMap[fvs]
            pac.getRange(pacRow, ppCol + 1, 1, 1).setValues([[fvs]])
          }
        }
      }

      updated++
    } else {
      // ─── New patient: create row ────────────────────────────────────────
      maxId++
      var blank = []
      for (var c = 0; c < lc; c++) {
        blank.push(c === COL.ID - 1 ? maxId : c === COL.VITAL - 1 ? 'VIGENTE' : '')
      }

      for (var i = 0; i < FORM_A_PAC.length; i++) {
        var fCol = FORM_A_PAC[i][0] - 1, pCol = FORM_A_PAC[i][1] - 1
        if (fCol < formRows[r].length && pCol >= 0 && pCol < lc) {
          var val = formRows[r][fCol]
          if (val !== undefined && val !== null && String(val).trim() !== '') {
            var vs = String(val).trim()
            var valMap = _FORM_VAL_MAP[pCol + 1]
            blank[pCol] = valMap && valMap[vs] ? valMap[vs] : vs
          }
        }
      }

      if (servicioCol && servicioCol <= lc) {
        blank[servicioCol - 1] = _SERVICIO_SI_MAP[prestNorm] ? 'SI' : fechaAtencion
      }

      var obsParts = []
      if (prestacion) obsParts.push('Prestación: ' + prestacion)
      if (observaciones) obsParts.push('Obs: ' + observaciones)
      if (visitaPerdida === 'SI') obsParts.push('Visita perdida')
      if (proxControl) obsParts.push('Próximo control: ' + Utilities.formatDate(proxControl, Session.getScriptTimeZone(), 'dd/MM/yyyy'))
      if (obsParts.length) blank[COL.OBSERVACIONES - 1] = '[' + fechaStr + '] ' + obsParts.join(' | ')

      if (prestacion) blank[COL.CONTROLES_MISCELANEOS - 1] = '[' + fechaStr + '] ' + prestacion

      newRows.push(blank)
    }
  }

  if (newRows.length > 0) {
    var nr = lr + 1
    if (nr + newRows.length - 1 > pac.getMaxRows()) {
      pac.insertRowsAfter(pac.getMaxRows(), nr + newRows.length - pac.getMaxRows())
    }
    pac.getRange(nr, 1, newRows.length, lc).setValues(newRows)
    var rng = pac.getRange(nr, 1, newRows.length, lc)
    rng.setFontColor('#000000')
    rng.setFontWeight('normal')
    rng.setFontStyle('normal')
    pac.getRange(nr, COL.OBSERVACIONES, newRows.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
    pac.getRange(nr, COL.CONTROLES_MISCELANEOS, newRows.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
    pac.setRowHeights(nr, newRows.length, 24)
  }

  try { _borrarFilasVacias(pac, 4) } catch (eV) {}

  ss.toast(updated + ' pacientes actualizados, ' + newRows.length + ' pacientes creados desde formularios', 'PADDS', 4)
  return updated + newRows.length
}

function rechazarFormularios() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    ss.toast('Rechazando formularios…', 'PADDS', 1)
    var sh = ss.getSheetByName(HOJA_FORM)
    if (!sh) return

    var lr = sh.getLastRow()
    if (lr < 5) { SpreadsheetApp.getUi().alert('No hay formularios para rechazar.'); return }

    var data = sh.getRange(5, 1, lr - 4, sh.getLastColumn()).getValues()
    var email = ''
    try { email = Session.getActiveUser().getEmail() } catch(e) { email = 'usuario' }

    var rechRows = []
    for (var r = 0; r < data.length; r++) {
      if (String(data[r][2] || '').trim() !== 'Pendiente') continue
      rechRows.push(r + 5)
    }

    if (rechRows.length === 0) { ss.toast('No hay Pendientes para rechazar', 'Formulario', 3); return }

    var rng = sh.getRange(rechRows[0], 3, rechRows.length, 3)
    var vals = []
    for (var i = 0; i < rechRows.length; i++) vals.push(['Rechazado', email, new Date()])
    rng.setValues(vals)
    sh.getRange(rechRows[0], 5, rechRows.length, 1).setNumberFormat('dd/mm/yyyy hh:mm')
    for (var i = 0; i < rechRows.length; i++) _colorEstado(sh, rechRows[i], 'Rechazado', sh.getLastColumn())

    ss.toast(rechRows.length + ' formularios marcados como Rechazado (usa Sincronizar para eliminar)', 'Formulario', 4)
    _miniResumenForm(sh)
  } catch(e) {
    SpreadsheetApp.getUi().alert('Error al rechazar: ' + e.message)
  }
}

function limpiarFormulariosAprobados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_FORM)
  if (!sh) return
  var lr = sh.getLastRow()
  var lc = sh.getLastColumn()
  if (lr < 5) return

  ss.toast('Buscando formularios procesados…', 'PADDS', 1)
  var data = sh.getRange(5, 1, lr - 4, lc).getValues()
  var ui = SpreadsheetApp.getUi()
  var aLimpiar = 0

  for (var r = data.length - 1; r >= 0; r--) {
    if (String(data[r][2] || '').trim() === 'Aprobado' || String(data[r][2] || '').trim() === 'Rechazado') {
      aLimpiar++
    }
  }

  if (aLimpiar === 0) { ui.alert('No hay formularios aprobados o rechazados para limpiar.'); return }
  if (ui.alert('Limpiar', '¿Eliminar ' + aLimpiar + ' formularios aprobados o rechazados?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return

  var grupos = [], start = -1
  for (var r = data.length - 1; r >= 0; r--) {
    if (String(data[r][2] || '').trim() === 'Aprobado' || String(data[r][2] || '').trim() === 'Rechazado') {
      if (start < 0) { start = r + 5; grupos.push([start, 1]) }
      else { grupos[grupos.length - 1][1]++ }
    } else { start = -1 }
  }
  for (var g = 0; g < grupos.length; g++) {
    sh.deleteRows(grupos[g][0], grupos[g][1])
  }

  _miniResumenForm(sh)
  SpreadsheetApp.getActiveSpreadsheet().toast(aLimpiar + ' formularios eliminados de la hoja', 'PADDS', 3)
}

