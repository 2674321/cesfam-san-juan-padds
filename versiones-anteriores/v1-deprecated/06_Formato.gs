// ─── FORMATO: VALIDACIONES, FORMATO CONDICIONAL, HOJAS DE SISTEMA ──────────

// ─── MIGRAR LAYOUT LEGACY → NUEVO ──────────────────────────────────────────
// Transformaciones para pasar del layout anterior (con TELEFONO CUIDADOR
// desplazado y nombres legacy) al layout actual. Solo actúa si detecta el
// layout legacy — es idempotente.

function _migrarLayoutPacientes(ss, sh) {
  var lc = sh.getLastColumn()
  if (lc < 20) return

  var didMigrate = false

  // 1. Renombrar CCV → CCV VIGENTE CUIDADOR si está en col 28
  var h28 = String(sh.getRange(3, COL.CCV).getValue() || '').trim().toUpperCase()
  if (h28 === 'CCV') {
    sh.getRange(3, COL.CCV).setValue('CCV VIGENTE CUIDADOR')
    ss.toast('Col ' + COL.CCV + ' renombrada a "CCV VIGENTE CUIDADOR"', '', 1)
    didMigrate = true
  }

  // 2. Insertar F. CONTROL MEDICO en col 57 si col 57 es ESTADO CONTROL MEDICO
  var h57 = String(sh.getRange(3, 57).getValue() || '').trim().toUpperCase()
  if (h57.indexOf('ESTADO CONTROL MEDICO') >= 0) {
    ss.toast('Insertando columna F. CONTROL MEDICO (col 57)…', '', 1)
    sh.insertColumns(57, 1)
    sh.getRange(3, 57).setValue('F. CONTROL MEDICO')
    didMigrate = true
  }

  // 3. Asegurar CONTROLES MISCELÁNEOS en col 79, F. ALTA LPP en col 80
  var h78 = String(sh.getRange(3, 78).getValue() || '').trim().toUpperCase()
  var h79 = String(sh.getRange(3, 79).getValue() || '').trim().toUpperCase()
  var h80 = lc >= 80 ? String(sh.getRange(3, 80).getValue() || '').trim().toUpperCase() : ''

  var tieneMisc79 = h79.indexOf('MISCELÁNEOS') >= 0
  var tieneMisc80 = h80.indexOf('MISCELÁNEOS') >= 0

  if (tieneMisc79 && tieneMisc80) {
    // Duplicado: CONTROLES MISCELÁNEOS está en ambas, eliminar col 80
    ss.toast('Eliminando columna duplicada CONTROLES MISCELÁNEOS (col 80)…', '', 1)
    sh.deleteColumns(80, 1)
    didMigrate = true
  } else if (tieneMisc80 && !tieneMisc79) {
    // Layout viejo: col 79 = F. ALTA LPP, col 80 = CONTROLES MISCELÁNEOS → intercambiar
    ss.toast('Intercambiando columnas 79 y 80 (layout antiguo)…', '', 1)
    // Mover CONTROLES MISCELÁNEOS de 80 a 79
    sh.insertColumns(79, 1)
    sh.getRange(3, 79).setValue('CONTROLES MISCELÁNEOS')
    // La vieja CONTROLES MISCELÁNEOS ahora está en 81, la F. ALTA LPP en 80
    // Borrar el duplicado en 81
    sh.deleteColumns(81, 1)
    didMigrate = true
  } else if (!tieneMisc79 && !tieneMisc80) {
    // No existe — verificar col 78 como última opción
    var tieneMisc78 = h78.indexOf('MISCELÁNEOS') >= 0
    if (!tieneMisc78) {
      ss.toast('Insertando columna CONTROLES MISCELÁNEOS (col 79)…', '', 1)
      sh.insertColumns(79, 1)
      sh.getRange(3, 79).setValue('CONTROLES MISCELÁNEOS')
      didMigrate = true
    }
  }

  // 4. Renombrar columnas cuyos nombres cambiaron
  var renames = {
    60: ['CCV MEDICO', 'F. CCV MEDICO'],
    62: ['CCV ENFERMERIA', 'F. CSCV ENFERMERIA'],
    65: ['ATENCION PODOLOGICA', 'F. PODOLOGO'],
    67: ['CONSULTA NUTRICIONAL', 'F. NUTRICIONISTA'],
    70: ['ATENCION FONOAUDIOLOGA', 'ESTADO FONOAUDIOLOGA'],
    72: ['CONSULTA KINE', 'F. CONTROL KINESICO'],
    85: ['CONSULTA PSICOLOGICA', 'CONSULTA PSICOLOGA'],
  }
  for (var col in renames) {
    var c = parseInt(col)
    if (c > lc) continue
    var actual = String(sh.getRange(3, c).getValue() || '').trim().toUpperCase()
    if (actual === renames[col][0].toUpperCase()) {
      sh.getRange(3, c).setValue(renames[col][1])
      didMigrate = true
    }
  }

  if (didMigrate) {
    ss.toast('Layout migrado correctamente', '', 2)
  }

  // 5. Normalizar mayúsculas de encabezados según _COLUMNAS
  for (var _nc = 0; _nc < _COLUMNAS.length; _nc++) {
    var _ci = _COLUMNAS[_nc]
    if (!_ci || !_ci.name) continue
    if (_nc > lc) break
    var _actual = String(sh.getRange(3, _nc + 1).getValue() || '').trim()
    if (_actual && _actual !== _ci.name && _actual.toUpperCase() === _ci.name) {
      sh.getRange(3, _nc + 1).setValue(_ci.name)
      didMigrate = true
    }
  }

  // 6. Insertar CONSULTA TRABAJADORA SOCIAL en col 88 si col 88 es SIGGES
  if (lc >= 88) {
    var h88 = String(sh.getRange(3, 88).getValue() || '').trim().toUpperCase()
    if (h88.indexOf('SIGGES') >= 0) {
      ss.toast('Insertando columna CONSULTA TRABAJADORA SOCIAL (col 88)…', '', 1)
      sh.insertColumns(88, 1)
      sh.getRange(3, 88).setValue('CONSULTA TRABAJADORA SOCIAL')
      didMigrate = true
    }
  }
}

// ─── RECONSTRUIR PACIENTES ─────────────────────────────────────────────────

function _agregarColumnasFaltantes(ss, sh, targetCols) {
  var currentCols = sh.getLastColumn()
  if (currentCols >= targetCols) return currentCols

  ss.toast('Agregando columnas faltantes…', '', 1)
  sh.insertColumns(currentCols + 1, targetCols - currentCols)

  // Write header names for new columns in row 3
  for (var c = currentCols + 1; c <= targetCols; c++) {
    var colInfo = _COLUMNAS[c]
    if (colInfo) sh.getRange(3, c).setValue(colInfo.name)
  }

  return targetCols
}

function repintarPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ss.toast('Hoja Pacientes no encontrada', '', 3); return }
  ss.toast('Preparando reconstrucción funcional…', '', 1)

  _migrarLayoutPacientes(ss, sh)

  var targetCols = _COLUMNAS._count || 111
  var currentCols = _agregarColumnasFaltantes(ss, sh, targetCols)

  var data = sh.getDataRange().getValues()
  var rows = data.length
  if (rows < 3) { ss.toast('Pacientes sin datos', '', 3); return }

  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Reconstruir Pacientes',
    'Se actualizarán validaciones, tooltips y formatos numéricos.\n'
    + 'Los datos y el formato visual actual se mantienen.\n'
    + '¿Continuar?',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return

  // Normalizar todos los encabezados según _COLUMNAS
  for (var _hc = 1; _hc <= currentCols; _hc++) {
    var _ci = _COLUMNAS[_hc]
    if (_ci && _ci.name) sh.getRange(3, _hc).setValue(_ci.name)
  }

  for (var _pr = 0; _pr < data.length; _pr++) {
    for (var _pc = data[_pr].length; _pc < currentCols; _pc++) {
      data[_pr].push('')
    }
  }

  if (rows >= 4) {
    var dataRows = data.slice(3)
    sh.getRange(4, 1, dataRows.length, currentCols).setValues(dataRows)
  }
  aplicarFormatoFuncional(sh, rows, currentCols)
  ponerTooltipsPacientes()
  _refrescarFormatoCondicional(sh, rows, currentCols)
  ss.toast('Pacientes reconstruido: validaciones, tooltips y formatos numéricos actualizados', '', 4)
}

function formatearPacientesVisual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ss.toast('Hoja Pacientes no encontrada', '', 3); return }
  var lr = sh.getLastRow()
  var lc = sh.getLastColumn()
  if (lr < 3) { ss.toast('Pacientes sin datos', '', 3); return }
  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Formatear Pacientes',
    'Se restaurarán colores, bordes, fuentes, anchos de columna\n'
    + 'y los encabezados de fila 1-3 (secciones, buscador, nombres de columna).\n'
    + 'Los datos no se pierden.\n¿Continuar?',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return
  _migrarLayoutPacientes(ss, sh)
  var targetCols = _COLUMNAS._count || 111
  _agregarColumnasFaltantes(ss, sh, targetCols)
  lc = sh.getLastColumn()
  _aplicarFormatoVisual(sh, lr, lc)
  ss.toast('Formato visual restaurado', '', 3)
}

function aplicarFormatoFuncional(sh, lr, lc) {
  var COLS_FECHA_EXTRAS = [9, 13, 14, 15, COL.CCV, COL.F_ALTA_LPP, COL.ZARIT, 87, 88]
  var colsFecha = []
  for (var fi = 0; fi < CONTROL_COLS.length; fi++) colsFecha.push(CONTROL_COLS[fi][1])
  for (var fi = 101; fi <= 108; fi++) colsFecha.push(fi)
  for (var fi = 0; fi < COLS_FECHA_EXTRAS.length; fi++) colsFecha.push(COLS_FECHA_EXTRAS[fi])
  colsFecha.sort(function(a,b) { return a - b })
  if (lr >= 4) {
    var dataRows = lr - 3
    var runCols = [8, 18]
    var runFmt = []
    for (var ri = 0; ri < runCols.length; ri++) {
      runFmt.push('R4C' + runCols[ri] + ':R' + lr + 'C' + runCols[ri])
    }
    sh.getRangeList(runFmt).setNumberFormat('@')

    var dateFmtRanges = []
    for (var fi = 0; fi < colsFecha.length; fi++) {
      var fc = colsFecha[fi]
      if (fc <= lc) dateFmtRanges.push('R4C' + fc + ':R' + lr + 'C' + fc)
    }
    sh.getRangeList(dateFmtRanges).setNumberFormat('dd/mm/yyyy')
  }

  sh.setFrozenRows(3)
  sh.setFrozenColumns(5)

  if (lr >= 3) {
    var rng = sh.getRange(3, 1, lr - 2, lc)
    var f = rng.getFilter()
    if (f) f.remove()
    rng.createFilter()
  }

  _aplicarValidaciones(sh, lr, lc, colsFecha)

  if (lr >= 4 && lc >= COL.CAPACITACIONES_INI) {
    var capFin = Math.min(COL.CAPACITACIONES_FIN, lc)
    var capCount = capFin - COL.CAPACITACIONES_INI + 1
    if (capCount > 0) {
      sh.getRange(4, COL.CAPACITACIONES_INI, lr - 3, capCount).setDataValidation(null)
      var dateVal = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build()
      sh.getRange(4, COL.CAPACITACIONES_INI, lr - 3, capCount).setDataValidation(dateVal)
    }
  }

  if (lr >= 4) {
    if (lc >= COL.OBSERVACIONES) sh.getRange(4, COL.OBSERVACIONES, lr - 3, lc - COL.OBSERVACIONES + 1).setDataValidation(null)
  }
}

function _aplicarFormatoVisual(sh, lr, lc) {
  for (var i = 0; i < Math.min(PAC_ANCHOS.length, lc); i++)
    sh.setColumnWidth(i + 1, PAC_ANCHOS[i])

  sh.setRowHeight(1, 40)
  sh.getRange(1, 1, 1, lc)
    .setFontFamily('Arial').setFontSize(13).setFontWeight('bold')
    .setVerticalAlignment('middle').setHorizontalAlignment('center')

  for (var s = 0; s < PAC_SECCIONES.length; s++) {
    var sec = PAC_SECCIONES[s]
    if (sec.ini > lc) break
    var fin = Math.min(sec.fin, lc)
    sh.getRange(1, sec.ini, 1, fin - sec.ini + 1)
      .setBackground(sec.bg).setFontColor(sec.fg)
    sh.getRange(1, sec.ini).setValue(sec.nombre)
  }

  sh.setRowHeight(2, 30)
  sh.getRange(2, 1, 1, lc).setBackground('#fff9c4')
    .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, '#d0d0d0', SpreadsheetApp.BorderStyle.SOLID)
  sh.getRange(2, 1).setValue('🔍').setFontSize(14).setHorizontalAlignment('center')
    .setNote('Buscar paciente: escribe el nombre (o parte) en la celda B2 y presiona Enter. La agenda se posicionará en el primer resultado.')

  var nf = SpreadsheetApp.newTextStyle().setFontFamily('Arial').setFontSize(11)
    .setForegroundColor('#666666').build()
  sh.getRange(2, 2, 1, 4).setTextStyle(nf).setHorizontalAlignment('center')

  sh.setRowHeight(3, 28)
  sh.getRange(3, 1, 1, lc).setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontColor('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true)
  for (var s = 0; s < PAC_SECCIONES.length; s++) {
    var sec = PAC_SECCIONES[s]
    if (sec.ini > lc) break
    var fin = Math.min(sec.fin, lc)
    sh.getRange(3, sec.ini, 1, fin - sec.ini + 1).setBackground(sec.bg2 || sec.bg)
  }

  if (lr >= 4) {
    var dataRows = lr - 3
    var bgs = []
    for (var r = 0; r < dataRows; r++) {
      var bgRow = []
      var frozenBg = r % 2 === 0 ? '#f0f4f8' : '#e4ecf3'
      var alt = r % 2 === 0 ? '#fafafa' : '#ffffff'
      for (var c = 0; c < lc; c++) {
        bgRow.push(c < 5 ? frozenBg : alt)
      }
      bgs.push(bgRow)
    }
    sh.getRange(4, 1, dataRows, lc).setBackgrounds(bgs)
      .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#d0d0d0', SpreadsheetApp.BorderStyle.SOLID)
    sh.getRange(4, COL.CONTROLES_MISCELANEOS, dataRows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    sh.getRange(4, COL.OBSERVACIONES, dataRows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    sh.setRowHeights(4, dataRows, 24)
  }

  try { _refrescarFormatoCondicional(sh, lr, lc) } catch(e) {}
  try { _colorearEdadPorEMPA(sh, lr, lc) } catch(e) {}
  try { colorearEMPAUsuario() } catch(e) {}
  try { colorearZARIT() } catch(e) {}
  try { colorearTS() } catch(e) {}
  try { colorearPS() } catch(e) {}
}

// ─── VALIDACIONES ──────────────────────────────────────────────────────────

function _aplicarValidaciones(sh, lr, lc, saltar) {
  var cols = Object.keys(PAC_VALIDACIONES).map(Number).filter(function(c) {
    return c <= lc && (!saltar || saltar.indexOf(c) === -1)
  }).sort(function(a, b) { return a - b })
  if (!cols.length) return

  var groups = []
  var cur = [cols[0]]
  for (var i = 1; i < cols.length; i++) {
    if (cols[i] === cols[i - 1] + 1 && PAC_VALIDACIONES[cols[i]] === PAC_VALIDACIONES[cols[i - 1]]) {
      cur.push(cols[i])
    } else {
      groups.push(cur)
      cur = [cols[i]]
    }
  }
  groups.push(cur)

  if (lr >= 4) {
    var rowCount = lr - 3
    for (var gi = 0; gi < groups.length; gi++) {
      var g = groups[gi]
      var dv = SpreadsheetApp.newDataValidation()
        .requireValueInList(PAC_VALIDACIONES[g[0]], true).setAllowInvalid(true).build()
      sh.getRange(4, g[0], rowCount, g.length).setDataValidation(dv)
    }
  }

  // Batch numberFormat for all CONTROL_COLS
  var fmtRanges = []
  for (var _ci = 0; _ci < CONTROL_COLS.length; _ci++) {
    var _fc = CONTROL_COLS[_ci][1]
    if (_fc <= lc) fmtRanges.push('R4C' + _fc + ':R' + lr + 'C' + _fc)
  }
  if (fmtRanges.length) sh.getRangeList(fmtRanges).setNumberFormat('dd/mm/yyyy')
}

// ─── FORMATO CONDICIONAL ──────────────────────────────────────────────────

function _refrescarFormatoCondicional(sh, lr, lc) {
  _aplicarValidaciones(sh, lr, lc)

  var cols = {}
  for (var c = 0; c < CONTROL_COLS.length; c++) {
    var ec = CONTROL_COLS[c][2]
    if (ec) cols[ec] = true
  }
  var statusCols = Object.keys(cols).map(Number).sort(function(a,b) { return a - b })

  var rules = []
  for (var i = 0; i < statusCols.length; i++) {
    var col = statusCols[i]
    var rng = sh.getRange(4, col, lr - 3, 1)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('AL DIA').setBackground('#c8e6c9').setFontColor('#2e7d32')
      .setRanges([rng]).build())
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('POR VENCER').setBackground('#ffe0b2').setFontColor('#e65100').setBold(true)
      .setRanges([rng]).build())
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('VENCIDO').setBackground('#ffcdd2').setFontColor('#c62828').setBold(true)
      .setRanges([rng]).build())
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('PENDIENTE').setBackground('#fff9c4').setFontColor('#f9a825')
      .setRanges([rng]).build())
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('N/A').setBackground('#f5f5f5').setFontColor('#999999')
      .setRanges([rng]).build())
  }
  var vitalRng = sh.getRange(4, 1, lr - 3, lc)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F4="FALLECIDO"')
    .setFontColor('#aaaaaa').setStrikethrough(true)
    .setRanges([vitalRng]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F4="SUSPENDIDO"')
    .setFontColor('#888888').setItalic(true)
    .setRanges([vitalRng]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F4="EGRESO"')
    .setFontColor('#888888')
    .setRanges([vitalRng]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F4="VIGENTE"')
    .setFontColor('#000000').setBold(false).setStrikethrough(false).setItalic(false)
    .setRanges([vitalRng]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F4="ALTA"')
    .setFontColor('#1565c0').setBold(false).setStrikethrough(false).setItalic(false)
    .setRanges([vitalRng]).build())

  // SECTOR (col B) conditional coloring
  var sectorRng = sh.getRange(4, 2, lr - 3, 1)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('VERDE').setBackground('#c8e6c9').setFontColor('#2e7d32')
    .setRanges([sectorRng]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('AMARILLO').setBackground('#fff9c4').setFontColor('#f9a825')
    .setRanges([sectorRng]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('NARANJO').setBackground('#ffe0b2').setFontColor('#e65100')
    .setRanges([sectorRng]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pendiente').setBackground('#f5f5f5').setFontColor('#999999')
    .setRanges([sectorRng]).build())

  // PRIORIDAD GENERAL (col 109) conditional coloring
  var prioridadRng = sh.getRange(4, 109, lr - 3, 1)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('URGENTE').setBackground('#ffcdd2').setFontColor('#c62828').setBold(true)
    .setRanges([prioridadRng]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('POR REVISAR').setBackground('#ffe0b2').setFontColor('#e65100').setBold(true)
    .setRanges([prioridadRng]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('AL DIA').setBackground('#c8e6c9').setFontColor('#2e7d32')
    .setRanges([prioridadRng]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('N/A').setBackground('#f5f5f5').setFontColor('#999999')
    .setRanges([prioridadRng]).build())

  var dateNAs = []
  for (var _di = 0; _di < CONTROL_COLS.length; _di++) {
    var _dc = CONTROL_COLS[_di][1]
    if (_dc <= lc) dateNAs.push(_dc)
  }
  for (var _di = 0; _di < dateNAs.length; _di++) {
    var _drng = sh.getRange(4, dateNAs[_di], lr - 3, 1)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('N/A').setBackground('#f5f5f5').setFontColor('#999999')
      .setRanges([_drng]).build())
  }
  try { sh.setConditionalFormatRules(rules) } catch(e) {}
  try { colorearCapacitaciones() } catch(e) {}
  try { colorearCCV() } catch(e) {}
  try { colorearEMPAUsuario() } catch(e) {}
  try { colorearZARIT() } catch(e) {}
  try { colorearTS() } catch(e) {}
  try { colorearPS() } catch(e) {}
  try { colorearSector(sh, lr) } catch(e) {}
  try { colorearEstado(sh, lr) } catch(e) {}
}

// ─── SECTOR Y ESTADO (colorear por valor) ─────────────────────────────────

function _aplicarColorPorValor(sh, lr, col, coloresMap) {
  if (!sh) sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pacientes')
  if (!sh) return
  if (!lr) lr = sh.getLastRow()
  if (lr < 4 || sh.getLastColumn() < col) return
  var data = sh.getRange(4, col, lr - 3, 1).getValues()
  var bgs = [], fgs = []
  for (var r = 0; r < data.length; r++) {
    var v = String(data[r][0] || '').trim().toUpperCase()
    var m = coloresMap[v]
    bgs.push([m ? m[0] : null])
    fgs.push([m ? m[1] : '#000000'])
  }
  sh.getRange(4, col, lr - 3, 1).setBackgrounds(bgs)
  sh.getRange(4, col, lr - 3, 1).setFontColors(fgs)
}

function colorearSector(sh, lr) { _aplicarColorPorValor(sh, lr, 2, _SECTOR_COLORS) }
function colorearEstado(sh, lr) { _aplicarColorPorValor(sh, lr, 6, _ESTADO_COLORS) }

// ─── COLOREAR POR VIGENCIA DE FECHA ──────────────────────────────────────

function _colorCeldaPorVigencia(val, meses) {
  var fecha = _parseDate(val)
  if (fecha) {
    var hoy = new Date()
    var dif = (hoy.getFullYear() - fecha.getFullYear()) * 12 + (hoy.getMonth() - fecha.getMonth())
    return dif <= meses ? '#c8e6c9' : '#ffcdd2'
  }
  return '#fff9c4'
}

function colorearCapacitaciones() {
  var params = leerParametros()
  var meses = params['CAPACITACIONES']
  if (!meses) meses = 12

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pacientes')
  if (!sh) return
  var lr = sh.getLastRow()
  if (lr < 4) return
  var lc = sh.getLastColumn()
  if (lc < 99) return

  var data = sh.getRange(4, COL.CAPACITACIONES_INI, lr - 3, 8).getValues()
  var bgs = new Array(data.length)
  for (var r = 0; r < data.length; r++) {
    bgs[r] = new Array(8)
    for (var c = 0; c < 8; c++) {
      bgs[r][c] = _colorCeldaPorVigencia(data[r][c], meses)
    }
  }

  sh.getRange(4, COL.CAPACITACIONES_INI, lr - 3, 8).setBackgrounds(bgs)
}

function colorearFilaCapacitacion(row, sh, _params, rowData) {
  if (!_params) _params = leerParametros()
  var meses = _params['CAPACITACIONES']
  if (!meses) meses = 12

  var vals = rowData ? [rowData.slice(COL.CAPACITACIONES_INI - 1, COL.CAPACITACIONES_INI - 1 + 8)] : sh.getRange(row, COL.CAPACITACIONES_INI, 1, 8).getValues()
  var bgs = [vals[0].map(function(v) { return _colorCeldaPorVigencia(v, meses) })]
  sh.getRange(row, COL.CAPACITACIONES_INI, 1, 8).setBackgrounds(bgs)
}

// ─── COLOREAR POR FECHA GENÉRICO (CCV, EMPA USUARIO) ─────────────────────

function _aplicarColorPorFecha(col, paramName) {
  var params = leerParametros()
  var meses = params[paramName] || 12
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pacientes')
  if (!sh) return
  var lr = sh.getLastRow()
  if (lr < 4 || sh.getLastColumn() < col) return
  var data = sh.getRange(4, col, lr - 3, 1).getValues()
  var bgs = data.map(function(r) { return [_colorCeldaPorVigencia(r[0], meses)] })
  sh.getRange(4, col, lr - 3, 1).setBackgrounds(bgs)
}

function _aplicarColorPorFechaFila(row, sh, col, paramName, rowData) {
  var params = leerParametros()
  var meses = params[paramName] || 12
  var val = rowData ? rowData[col - 1] : sh.getRange(row, col, 1, 1).getValues()[0][0]
  sh.getRange(row, col).setBackground(_colorCeldaPorVigencia(val, meses))
}

function colorearCCV() { _aplicarColorPorFecha(COL.CCV, 'CCV VIGENTE CUIDADOR') }
function colorearFilaCCV(row, sh, _params, rowData) { _aplicarColorPorFechaFila(row, sh, COL.CCV, 'CCV VIGENTE CUIDADOR', rowData) }
function colorearEMPAUsuario() { _aplicarColorPorFecha(COL.EMPA_USUARIO_FECHA, 'EMPA/EMPAM USUARIO') }
function colorearFilaEMPAUsuario(row, sh, _params, rowData) { _aplicarColorPorFechaFila(row, sh, COL.EMPA_USUARIO_FECHA, 'EMPA/EMPAM USUARIO', rowData) }
function colorearZARIT() { _aplicarColorPorFecha(COL.ZARIT, 'ZARIT') }
function colorearFilaZARIT(row, sh, _params, rowData) { _aplicarColorPorFechaFila(row, sh, COL.ZARIT, 'ZARIT', rowData) }
function colorearTS() { _aplicarColorPorFecha(88, 'CONSULTA TRABAJADORA SOCIAL') }
function colorearFilaTS(row, sh, _params, rowData) { _aplicarColorPorFechaFila(row, sh, 88, 'CONSULTA TRABAJADORA SOCIAL', rowData) }
function colorearPS() { _aplicarColorPorFecha(87, 'CONSULTA PSICOLOGA') }
function colorearFilaPS(row, sh, _params, rowData) { _aplicarColorPorFechaFila(row, sh, 87, 'CONSULTA PSICOLOGA', rowData) }

// ─── LIMPIAR FORMATO EN COLUMNAS DE TEXTO LIBRE ──────────────────────────

function _FORMATEADAS() {
  var s = {}
  // CONTROL_COLS status columns
  for (var i = 0; i < CONTROL_COLS.length; i++) {
    if (CONTROL_COLS[i][2]) s[CONTROL_COLS[i][2]] = true
    s[CONTROL_COLS[i][1]] = true
  }
  // Additional formatted columns
  s[2] = true  // SECTOR
  s[6] = true  // ESTADO
  s[28] = true // CCV

  s[87] = true // CONSULTA PSICOLOGA
  s[109] = true // PRIORIDAD
  for (var c = 101; c <= 108; c++) s[c] = true // CAPACITACIONES
  s[10] = true // EDAD USUARIO (EMPA coloring)
  s[19] = true // EDAD CUIDADOR (EMPA coloring)
  s[COL.EMPA_USUARIO_FECHA] = true // EMPA USUARIO fecha (EMPA date coloring)
  return s
}
var _COLS_FORMATEADAS = _FORMATEADAS()

function _limpiarFormatoCelda(sh, row, col) {
  if (_COLS_FORMATEADAS[col]) return
  sh.getRange(row, col)
    .setFontWeight('normal').setFontStyle('normal')
    .setStrikethrough(false).setUnderline(false)
}

// ─── REPARAR FÓRMULAS ─────────────────────────────────────────────────────

function repararFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var viejoNom = 'Pacientes (anterior)'
  var nuevoNom = 'Pacientes'
  var total = 0

  for (var si = 0; si < ss.getSheets().length; si++) {
    var sh = ss.getSheets()[si]
    if (sh.getSheetName().indexOf(viejoNom) !== -1) continue
    var rng = sh.getDataRange()
    var formulas = rng.getFormulas()
    var cambio = false

    for (var r = 0; r < formulas.length; r++) {
      for (var c = 0; c < formulas[r].length; c++) {
        var f = formulas[r][c]
        if (f === '') continue
        var nf = f

        nf = nf.replace(/'?Pacientes \(anterior\)'?!?/g, "'" + nuevoNom + "'!")
        nf = nf.replace(/\$B\$?\d*:\$B\$?\d*/g, function(m) { return m.replace(/\$B/g, '$F') })
        nf = nf.replace(/\$C\$?\d*:\$C\$?\d*/g, function(m) { return m.replace(/\$C/g, '$B') })

        if (nf !== f) {
          formulas[r][c] = nf
          cambio = true
          total++
        }
      }
    }

    if (cambio) rng.setFormulas(formulas)
  }

  ss.toast(total + ' fórmulas reparadas en Dashboard', '', 4)
  try { recalcularTodo() } catch(e) {}
}

// ─── CREAR HOJA PARÁMETROS ─────────────────────────────────────────────────

function crearParametros() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Creando hoja Parámetros…', '', 1)
  var old = ss.getSheetByName('Parametros')
  if (old) ss.deleteSheet(old)

  var sh = ss.insertSheet('Parametros')
  sh.setTabColor('#1a237e')
  sh.setColumnWidth(1, 340)
  sh.setColumnWidth(2, 110)
  sh.setColumnWidth(3, 580)

  // ── Paleta ──
  var NAVY = '#1a237e', MID = '#283593', ACCENT = '#3f51b5'
  var LIGHT = '#e8eaf6', WHITE = '#ffffff', BORDER = '#c5cae9'
  var LIGHT_GREEN = '#e8f5e9', LIGHT_BLUE = '#e3f2fd', LIGHT_ORANGE = '#fff3e0'
  var LIGHT_PURPLE = '#f3e5f5', LIGHT_TEAL = '#e0f2f1', LIGHT_PINK = '#fce4ec'
  var BS = SpreadsheetApp.BorderStyle.SOLID

  function _secColor(i) {
    var palette = [LIGHT_GREEN, LIGHT_BLUE, LIGHT_ORANGE, LIGHT_PURPLE, LIGHT_TEAL, LIGHT_PINK]
    return palette[i % palette.length]
  }

  // ── Encabezado principal ──
  sh.getRange(1, 1, 1, 3).merge()
  sh.getRange(1, 1)
    .setValue('⚙  PARÁMETROS DE VIGENCIA  —  PADDS 2026')
    .setFontFamily('Calibri').setFontSize(22).setFontWeight('bold')
    .setFontColor(WHITE).setBackground(NAVY)
    .setVerticalAlignment('middle').setHorizontalAlignment('left')
  sh.setRowHeight(1, 52)

  sh.getRange(2, 1, 1, 3).merge()
  sh.getRange(2, 1)
    .setValue('Configure aquí los plazos de vigencia de cada prestación. Al modificar un valor, el cambio se aplica automáticamente al ejecutar "Recalcular estados" o "Reconstruir Pacientes".')
    .setFontFamily('Calibri').setFontSize(11).setFontColor('#7986cb')
    .setBackground(LIGHT).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  sh.setRowHeight(2, 30)

  // ── Grupos de parámetros ──
  var GRUPOS = [
    {
      nombre: '🩺  CONTROLES MÉDICOS Y EXÁMENES',
      color: '#1565c0',
      items: [
        ['Vigencia examenes usuario',    6, 'Plazo máximo entre cada toma de exámenes del usuario. Si la última fecha registrada supera este período, el estado del control pasa a "VENCIDO".'],
        ['Vigencia examenes cuidador',  12, 'Plazo máximo entre exámenes del cuidador. Similar al del usuario, pero con un período más amplio (12 meses).'],
        ['Vigencia control medico',      6, 'Tiempo máximo entre controles médicos. Si han pasado más de 6 meses desde el último control, se marca como vencido y se activa la alerta de prioridad.'],
        ['Vigencia PIC-1C',             12, 'Vigencia del Plan de Intervención Individual (1 componente). Se debe actualizar una vez al año para mantener el plan activo y al día.'],
        ['Vigencia PIC-2C',             12, 'Vigencia del Plan de Intervención Individual (2 componentes). Al igual que el PIC-1C, se renueva anualmente.'],
      ],
    },
    {
      nombre: '💓  CONTROLES CARDIOVASCULARES Y ESPECIALIDADES',
      color: '#c62828',
      items: [
        ['Vigencia CCV medico',          6, 'Plazo entre controles cardiovasculares realizados por médico. El control CCV debe repetirse cada 6 meses para pacientes con riesgo cardiovascular.'],
        ['Vigencia CSCV enfermeria',     6, 'Plazo entre controles cardiovasculares realizados por enfermería. La periodicidad es la misma que la del médico: cada 6 meses.'],
        ['Vigencia podologo',            6, 'Plazo entre atenciones de podología. El paciente debe ser evaluado cada 6 meses, especialmente si tiene diabetes o riesgo de pie diabético.'],
        ['Vigencia nutricionista',       6, 'Plazo entre controles nutricionales. Se recomienda evaluación cada 6 meses para ajustar el plan alimentario según la evolución del paciente.'],
        ['Vigencia fonoaudiologa',      12, 'Plazo entre atenciones de fonoaudiología. La vigencia es anual, dado que los tratamientos fonoaudiológicos suelen tener seguimientos más espaciados.'],
        ['Vigencia control kinesico',    6, 'Plazo entre controles kinésicos. El paciente debe ser reevaluado cada 6 meses para verificar la evolución de su funcionalidad y movilidad.'],
        ['Vigencia odontologia',        12, 'Plazo entre atenciones odontológicas. La revisión dental se programa una vez al año, salvo que existan patologías que requieran controles más frecuentes.'],
      ],
    },
    {
      nombre: '📋  EMPA / EMPAM Y PREVENCIÓN',
      color: '#2e7d32',
      items: [
        ['Vigencia EMPA/EMPAM cuidador',12, 'Vigencia del examen preventivo del cuidador principal. Se debe renovar anualmente para mantener actualizado el estado de salud del cuidador.'],
        ['Vigencia EMPA/EMPAM usuario', 12, 'Vigencia del examen preventivo del usuario. El EMPA o EMPAM debe aplicarse una vez al año. El color de la fecha en columna 82 indica si está al día o vencido.'],
        ['CCV VIGENTE CUIDADOR',        12, 'Vigencia de la CCV (cardiopatía descompensada) del cuidador. Se controla anualmente y la celda se colorea según la fecha registrada en columna 28.'],
      ],
    },
    {
      nombre: '🧠  SOCIAL / PSICOLÓGICO',
      color: '#6a1b9a',
      items: [
        ['ZARIT',                       12, 'Plazo entre aplicaciones de la escala Zarit (sobrecarga del cuidador). Debe aplicarse una vez al año. La fecha en columna 85 se colorea según su vigencia.'],
        ['CONSULTA TRABAJADORA SOCIAL', 12, 'Plazo entre consultas de trabajadora social. La fecha registrada en columna 88 se colorea según este plazo: verde si está al día, rojo si ha vencido.'],
        ['CONSULTA PSICOLOGA',          12, 'Plazo entre consultas psicológicas. Funciona igual que la trabajadora social: la fecha en columna 87 se evalúa contra este valor para colorear la celda.'],
      ],
    },
    {
      nombre: '📚  CAPACITACIONES Y AVISOS',
      color: '#e65100',
      items: [
        ['CAPACITACIONES',               6, 'Vigencia de las capacitaciones registradas (columnas 101 a 108). Cada taller o capacitación tiene una validez de 6 meses; pasado ese tiempo se considera vencida y se colorea en rojo.'],
        ['Días aviso',                  15, 'Cantidad de días antes del vencimiento de un control para que el sistema lo marque como "POR VENCER". Si un control vence en 15 días o menos, aparece en naranja como alerta temprana.'],
      ],
    },
  ]

  // ── Render ──
  var row = 3
  var BS2 = SpreadsheetApp.BorderStyle.SOLID_MEDIUM

  for (var g = 0; g < GRUPOS.length; g++) {
    var grupo = GRUPOS[g]
    var secBg = _secColor(g)

    // Fila de grupo
    sh.getRange(row, 1, 1, 3).merge()
    sh.getRange(row, 1)
      .setValue(grupo.nombre)
      .setFontFamily('Calibri').setFontSize(12).setFontWeight('bold')
      .setFontColor(WHITE).setBackground(grupo.color)
      .setVerticalAlignment('middle').setHorizontalAlignment('left')
      .setBorder(true, true, true, true, true, true, BORDER, BS)
    sh.setRowHeight(row, 28)
    row++

    // Subtítulos
    sh.getRange(row, 1).setValue('PARÁMETRO')
      .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold')
      .setFontColor(grupo.color).setBackground('#f5f5f5')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, BORDER, BS2)
    sh.getRange(row, 2).setValue('MESES')
      .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold')
      .setFontColor(grupo.color).setBackground('#f5f5f5')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, BORDER, BS2)
    sh.getRange(row, 3).setValue('DESCRIPCIÓN')
      .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold')
      .setFontColor(grupo.color).setBackground('#f5f5f5')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, BORDER, BS2)
    sh.setRowHeight(row, 22)
    row++

    // Items
    for (var i = 0; i < grupo.items.length; i++) {
      var item = grupo.items[i]
      var bg = i % 2 === 0 ? WHITE : secBg

      sh.getRange(row, 1).setValue(item[0])
        .setFontFamily('Calibri').setFontSize(11).setFontWeight('bold')
        .setFontColor(NAVY).setBackground(bg)
        .setVerticalAlignment('middle')
        .setBorder(false, true, false, true, false, false, BORDER, BS)
        .setNote(item[2])

      sh.getRange(row, 2).setValue(item[1])
        .setFontFamily('Calibri').setFontSize(14).setFontWeight('bold')
        .setFontColor(grupo.color).setBackground(WHITE)
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setNumberFormat('0')
        .setBorder(true, true, true, true, false, false, grupo.color, BS2)

      sh.getRange(row, 3).setValue(item[2])
        .setFontFamily('Calibri').setFontSize(10).setFontColor('#424242')
        .setBackground(bg).setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setBorder(false, true, false, true, false, false, BORDER, BS)

      sh.setRowHeight(row, 40)
      row++
    }

    // Separador entre grupos
    sh.setRowHeight(row, 6)
    sh.getRange(row, 1, 1, 3).setBackground('#f5f5f5').setBorder(false, false, false, false, false, false)
    row++
  }

  // ── Footer ──
  sh.getRange(row, 1, 1, 3).merge()
  sh.getRange(row, 1)
    .setValue('💡  Los cambios se aplican automáticamente al ejecutar "Recalcular estados" o "Reconstruir Pacientes". También se actualiza al editar una fecha en Pacientes.')
    .setFontFamily('Calibri').setFontSize(10).setFontColor('#7986cb')
    .setBackground(LIGHT).setVerticalAlignment('middle').setHorizontalAlignment('center')
  sh.setRowHeight(row, 28)
  row++

  sh.getRange(row, 1, 1, 3).merge()
  sh.getRange(row, 1)
    .setValue('PADDS 2026 — Centro de Salud Familiar San Joaquín')
    .setFontFamily('Calibri').setFontSize(9).setFontColor('#b0bec5')
    .setBackground(LIGHT).setVerticalAlignment('middle').setHorizontalAlignment('center')
  sh.setRowHeight(row, 22)

  sh.setFrozenRows(1)

  var maxRows = sh.getMaxRows()
  var neededRows = row + 1
  if (maxRows > neededRows) sh.deleteRows(neededRows, maxRows - neededRows + 1)

  ss.toast('Hoja Parámetros creada/actualizada', '', 4)
}

// ─── CREAR HOJA INSTRUCCIONES ──────────────────────────────────────────────

function crearInstrucciones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var old = ss.getSheetByName('Referencia Columnas')
  if (old) ss.deleteSheet(old)

  var sh = ss.insertSheet('Referencia Columnas')
  sh.setTabColor('#1b5e20')
  sh.setColumnWidth(1, 50)
  sh.setColumnWidth(2, 220)
  sh.setColumnWidth(3, 340)
  sh.setColumnWidth(4, 320)
  sh.setColumnWidth(5, 200)

  var DARK = '#1b5e20', MID = '#2e7d5b', LIGHT = '#e8f5e9', WHITE = '#ffffff'
  var BS = SpreadsheetApp.BorderStyle.SOLID

  sh.getRange(1, 1, 1, 5).merge()
  sh.getRange(1, 1).setValue('PADDS 2026  —  Referencia de Columnas')
    .setFontFamily('Calibri').setFontSize(20).setFontWeight('bold')
    .setFontColor(WHITE).setBackground(DARK)
    .setVerticalAlignment('middle')
  sh.setRowHeight(1, 44)

  sh.getRange(2, 1, 1, 5).merge()
  sh.getRange(2, 1).setValue('Pase el mouse sobre los encabezados en Pacientes (fila 3) para ver la descripción · Menú → Herramientas → Poner Tooltips')
    .setFontFamily('Calibri').setFontSize(11).setFontStyle('italic')
    .setFontColor('#e8f5e9').setBackground(MID)
    .setVerticalAlignment('middle')
  sh.setRowHeight(2, 24)

  // Header columns
  var HD = ['N°', 'CAMPO', 'DESCRIPCIÓN', 'VALORES', 'AUTOMATIZACIÓN']
  for (var c = 0; c < HD.length; c++) {
    sh.getRange(3, c + 1).setValue(HD[c])
      .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
      .setFontColor(WHITE).setBackground(MID)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true)
  }
  sh.setRowHeight(3, 24)

  var row = 4
  for (var si = 0; si < PAC_SECCIONES.length; si++) {
    var sec = PAC_SECCIONES[si]
    var colBg = sec.bg2 || sec.bg
    var colBgLight = _lightenHex(colBg, 130, 100, 110)

    // Section header
    sh.getRange(row, 1, 1, 5).merge()
    sh.getRange(row, 1).setValue(sec.nombre)
      .setFontFamily('Calibri').setFontSize(11).setFontWeight('bold')
      .setFontColor(WHITE).setBackground(colBg)
      .setVerticalAlignment('middle')
    sh.setRowHeight(row, 22)
    row++

    // Column rows
    for (var c = sec.ini; c <= Math.min(sec.fin, _COLUMNAS.length); c++) {
      var col = _COLUMNAS[c]
      if (!col) continue
      var alt = (c - sec.ini) % 2 === 0
      var bg = alt ? WHITE : '#f5f7fa'

      sh.getRange(row, 1).setValue(col.n)
        .setFontFamily('Consolas,monospace').setFontSize(9).setFontColor('#888')
        .setBackground(bg).setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true, '#e0e0e0', BS)

      sh.getRange(row, 2).setValue(col.name)
        .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
        .setFontColor('#212121').setBackground(bg).setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true, '#e0e0e0', BS)

      sh.getRange(row, 3).setValue(col.desc)
        .setFontFamily('Calibri').setFontSize(9).setFontColor('#424242')
        .setBackground(bg).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setBorder(true, true, true, true, true, true, '#e0e0e0', BS)

      sh.getRange(row, 4).setValue(col.vals)
        .setFontFamily('Calibri').setFontSize(9).setFontColor('#555')
        .setBackground(bg).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setBorder(true, true, true, true, true, true, '#e0e0e0', BS)

      sh.getRange(row, 5).setValue(col.auto)
        .setFontFamily('Calibri').setFontSize(9).setFontStyle('italic')
        .setFontColor('#1565c0').setBackground(bg).setVerticalAlignment('middle')
        .setHorizontalAlignment('center')
        .setBorder(true, true, true, true, true, true, '#e0e0e0', BS)

      sh.setRowHeight(row, 20)
      row++
    }
    row++ // gap between sections
  }

  // Info footer
  sh.getRange(row, 1, 1, 5).merge()
  sh.getRange(row, 1).setValue('Los colores de cada sección coinciden con los de la fila de encabezados en Pacientes. Tooltips disponibles al pasar el mouse sobre fila 3.')
    .setFontFamily('Calibri').setFontSize(10).setFontColor('#757575').setBackground('#f5f5f5')
    .setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  sh.setRowHeight(row, 24)

  var _mr = sh.getMaxRows()
  if (_mr > row + 3) sh.deleteRows(row + 3, _mr - row - 2)

  ss.toast('Guía de columnas creada (' + (_COLUMNAS._count || _COLUMNAS.length) + ' columnas)', '', 4)
}

// ─── CREAR DASHBOARD ──────────────────────────────────────────────────────

function crearDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var pac = ss.getSheetByName('Pacientes')
  if (!pac) { ss.toast('Hoja Pacientes no encontrada', '', 3); return }
  ss.toast('Generando Dashboard…', '', 1)

  var old = ss.getSheetByName('Dashboard')
  if (old) ss.deleteSheet(old)

  var sh = ss.insertSheet('Dashboard')
  sh.setTabColor('#1b5e20')

  var P = "'Pacientes'!"
  var SEP = getFormulaSep()
  var BS = SpreadsheetApp.BorderStyle.SOLID
  var BSM = SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  var GREEN = '#1b5e20', GREEN_L = '#e8f5e9', GREEN_M = '#2e7d5b'
  var W = '#ffffff', T = '#212121', TG = '#757575'
  var SEXO = colToLetter(7), EDAD = colToLetter(10)
  var PRIO = colToLetter(109), DEP = colToLetter(29)
  var EMPC = colToLetter(20), EMPU = colToLetter(81)

  var CARD_BORDER = '#4caf50'

  sh.setColumnWidth(1, 240)
  for (var _ci = 2; _ci <= 8; _ci++) sh.setColumnWidth(_ci, 130)
  sh.setColumnWidth(9, 100)

  function sec(t, r) {
    sh.getRange(r, 1, 1, 9).merge()
    sh.getRange(r, 1).setValue(t).setFontFamily('Calibri').setFontSize(13).setFontWeight('bold')
      .setFontColor(W).setBackground(GREEN).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, W, BSM)
    sh.setRowHeight(r, 30)
    sh.getRange(r + 1, 1).setBackground(GREEN_L)
    sh.setRowHeight(r + 1, 4)
  }
  function cardLabel(r, c, w, t, fc) {
    sh.getRange(r, c, 1, w).merge()
      .setValue(t).setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
      .setFontColor(fc || T).setBackground('#f9fafb').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, false, true, true, false, CARD_BORDER, BS)
  }
  function cardVal(r, c, w, f, fc) {
    if (f) sh.getRange(r, c).setFormula(f)
    var rng = sh.getRange(r, c, 1, w).merge()
    rng.setFontFamily('Calibri').setFontSize(22).setFontWeight('bold')
      .setFontColor(fc || GREEN).setBackground(W).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(false, true, true, true, true, true, CARD_BORDER, BS)
  }
  function cf(rng, criteria) { return '=COUNTIF(' + rng + SEP + ' "' + criteria + '")' }
  function cf2(rng, lo, hi) {
    return '=COUNTIF(' + rng + SEP + ' ">=' + lo + '")-COUNTIF(' + rng + SEP + ' ">=' + hi + '")'
  }

  var R = 1
  sh.getRange(R, 1, 1, 9).merge()
  sh.getRange(R, 1).setValue('SISTEMA PADDS — DASHBOARD')
    .setFontFamily('Calibri').setFontSize(22).setFontWeight('bold')
    .setFontColor(W).setBackground(GREEN).setHorizontalAlignment('center').setVerticalAlignment('middle')
  sh.setRowHeight(R, 52)
  sh.getRange(R, 1).setBorder(true, true, true, true, true, true, '#0d3b14', BSM)
  R++
  sh.getRange(R, 1, 1, 9).merge()
  sh.getRange(R, 1).setValue('Actualizado: ' + Utilities.formatDate(new Date(), 'America/Santiago', 'dd/MM/yyyy HH:mm') + '  •  Datos en vivo desde la hoja Pacientes')
    .setFontFamily('Calibri').setFontSize(9).setFontStyle('italic').setFontColor('#e8f5e9')
    .setBackground(GREEN_M).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(false, true, true, true, true, true, W, BS)
  sh.setRowHeight(R, 24)
  R += 2

  sec('RESUMEN GENERAL', R); R += 2
  var RC = R
  var resumen = [
    [1, 2, 'Total Pacientes', '=MAX(' + P + 'A4:A)', GREEN],
    [3, 2, 'Vigentes',        cf(P + 'F4:F', 'VIGENTE'), '#2e7d32'],
    [5, 2, 'Fallecidos',      cf(P + 'F4:F', 'FALLECIDO'), '#c62828'],
    [7, 2, 'Egresados',       cf(P + 'F4:F', 'EGRESO'), '#e65100'],
  ]
  resumen.forEach(function(x) { cardLabel(RC, x[0], x[1], x[2], x[3]); cardVal(RC + 1, x[0], x[1], x[3], x[3]) })
  sh.setRowHeight(RC, 22); sh.setRowHeight(RC + 1, 48)
  var _totalValCell = colToLetter(1) + (RC + 1)
  R = RC + 4

  sec('DEMOGRAFIA  —  edad y sexo', R); R += 2
  var RC = R
  var demo = [
    [1, 'Hombres',            cf(P + SEXO + '4:' + SEXO, 'F'), '#1565c0'],
    [3, 'Mujeres',            cf(P + SEXO + '4:' + SEXO, 'M'), '#c62828'],
    [5, 'Edad Promedio',      '=ROUND(AVERAGE(' + P + EDAD + '4:' + EDAD + ')' + SEP + ' 0)', GREEN],
    [7, 'Electrodependiente', cf(P + colToLetter(29) + '4:' + colToLetter(29), 'SI'), '#6a1b9a'],
  ]
  demo.forEach(function(x) { cardLabel(RC, x[0], 2, x[1], x[3]); cardVal(RC + 1, x[0], 2, x[2], x[3]) })
  sh.setRowHeight(RC, 22); sh.setRowHeight(RC + 1, 44)
  R = RC + 4
  var ed = P + EDAD + '4:' + EDAD
  var RA = R
  var ages1 = [
    [1, '< 20',  cf(ed, '<20')],
    [3, '20-44',  cf2(ed, 20, 45)],
    [5, '45-59',  cf2(ed, 45, 60)],
    [7, '60-64',  cf2(ed, 60, 65)],
    [9, '65-74',  cf2(ed, 65, 75)],
  ]
  ages1.forEach(function(x) { cardLabel(RA, x[0], 1, x[1], GREEN); cardVal(RA + 1, x[0], 1, x[2], GREEN) })
  sh.setRowHeight(RA, 16); sh.setRowHeight(RA + 1, 36)

  R = RA + 3
  var ages2 = [
    [1, 4, '75+',          cf(ed, '>=75')],
    [5, 4, 'Total edades', '=COUNTIF(' + ed + SEP + ' ">0")'],
  ]
  ages2.forEach(function(x) { cardLabel(R, x[0], x[1], x[2], GREEN); cardVal(R + 1, x[0], x[1], x[3], GREEN) })
  sh.setRowHeight(R, 16); sh.setRowHeight(R + 1, 36)
  R += 4

  sec('PRIORIDAD GENERAL', R); R += 2
  var RP = R
  var prio = [
    ['URGENTE',     cf(P + PRIO + '4:' + PRIO, 'URGENTE'),     '#c62828', '#ffebee'],
    ['POR REVISAR', cf(P + PRIO + '4:' + PRIO, 'POR REVISAR'), '#e65100', '#fff3e0'],
    ['AL DIA',      cf(P + PRIO + '4:' + PRIO, 'AL DIA'),      '#2e7d32', '#e8f5e9'],
    ['N/A',          cf(P + PRIO + '4:' + PRIO, 'N/A'),         TG,        '#f5f5f5'],
  ]
  prio.forEach(function(x, i) {
    var c = 1 + i * 2
    cardLabel(RP, c, 2, x[0], x[2]); cardVal(RP + 1, c, 2, x[1], x[2])
  })
  sh.setRowHeight(RP, 20); sh.setRowHeight(RP + 1, 44)
  R = RP + 4

  sec('SECTOR  —  distribucion geografica', R); R += 2
  var RS = R
  var sect = [
    ['VERDE',     cf(P + 'B4:B', 'VERDE'),     '#2e7d32', '#e8f5e9'],
    ['AMARILLO',  cf(P + 'B4:B', 'AMARILLO'),  '#f9a825', '#fffde7'],
    ['NARANJO',   cf(P + 'B4:B', 'NARANJO'),   '#e65100', '#fff3e0'],
    ['PENDIENTE', cf(P + 'B4:B', 'Pendiente'), TG,        '#f5f5f5'],
  ]
  sect.forEach(function(x, i) {
    var c = 1 + i * 2
    cardLabel(RS, c, 2, x[0], x[2]); cardVal(RS + 1, c, 2, x[1], x[2])
  })
  sh.setRowHeight(RS, 20); sh.setRowHeight(RS + 1, 44)
  R = RS + 4

  sec('DEPENDENCIA  —  clasificacion funcional', R); R += 2
  var RD = R
  var dlist = [
    ['SEVERO',        cf(P + DEP + '4:' + DEP, 'SEVERO'),        '#c62828', '#ffebee'],
    ['MODERADO',      cf(P + DEP + '4:' + DEP, 'MODERADO'),      '#e65100', '#fff3e0'],
    ['LEVE',          cf(P + DEP + '4:' + DEP, 'LEVE'),          '#f9a825', '#fffde7'],
    ['INDEPENDIENTE', cf(P + DEP + '4:' + DEP, 'INDEPENDIENTE'), '#2e7d32', '#e8f5e9'],
  ]
  dlist.forEach(function(x, i) {
    var c = 1 + i * 2
    cardLabel(RD, c, 2, x[0], x[2]); cardVal(RD + 1, c, 2, x[1], x[2])
  })
  sh.setRowHeight(RD, 20); sh.setRowHeight(RD + 1, 44)
  R = RD + 4

  sec('COBERTURA EMPA / EMPAM', R); R += 2
  var RE = R
  var empa = [
    [1, 2, 'Cuidador EMPA',  cf(P + EMPC + '4:' + EMPC, 'EMPA'),  '#1565c0'],
    [3, 2, 'Cuidador EMPAM', cf(P + EMPC + '4:' + EMPC, 'EMPAM'), '#6a1b9a'],
    [5, 2, 'Usuario EMPA',   cf(P + EMPU + '4:' + EMPU, 'EMPA'),  '#1565c0'],
    [7, 2, 'Usuario EMPAM',  cf(P + EMPU + '4:' + EMPU, 'EMPAM'), '#6a1b9a'],
  ]
  empa.forEach(function(x) { cardLabel(RE, x[0], x[1], x[2], x[4]); cardVal(RE + 1, x[0], x[1], x[3], x[4]) })
  sh.setRowHeight(RE, 20); sh.setRowHeight(RE + 1, 44)
  R = RE + 4

  sec('PATOLOGIAS CRONICAS  —  prevalencia', R); R += 2
  var pHeaders = pac.getRange(3, 37, 1, 17).getValues()[0]
  var pathRow = R

  var pathHeaders = ['Patologia', 'Pacientes', '%']
  for (var _phi = 0; _phi < pathHeaders.length; _phi++) {
    sh.getRange(pathRow, _phi + 1).setValue(pathHeaders[_phi])
      .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
      .setFontColor(W).setBackground(GREEN).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, GREEN, BS)
  }
  sh.setRowHeight(pathRow, 24)

  var pacNums = pac.getRange(4, 1, pac.getLastRow() - 3, 1).getValues().reduce(function(m, r) { return Math.max(m, Number(r[0]) || 0) }, 0)
  sh.getRange(pathRow, 4).setValue('Base: ' + pacNums + ' pac. (MAX N°)')
    .setFontFamily('Calibri').setFontSize(8).setFontStyle('italic').setFontColor('#81c784')
    .setBackground(GREEN).setVerticalAlignment('middle').setHorizontalAlignment('left')
    .setBorder(true, true, true, true, true, true, GREEN, BS)

  for (var p = 0; p < pHeaders.length; p++) {
    var pr = pathRow + 1 + p
    var cn = 37 + p
    var cl = colToLetter(cn)
    var nm = String(pHeaders[p] || '').trim() || 'Patologia ' + cn
    var bg = p % 2 === 0 ? W : '#f5faf5'

    var countFormula = 'COUNTIF(' + P + cl + '4:' + cl + ';"SI")'
    if (cn === 53) countFormula = 'COUNTA(' + P + cl + '4:' + cl + ')'

    sh.getRange(pr, 1).setValue(nm)
      .setFontFamily('Calibri').setFontSize(10).setFontColor(T).setBackground(bg).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#c8e6c9', BS)
    sh.getRange(pr, 2).setFormula('=' + countFormula)
      .setFontFamily('Calibri').setFontSize(10).setFontColor(GREEN).setBackground(bg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#c8e6c9', BS)
    sh.getRange(pr, 3).setFormula('=IFERROR(' + countFormula + '/' + _totalValCell + ';0)')
      .setNumberFormat('0%')
      .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold').setFontColor(GREEN).setBackground(bg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#c8e6c9', BS)
    sh.setRowHeight(pr, 22)
  }
  R = pathRow + 1 + pHeaders.length + 2

  sec('VIGENCIAS POR AREA', R); R += 2
  var vigHeaders = ['Control / Area', 'AL DIA', 'POR\nVENCER', 'VENCIDO', 'PENDTE', 'N/A', 'TOTAL', '% AL DIA', 'BARRAS']
  var vigBgs = [GREEN, '#2e7d32', '#e65100', '#c62828', '#f9a825', TG, GREEN, '#2e7d32', '#37474f']
  vigHeaders.forEach(function(h, i) {
    sh.getRange(R, i + 1).setValue(h)
      .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold')
      .setFontColor(W).setBackground(vigBgs[i] || GREEN_M).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, GREEN, BS)
  })
  sh.setRowHeight(R, 30)
  R++
  for (var v = 0; v < CONTROL_COLS.length; v++) {
    var vr = R + v
    var ec = CONTROL_COLS[v][2], cl2 = colToLetter(ec)
    var nm2 = CONTROL_COLS[v][0]
    var bg2 = v % 2 === 0 ? '#fafafa' : '#f0f8f0'

    sh.getRange(vr, 1).setValue(nm2)
      .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold').setFontColor(T).setBackground(bg2).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#c8e6c9', BS)
    var sts = ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A']
    sts.forEach(function(s, si) {
      sh.getRange(vr, 2 + si).setFormula(cf(P + cl2 + '4:' + cl2, s))
        .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold').setFontColor(T).setBackground(bg2)
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true, '#c8e6c9', BS)
    })
    sh.getRange(vr, 7).setFormula('=SUM(B' + vr + ':F' + vr + ')')
      .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold').setFontColor(GREEN).setBackground(bg2)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#c8e6c9', BS)
    sh.getRange(vr, 8).setFormula('=IF(G' + vr + '=0' + SEP + ' ""' + SEP + ' ROUND(B' + vr + '/G' + vr + '*100' + SEP + ' 0))')
      .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold').setFontColor('#2e7d32').setBackground(bg2)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#c8e6c9', BS)
    sh.getRange(vr, 9).setFormula('=SPARKLINE(B' + vr + ':F' + vr + ')')
      .setBackground(bg2).setVerticalAlignment('middle').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true, '#c8e6c9', BS)
    sh.setRowHeight(vr, 22)
  }
  R += CONTROL_COLS.length + 2

  sec('ALERTAS RESUMEN', R); R += 2
  var ar = R
  var stCols = CONTROL_COLS.map(function(c) { return c[2] }).filter(function(x) { return x })
  function _rcf(r, c) { return 'COUNTIF(' + r + SEP + ' "' + c + '")' }
  var alertas = [
    ['VENCIDOS',    stCols.map(function(c) { return _rcf(P + colToLetter(c) + '4:' + colToLetter(c), 'VENCIDO') }).join('+'), '#c62828', '#ffebee'],
    ['POR VENCER',  stCols.map(function(c) { return _rcf(P + colToLetter(c) + '4:' + colToLetter(c), 'POR VENCER') }).join('+'), '#e65100', '#fff3e0'],
    ['PENDIENTES',  stCols.map(function(c) { return _rcf(P + colToLetter(c) + '4:' + colToLetter(c), 'PENDIENTE') }).join('+'), '#f9a825', '#fffde7'],
  ]
  alertas.forEach(function(x, i) {
    var ac = 1 + i * 3
    sh.getRange(ar, ac, 1, 3).merge()
    sh.getRange(ar, ac).setValue(x[0])
      .setFontFamily('Calibri').setFontSize(13).setFontWeight('bold')
      .setFontColor(x[2]).setBackground(x[3]).setVerticalAlignment('middle').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true, x[2], BS)
    sh.getRange(ar + 1, ac, 1, 3).merge()
    sh.getRange(ar + 1, ac).setFormula('=' + x[1])
      .setFontFamily('Calibri').setFontSize(26).setFontWeight('bold')
      .setFontColor(x[2]).setBackground(W).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, x[2], BS)
    sh.setRowHeight(ar, 24)
    sh.setRowHeight(ar + 1, 52)
  })
  R = ar + 4

  sh.getRange(R, 1, 1, 9).merge()
  sh.getRange(R, 1).setValue('Leyenda de clasificacion de controles')
    .setFontFamily('Calibri').setFontSize(12).setFontWeight('bold').setFontColor(W).setBackground('#37474f')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, W, BSM)
  sh.setRowHeight(R, 28)
  R += 2
  var _leyColors = { bg: ['#e8f5e9', '#fff8e1', '#ffebee', '#fff9c4', '#f5f5f5'], fg: ['#2e7d32', '#e65100', '#c62828', '#f9a825', TG] }
  var stLbl = [
    ['AL DIA',       'Control vigente — fecha dentro del plazo configurado en Parametros'],
    ['POR VENCER',   'Proximo a vencer — segun DIAS_AVISO en Parametros'],
    ['VENCIDO',      'Fecha supero el plazo maximo de vigencia'],
    ['PENDIENTE',    'Sin fecha registrada o formato no valido'],
    ['N/A',          'No aplica al paciente (ej. menor de edad)'],
  ]
  for (var _li = 0; _li < stLbl.length; _li++) {
    var s = stLbl[_li]
    sh.getRange(R + _li, 1).setValue(s[0])
      .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
      .setFontColor(_leyColors.fg[_li]).setBackground(_leyColors.bg[_li]).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#c8e6c9', BS)
    sh.getRange(R + _li, 2, 1, 8).merge().setValue(s[1])
      .setFontFamily('Calibri').setFontSize(10).setFontColor('#333').setBackground(_leyColors.bg[_li])
      .setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, '#c8e6c9', BS)
    sh.setRowHeight(R + _li, 22)
  }
  R += stLbl.length + 2

  sh.getRange(R, 1).setValue('Areas de control (fecha de ultimo control → estado de vigencia):')
    .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold').setFontColor(W).setBackground('#455a64')
    .setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, GREEN, BS)
  sh.getRange(R, 2, 1, 8).merge()
  sh.setRowHeight(R, 24)
  R++
  for (var _ai2 = 0; _ai2 < CONTROL_COLS.length; _ai2++) {
    var c = CONTROL_COLS[_ai2]
    var bgA2 = _ai2 % 2 === 0 ? '#f5f7fa' : '#eef1f5'
    sh.getRange(R + _ai2, 1).setValue(c[0])
      .setFontFamily('Calibri').setFontSize(9).setFontColor(T).setBackground(bgA2).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#c8e6c9', BS)
    var clF = colToLetter(c[1]), clE = colToLetter(c[2])
    sh.getRange(R + _ai2, 2, 1, 8).merge().setValue('Fecha: columna ' + clF + '  →  Estado: columna ' + clE)
      .setFontFamily('Consolas,monospace').setFontSize(9).setFontColor('#555').setBackground(bgA2)
      .setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, '#c8e6c9', BS)
    sh.setRowHeight(R + _ai2, 20)
  }
  R += CONTROL_COLS.length + 2

  sh.getRange(R, 1, 1, 9).merge()
  sh.getRange(R, 1).setValue('Prioridad General (col 109): VENCIDO → URGENTE | POR VENCER → POR REVISAR | AL DIA/N/A → AL DIA | sin datos → N/A')
    .setFontFamily('Calibri').setFontSize(9).setFontStyle('italic').setFontColor('#555').setBackground('#f5f5f5')
    .setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#c8e6c9', BS)
  sh.setRowHeight(R, 24)
  R++

  var _mr2 = sh.getMaxRows()
  if (_mr2 > R + 2) sh.deleteRows(R + 2, _mr2 - R - 1)

  ss.toast('Dashboard creado con estadisticas en vivo', '', 4)
}

// ─── RESALTAR FILA ACTIVA (ondSelectionChange) ──────────────────────────────

function configurarResaltadoFila() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  try {
    var helper = ss.getSheetByName('_Resalte')
    if (!helper) {
      helper = ss.insertSheet('_Resalte')
      helper.hideSheet()
      helper.getRange(1, 1, 1, 2).setValues([['Hoja', 'Fila']])
      helper.getRange(1, 2).setValue(0)
    }

    var sheets = ['Pacientes', 'Formulario Usuario / Profesional']
    for (var si = 0; si < sheets.length; si++) {
      var sh = ss.getSheetByName(sheets[si])
      if (!sh) continue

      // Remove old _Resalte rules safely, then add fresh one
      var existing = []
      try {
        existing = sh.getConditionalFormatRules()
      } catch(e2) {
        existing = []
      }
      var keep = []
      for (var ri = 0; ri < existing.length; ri++) {
        try {
          var hasResalte = false
          var rngs = existing[ri].getRanges()
          for (var ri2 = 0; ri2 < rngs.length; ri2++) {
            try {
              if (rngs[ri2].getA1Notation().indexOf('_Resalte') !== -1) { hasResalte = true; break }
            } catch(e3) { hasResalte = true; break }
          }
          if (!hasResalte) keep.push(existing[ri])
        } catch(e4) { /* skip broken rule */ }
      }

      var mr = sh.getMaxRows()
      var mc = sh.getMaxColumns()
      var range = sh.getRange(1, 1, mr, mc)
      var formula = '=AND(ROW()>=4, ROW()=INDIRECT("_Resalte!B1"), INDIRECT("_Resalte!A1")="' + sheets[si] + '")'
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground('#E8F0FE')
        .setRanges([range])
        .build()
      keep.push(rule)
      sh.setConditionalFormatRules(keep)
    }
  } catch(e) {
    try { ss.toast('Error al configurar resaltado: ' + e.message, '', 5) } catch(e5) {}
  }
}



