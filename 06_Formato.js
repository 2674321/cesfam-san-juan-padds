// ════════════════════════════════════════════════════════
// 06_Formato.gs │ formato/validaciones/colores/dashboard/parametros (112 cols 'Pacientes')

// ─────────────────────────────────────────────────────────
// ─── FORMATO: VALIDACIONES, FORMATO CONDICIONAL, HOJAS DE SISTEMA ───────────

// ─── RECONSTRUIR PACIENTES ─────────────────────────────────────────────────

function _agregarColumnasFaltantes(ss, sh, targetCols) {
  var currentCols = sh.getLastColumn()
  if (currentCols >= targetCols) return currentCols

  ss.toast('Agregando columnas faltantes…', 'PADDS', 1)
  sh.insertColumns(currentCols + 1, targetCols - currentCols)

  var names = []
  for (var c = currentCols + 1; c <= targetCols; c++) {
    var colInfo = _COLUMNAS[c]
    names.push(colInfo ? colInfo.name : '')
  }
  if (names.length) sh.getRange(3, currentCols + 1, 1, names.length).setValues([names])

  return targetCols
}

function _repintarSinConfirmar(sh, ss) {
  ss.toast('Preparando reconstrucción funcional…', 'PADDS', 1)

  var _migroIns = false
  try { _migroIns = _asegurarColumnaInsulino(sh) } catch(eM) {}
  if (_migroIns) ss.toast('Columna INSULINO DEPENDIENTE agregada…', 'PADDS', 2)

  var targetCols = _COLUMNAS._count || 112
  var currentCols = _agregarColumnasFaltantes(ss, sh, targetCols)

  var data = sh.getDataRange().getValues()
  var rows = data.length
  if (rows < 3) {   ss.toast(HOJA_PAC + ' sin datos para formatear', 'Pacientes', 4); return }

  var _hdr = []
  for (var _hc = 1; _hc <= currentCols; _hc++) {
    var _ci = _COLUMNAS[_hc]
    _hdr.push(_ci && _ci.name ? _ci.name : '')
  }
  sh.getRange(3, 1, 1, currentCols).setValues([_hdr])

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
  try { recalcularTodo() } catch(e) {}
  if (_migroIns) { try { crearDashboard() } catch(eD) {} }
}

function _restaurarVisualSinConfirmar(sh, ss) {
  var lr = sh.getLastRow()
  var lc = sh.getLastColumn()
  if (lr < 3) {   ss.toast(HOJA_PAC + ' sin datos para formatear', 'Pacientes', 4); return }
  var targetCols = _COLUMNAS._count || 112
  _agregarColumnasFaltantes(ss, sh, targetCols)
  lc = sh.getLastColumn()
  _aplicarFormatoVisual(sh, lr, lc)
  try { recalcularTodo() } catch(e) {}
}

// ─── FORMATEAR DATOS: un solo paso (diseño completo) ───────────────────────

function formatearDatos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var chk = _hojaPacientesValida(ss)
  if (!chk.ok) { ss.toast(chk.msg, 'Pacientes', 6); return }
  var sh = chk.sh
  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Formatear datos de ' + HOJA_PAC,
    'Se aplicará el diseño completo de la plantilla:\n' +
    '· Validaciones (casillas, dropdowns, fechas) y tooltips\n' +
    '· Secciones, buscador, contador y dropdown de sección (F2)\n' +
    '· Colores de vigencia, bordes, fuentes y anchos de columna\n' +
    '· Filas completamente vacías (sin datos) se ELIMINAN\n\n' +
    'Los datos de pacientes no se pierden.\n¿Continuar?',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return
  _formatearDatosSinConfirmar(sh, ss)
}

function _formatearDatosSinConfirmar(sh, ss) {
  ss.toast('Formateando datos y diseño…', 'PADDS', 1)

  var eliminadas = 0
  try { eliminadas = _borrarFilasVacias(sh, 4) } catch(eV) {}
  if (eliminadas) ss.toast(eliminadas + ' filas vacías eliminadas. Formateando…', 'PADDS', 2)

  var errores = []
  try { _repintarSinConfirmar(sh, ss) }
  catch(eE1) { errores.push('Formato funcional: ' + eE1.message) }

  var corrDrop = 0
  try { corrDrop = _corregirDropdownsPacientes(sh, sh.getLastRow()) }
  catch(eC) { errores.push('Dropdowns: ' + eC.message) }
  if (corrDrop > 0) ss.toast(corrDrop + ' valor(es) de dropdown corregido(s)', 'PADDS', 3)

  try { _restaurarVisualSinConfirmar(sh, ss) }
  catch(eE2) { errores.push('Diseño visual: ' + eE2.message) }

  try { _pintarOpciones(sh, sh.getLastRow(), sh.getLastColumn()) }
  catch(eP) { errores.push('Colores de opciones: ' + eP.message) }

  try { _pintarFechasInvalidas(sh, sh.getLastRow(), sh.getLastColumn()) }
  catch(ePI) { errores.push('Fechas inválidas: ' + ePI.message) }

  try { formatearRUTPacientes(true) }
  catch(eR) { errores.push('RUN: ' + eR.message) }

  var ui = SpreadsheetApp.getUi()
  if (errores.length) {
    ui.alert('Formatear datos', '⚠️ Algunos pasos fallaron:\n- ' + errores.join('\n- '), ui.ButtonSet.OK)
  } else {
    ss.toast('Formatear datos: diseño completo aplicado a ' + HOJA_PAC +
      (corrDrop ? ' (' + corrDrop + ' valores de dropdown corregidos)' : '') +
      (eliminadas ? ' (' + eliminadas + ' filas vacías eliminadas)' : ''), 'PADDS', 4)
  }
}

// ─── REPARAR FORMATO (reubicar sin tocar datos) ──────────────────────────────

function repararFormatoPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var chk = _hojaPacientesValida(ss)
  if (!chk.ok) { ss.toast(chk.msg, 'Pacientes', 6); return }
  var sh = chk.sh
  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Reparar formato de ' + HOJA_PAC,
    'Esto REUBICA todo el formato en las columnas correctas:\n' +
    '· Elimina filas sobrantes que contengan SOLO el encabezado de columnas\n' +
    '  (tipo "N° · SECTOR · NOMBRE …") que hayan quedado como datos\n' +
    '· Limpia fondos, bordes, celdas combinadas, validaciones y reglas que\n' +
    '  vengan de la hoja anterior (por pegar con formato)\n' +
    '· Re-aplica el diseño completo (secciones, cabeceras, buscador,\n' +
    '  validaciones, colores, anchos) a las columnas correctas\n\n' +
    'LOS DATOS NO SE BORRAN ni se mueven.\n¿Continuar?',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return

  ss.toast('Reubicando formato…', 'PADDS', 1)
  var pasos = []

  try {
    var eHdr = _eliminarFilasConCabecera(sh)
    if (eHdr) pasos.push('✔ ' + eHdr + ' fila(s) con encabezado duplicado eliminada(s)')
  } catch(eH) { pasos.push('✘ encabezados: ' + eH.message) }

  try {
    var _merges = sh.getMergedRanges()
    for (var _mi = 0; _mi < _merges.length; _mi++) _merges[_mi].breakApart()
    var _maxR = sh.getMaxRows()
    var _maxC = sh.getMaxColumns()
    sh.getRange(1, 1, _maxR, _maxC).clearFormat()
    sh.getRange(1, 1, _maxR, _maxC).clearDataValidations()
    sh.setConditionalFormatRules([])
    pasos.push('✔ Formato anterior limpiado')
  } catch(eL) { pasos.push('✘ limpiar formato: ' + eL.message) }

  try { _formatearDatosSinConfirmar(sh, ss); pasos.push('✔ Diseño completo re-aplicado') }
  catch(eF) { pasos.push('✘ formatear: ' + eF.message) }

  ss.toast('Reparar formato: listo', 'PADDS', 5)
  ui.alert('Reparar formato de ' + HOJA_PAC, pasos.join('\n'), ui.ButtonSet.OK)
}

// Elimina filas de datos que contienen SOLO el encabezado de columnas (la fila
// de títulos copiada por error al transferir, ej: "N° SECTOR NOMBRE …").
function _eliminarFilasConCabecera(sh) {
  var lr = sh.getLastRow()
  if (lr < 5) return 0
  var lc = Math.min(sh.getLastColumn(), _COLUMNAS._count || 112)
  var data = sh.getRange(4, 1, lr - 3, lc).getValues()
  var aEliminar = []
  for (var r = 0; r < data.length; r++) {
    var c1 = String(data[r][0] || '').trim()
    var c2 = String(data[r][1] || '').trim()
    var c3 = String(data[r][2] || '').trim()
    if (c1 === 'N°' && c2 === 'SECTOR' && c3 === 'NOMBRE') aEliminar.push(r)
  }
  if (!aEliminar.length) return 0
  for (var i = aEliminar.length - 1; i >= 0; i--) sh.deleteRows(4 + aEliminar[i])
  return aEliminar.length
}

function aplicarFormatoFuncional(sh, lr, lc) {

  sh.setFrozenColumns(0)
  sh.setFrozenRows(0)

  if (lr >= 4) {
    var dataRows = lr - 3
    var runCols = [COL.RUN, COL.RUN_CUIDADOR]
    var runFmt = []
    for (var ri = 0; ri < runCols.length; ri++) {
      runFmt.push('R4C' + runCols[ri] + ':R' + lr + 'C' + runCols[ri])
    }
    sh.getRangeList(runFmt).setNumberFormat('@')

    var dateFmtRanges = []
    for (var fi = 0; fi < _FECHAS_VA.length; fi++) {
      var fc = _FECHAS_VA[fi]
      if (fc <= lc) {
        dateFmtRanges.push('R4C' + fc + ':R' + lr + 'C' + fc)
        sh.getRange(4, fc, dataRows, 1).setHorizontalAlignment('center')
      }
    }
    sh.getRangeList(dateFmtRanges).setNumberFormat('dd/mm/yyyy')

    var _txtCols = [COL.TELEFONO, COL.PATOLOGIAS_CUIDADOR, COL.OTRAS_PATOLOGIAS,
      COL.MORBILIDAD, COL.CONTROLES_MISCELANEOS, COL.RECETAS_CONTROLADAS,
      COL.ZONA_EVACUACION, COL.UBICACION_LPP, COL.TTO_INVASIVOS, COL.OBSERVACIONES,
      COL.P_A, COL.HEMOGLOBINA_GLICOCILADA, COL.LDL_70, COL.RAC, COL.VFG]
    var _txtRanges = []
    for (var _ti = 0; _ti < _txtCols.length; _ti++) {
      var _tc = _txtCols[_ti]
      if (_tc <= lc) _txtRanges.push('R4C' + _tc + ':R' + lr + 'C' + _tc)
    }
    if (_txtRanges.length) sh.getRangeList(_txtRanges).setNumberFormat('@')

    for (var bc = 0; bc < _CHECKBOX_COLS.length; bc++) {
      var bcol = _CHECKBOX_COLS[bc]
      if (bcol <= lc) {
        sh.getRange(4, bcol, dataRows, 1)
          .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build())
          .setHorizontalAlignment('center')
      }
    }

    sh.getRange(4, 10, dataRows, 1).setHorizontalAlignment('center')
    if (19 <= lc) sh.getRange(4, 19, dataRows, 1).setHorizontalAlignment('center')

    var _centro = [2, 6, 7, 17, 22, 26, 27, 28, 49, 59, 66, 68, 71, 72, 73, 74, 75,
      79, 81, 82, 83, 84, 93, 94, 95, 97, 99, 101, 105, 106, 107, 108, 109, 110]
    for (var _cc = 0; _cc < _centro.length; _cc++) {
      if (_centro[_cc] <= lc) sh.getRange(4, _centro[_cc], dataRows, 1).setHorizontalAlignment('center')
    }
  }

  sh.setFrozenRows(3)

  if (lr >= 3) {
    var rng = sh.getRange(3, 1, lr - 2, lc)
    var f = rng.getFilter()
    if (f) f.remove()
    rng.createFilter()
  }

  _aplicarValidaciones(sh, lr, lc)

  if (lr >= 4) {

    var _vacDef = typeof _VACUNA_COLS !== 'undefined' ? _VACUNA_COLS : [81, 82, 83, 84]
    var _vacVals = typeof _VACUNA_VALS !== 'undefined' ? _VACUNA_VALS : ['SI', 'NO', 'N/A', 'R', 'P']
    var dateVal = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build()
    for (var fi = 0; fi < _FECHAS_VA.length; fi++) {
      var fc2 = _FECHAS_VA[fi]
      if (fc2 <= lc && _vacDef.indexOf(fc2) < 0) sh.getRange(4, fc2, lr - 3, 1).setDataValidation(dateVal)
    }

    var libres = typeof PAC_LIBRES !== 'undefined' ? PAC_LIBRES : [20, 48, 50, 63, 76, 98, 100, 105, 106, 107, 108, 109, 111]
    for (var li = 0; li < libres.length; li++) {
      if (libres[li] <= lc) sh.getRange(4, libres[li], lr - 3, 1).setDataValidation(null)
    }

    for (var _vai = 0; _vai < _vacDef.length; _vai++) {
      var _vacC = _vacDef[_vai]
      if (_vacC <= lc) {
        sh.getRange(4, _vacC, lr - 3, 1)
          .setNumberFormat('@')
          .setHorizontalAlignment('center')
      }
    }
  }
}

function _aplicarFormatoVisual(sh, lr, lc) {

  sh.setFrozenColumns(0)
  sh.setFrozenRows(0)

  for (var i = 0; i < Math.min(PAC_ANCHOS.length, lc); i++)
    sh.setColumnWidth(i + 1, PAC_ANCHOS[i])

  sh.setRowHeight(1, 38)
  sh.getRange(1, 1, 1, lc)
    .setFontFamily('Arial').setFontSize(12).setFontWeight('bold')
    .setVerticalAlignment('middle').setHorizontalAlignment('center')

  var _frozenLegacy = sh.getFrozenColumns()
  if (_frozenLegacy > 0) sh.setFrozenColumns(0)
  try {
    var _row1merges = sh.getRange(1, 1, 1, lc).getMergedRanges()
    for (var _m1 = 0; _m1 < _row1merges.length; _m1++) _row1merges[_m1].breakApart()
    sh.getRange(1, 1, 1, lc).clearContent()
  } catch(eB) {}
  var _FREEZE_COLS = 5
  for (var s = 0; s < PAC_SECCIONES.length; s++) {
    var sec = PAC_SECCIONES[s]
    if (sec.ini > lc) break
    var fin = Math.min(sec.fin, lc)
    if (sec.ini <= _FREEZE_COLS && _FREEZE_COLS < fin) {

      var part1 = sh.getRange(1, sec.ini, 1, _FREEZE_COLS - sec.ini + 1)
      try { part1.merge() } catch(eM1) {}
      part1.setBackground(sec.bg).setFontColor(sec.fg)
        .setValue(sec.nombre).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
        .setBorder(false, false, true, false, false, false, sec.bg2 || sec.bg, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
      var part2 = sh.getRange(1, _FREEZE_COLS + 1, 1, fin - _FREEZE_COLS)
      try { part2.merge() } catch(eM2) {}
      part2.setBackground(sec.bg).setFontColor(sec.fg).setValue('')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
        .setBorder(false, false, true, false, false, false, sec.bg2 || sec.bg, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
      try {
        part1.setNote(sec.nombre + ' · columnas ' + sec.ini + ' a ' + fin)
        part2.setNote(sec.nombre + ' · columnas ' + sec.ini + ' a ' + fin)
      } catch(eN) {}
      continue
    }
    var secRng = sh.getRange(1, sec.ini, 1, fin - sec.ini + 1)
    try { secRng.merge() } catch(eM) {}
    secRng.setBackground(sec.bg).setFontColor(sec.fg)
      .setValue(sec.nombre).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
      .setBorder(false, false, true, false, false, false, sec.bg2 || sec.bg, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    try {
      secRng.setNote(sec.nombre === 'GESTIÓN'
        ? 'Colores de vigencia de las celdas de fecha:\n' +
          '🟢 AL DIA · 🟠 POR VENCER · 🔴 VENCIDO · 🟡 PENDIENTE · ⚪ N/A'
        : sec.nombre === 'CONTROLES Y SEGUIMIENTO'
          ? 'SERVICIO CSV — controles de salud vigente del usuario (Exámenes 51 · Control Médico 52 · CCV Médico 53 · CSCV Enfermería 54) y del cuidador (EMPA/EMPAM 23 · Exámenes 24 · CCV 25).\nVigencia por fecha: 🟢 AL DIA · 🟠 POR VENCER · 🔴 VENCIDO · 🟡 PENDIENTE · ⚪ N/A'
        : sec.nombre === 'CONTROL DE SIGNOS VITALES (CSV)'
          ? 'Últimos valores de signos vitales y laboratorio: P/A (105) · Hemoglobina Glicosilada (106) · LDL menor a 70 (107) · RAC (108) · VFG (109).\nTexto libre: escribe el valor (ej: 120/80, 6.5, 70).'
          : sec.nombre + ' · columnas ' + sec.ini + ' a ' + fin)
    } catch(eN) {}
  }

  var _freezeCols = lc >= _FREEZE_COLS ? _FREEZE_COLS : lc
  try { sh.setFrozenColumns(_freezeCols) } catch(eF) {}
  // Cabecera de secciones + buscador + títulos de columna siempre visibles.
  try { sh.setFrozenRows(3) } catch(eFR) {}

    sh.setRowHeight(2, 30)
  sh.getRange(2, 1, 1, lc).setBackground('#F0FDFA')
    .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, _UI.accentL, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  for (var s2 = 0; s2 < PAC_SECCIONES.length; s2++) {
    var sec2 = PAC_SECCIONES[s2]
    if (sec2.ini > lc) break
    sh.getRange(2, sec2.ini, 1, Math.min(sec2.fin, lc) - sec2.ini + 1)
      .setBackground(_lightenHex(sec2.bg2 || sec2.bg, 180, 150, 130))
  }
  sh.getRange(2, 1).setValue('🔍').setFontSize(14).setHorizontalAlignment('center')
    .setNote('Buscador vivo: escribe el nombre, RUN o teléfono en B2 y presiona Enter (las filas que no coinciden se ocultan temporalmente).\n\nSecciones: en F2 elige qué sección mostrar ("TODAS" restaura). D2 muestra cuántos pacientes coinciden.')

  var nf = SpreadsheetApp.newTextStyle().setFontFamily('Arial').setFontSize(11)
    .setForegroundColor('#0F766E').build()

  sh.getRange(2, 2, 1, 2).merge()
  sh.getRange(2, 2, 1, 2).setTextStyle(nf).setHorizontalAlignment('center')
  sh.getRange('D2').setFontWeight('bold').setFontSize(11).setFontColor('#115E59')
    .setHorizontalAlignment('center').setBackground('#CCFBF1')
    .setNote('Resultados del buscador: "X de Y pacientes". Límite: 50 resultados en la ventana de búsqueda.')

  sh.getRange('E2').setValue('Mostrar ▾').setFontStyle('italic')
    .setFontColor('#64748B').setFontSize(12).setHorizontalAlignment('center')
    .setNote('Elige en F2 (una celda) qué sección mostrar temporalmente. Las demás columnas se ocultan; "TODAS" restaura. El buscador (B2), el filtro de estado (G2) y esta fila siempre quedan visibles.')
  var listaF = ['TODAS']
  for (var _lf = 0; _lf < PAC_SECCIONES.length; _lf++) {
    if (PAC_SECCIONES[_lf].ini <= lc) listaF.push(PAC_SECCIONES[_lf].nombre)
  }
  sh.getRange('F2').setValue('TODAS')
  sh.getRange('F2').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(listaF, true)
      .setHelpText('Elige qué sección mostrar. Las demás columnas quedan ocultas temporalmente. "TODAS" restaura todo.')
      .build())
    .setFontWeight('bold').setFontColor('#1E293B').setFontSize(11).setHorizontalAlignment('center')
    .setNote('Mostrar sección: elige aquí qué sección ver (las demás columnas se ocultan temporalmente). "TODAS" restaura la vista completa. El buscador (B2) y esta fila siempre quedan visibles.')

  sh.getRange('G2').setValue('TODOS')
  sh.getRange('G2').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['TODOS'].concat(PAC_VALIDACIONES[6]), true)
      .setHelpText('Filtra los pacientes por estado de vigencia. "TODOS" restaura. Se combina con el buscador (B2).')
      .build())
    .setFontWeight('bold').setFontColor('#1E293B').setFontSize(11).setHorizontalAlignment('center')
    .setNote('Filtro por estado: elige aquí qué estado ver ("TODOS" restaura). Se combina con el buscador de B2 y el contador D2 se actualiza solo.')

  sh.setRowHeight(3, 30)
  sh.getRange(3, 1, 1, lc).setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
    .setFontColor('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setBorder(true, true, true, true, true, true)
  for (var s = 0; s < PAC_SECCIONES.length; s++) {
    var sec = PAC_SECCIONES[s]
    if (sec.ini > lc) break
    var fin = Math.min(sec.fin, lc)
    sh.getRange(3, sec.ini, 1, fin - sec.ini + 1)
      .setBackground(sec.bg2 || sec.bg)
      .setBorder(false, false, true, false, false, false, sec.bg2 || sec.bg, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  }

  if (lr >= 4) {
    var dataRows = lr - 3
    var bgs = []
    for (var r = 0; r < dataRows; r++) {
      var bgRow = []
      var alt = _UI.zebraBg[r % 2]
      for (var c = 0; c < lc; c++) bgRow.push(alt)
      bgs.push(bgRow)
    }
    sh.getRange(4, 1, dataRows, lc).setBackgrounds(bgs)
      .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, _UI.border, SpreadsheetApp.BorderStyle.SOLID)

    var _overflowCols = [11, 20, 55, 62, 64, 77]
    for (var _wc = 0; _wc < _overflowCols.length; _wc++) {
      if (_overflowCols[_wc] <= lc) sh.getRange(4, _overflowCols[_wc], dataRows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
    }
    var _clipCols = [48, 63, 111]
    for (var _cc = 0; _cc < _clipCols.length; _cc++) {
      if (_clipCols[_cc] <= lc) sh.getRange(4, _clipCols[_cc], dataRows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    }

    if (50 <= lc) sh.getRange(4, 50, dataRows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    sh.setRowHeights(4, dataRows, 26)
  }

  _ajustarAnchosAlContenido(sh, lr, lc)

  for (var sD = 0; sD < PAC_SECCIONES.length; sD++) {
    var secD = PAC_SECCIONES[sD]
    if (secD.fin >= lc) break
    sh.getRange(2, secD.fin, lr - 1, 1).setBorder(
      null, null, null, true, null, null,
      secD.bg2 || secD.bg, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  }
  try { sh.setGridLines(false) } catch(eG) {}

  try { _refrescarFormatoCondicional(sh, lr, lc) } catch(e) {}

  try { sh.showColumns(1, sh.getMaxColumns()) } catch(eS) {}
  try { aplicarFiltroSecciones(sh) } catch(eF2) {}
}

// ─── VALIDACIONES ──────────────────────────────────────────────────────────

function _anchoTextoPx(txt, bold) {
  var s = String(txt == null ? '' : txt)
  if (!s) return 0
  var line = s.split('\n')[0]
  if (line.length > 50) line = line.slice(0, 50)
  var w = 0
  for (var i = 0; i < line.length; i++) {
    var ch = line.charAt(i)
    if (/[0-9]/.test(ch)) w += 6
    else if (ch === ' ') w += 4
    else if (/[.,:;'()/\\\-]/.test(ch)) w += 5
    else if (/[a-záéíóúñü]/.test(ch)) w += 8
    else if (/[A-ZÁÉÍÓÚÑÜ]/.test(ch)) w += 10.5
    else w += 9
  }
  w = Math.min(Math.round(w), 450)
  return bold ? Math.round(w * 1.08) : w
}

// (primera línea, con tope). Solo crece, nunca encoge lo ya ajustado.
function _ajustarAnchosAlContenido(sh, lr, lc) {
  if (!sh || lc < 1) return
  var heads = sh.getRange(3, 1, 1, lc).getValues()[0]
  var data = lr >= 4 ? sh.getRange(4, 1, lr - 3, lc).getValues() : []
  var crece = {}
  for (var c = 1; c <= lc; c++) {
    var base = c <= PAC_ANCHOS.length ? PAC_ANCHOS[c - 1] : 100
    var need = base
    var hw = _anchoTextoPx(heads[c - 1], true) + 12
    if (hw > need) need = hw
    for (var r = 0; r < data.length; r++) {
      var v = data[r][c - 1]
      var cw = (typeof v === 'object' && v instanceof Date && !isNaN(v.getTime()))
        ? 62
        : (_anchoTextoPx(v, false) + 8)
      if (cw > need) need = cw
    }
    crece[c] = need
  }
  for (var c = 1; c <= lc; c++) {
    var cur = sh.getColumnWidth(c)
    if (crece[c] > cur) sh.setColumnWidth(c, crece[c])
  }
}

// ─── DROPDOWNS CON COLOR POR OPCIÓN ─────────────────────────────────────────
//
// La API de Sheets no permite colores por opción en un dropdown directo
// (ONE_OF_LIST). La solución 100% Apps Script: "Dropdown (desde un rango)"
// (ONE_OF_RANGE) apuntando a celdas coloreadas de la hoja oculta _Opciones.
// Google Sheets muestra cada opción con el color de su celda fuente y el chip
// del valor seleccionado toma ese color. Si algo falla, se cae a
// requireValueInList (dropdown simple, mismo comportamiento actual).

function _fgSobre(bg) {
  var h = String(bg || '').replace('#', '')
  if (h.length !== 6) return '#1a1a1a'
  var r = parseInt(h.substr(0, 2), 16)
  var g = parseInt(h.substr(2, 2), 16)
  var b = parseInt(h.substr(4, 2), 16)
  return (0.299 * r + 0.587 * g + 0.114 * b) > 150 ? '#1a1a1a' : '#ffffff'
}

function _claveColorOpcion(o) {
  return _quitarAcentos(String(o == null ? '' : o)).trim().replace(/\s+/g, ' ').toUpperCase()
}

// Índice (posición en PAC_VALIDACIONES) de cada opción, por columna. Se calcula
// una sola vez (caché) y evita escanear la lista por cada celda al pintar.
var _OPC_IDX_CACHE = {}
function _idxOpcionCol(col) {
  if (!_OPC_IDX_CACHE[col]) {
    var m = {}
    var lista = PAC_VALIDACIONES[col] || []
    for (var i = 0; i < lista.length; i++) m[_claveColorOpcion(lista[i])] = i
    _OPC_IDX_CACHE[col] = m
  }
  return _OPC_IDX_CACHE[col]
}

// Fondo semántico de una opción: color propio si existe en _OPC_BG; si no,
// el de la paleta según su posición en la lista de la columna (así TODA opción
// de dropdown tiene color). null solo si el valor no pertenece a la columna.
function _bgOpcionCol(col, clave) {
  var porCol = _OPC_BG_COL[col]
  if (porCol && porCol[clave]) return porCol[clave]
  var bg = _OPC_BG[clave]
  if (bg) return bg
  var idx = _idxOpcionCol(col)[clave]
  if (idx === undefined) return null
  return _PALETA_OPCIONES[idx % _PALETA_OPCIONES.length]
}

// Color para el chip del dropdown (hoja _Opciones): [fondo, letra].
function _colorOpcion(col, opcion, idx) {
  var bg = _bgOpcionCol(col, _claveColorOpcion(opcion))
  if (!bg) bg = _PALETA_OPCIONES[idx % _PALETA_OPCIONES.length]
  return [bg, _fgSobre(bg)]
}

// ─── COLOR POR OPCIÓN EN LA CELDA (siempre visible) ─────────────────────────
//
// Además del dropdown "desde un rango" (chips de color), cada celda se pinta
// con un tinte claro del color de su opción + texto en ese color. Así el color
// según la opción seleccionada se ve siempre, aunque el chip no lo muestre.

function _tintarColor(hex, t) {
  var h = String(hex || '').replace('#', '')
  if (h.length !== 6) return '#ffffff'
  var r = parseInt(h.substr(0, 2), 16)
  var g = parseInt(h.substr(2, 2), 16)
  var b = parseInt(h.substr(4, 2), 16)
  r = Math.round(r + (255 - r) * t)
  g = Math.round(g + (255 - g) * t)
  b = Math.round(b + (255 - b) * t)
  return '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1)
}

// Devuelve { bg, fg } para una opción conocida de la columna, o null si el valor
// no pertenece a ninguna opción del dropdown (así TODA opción válida tiene color).
// Un valor con comas (multiselección, ej. "SILLA RUEDAS, CAMA") usa estilo azul.
function _colorCeldaOpcion(valor, col) {
  if (String(valor).indexOf(',') >= 0) return { bg: '#F0FDFA', fg: '#0F766E' }
  var bg = _bgOpcionCol(col, _claveColorOpcion(valor))
  if (!bg) return null
  return { bg: _tintarColor(bg, 0.75), fg: bg }
}

// Recolorea UNA celda de dropdown según su valor (usado en onEdit).
// Las celdas vacías vuelven al color de cebra de su fila.
function _pintarCeldaOpcion(sh, row, col) {
  if (_CHECKBOX_COLS.indexOf(col) >= 0) return
  if (col === COL.SECTOR || col === COL.VITAL) return
  var v = String(sh.getRange(row, col).getValue() || '').trim()
  var zebra = _UI.zebraBg[row % 2]
  if (!v) { sh.getRange(row, col).setBackground(zebra).setFontWeight('normal'); return }
  var cc = _colorCeldaOpcion(v, col)
  if (!cc) return
  var bold = PAC_MULTISELECT.indexOf(col) >= 0 && v.indexOf(',') >= 0
  sh.getRange(row, col).setBackground(cc.bg).setFontColor(cc.fg).setFontWeight(bold ? 'bold' : 'normal')
}

// Resalta ligeramente una celda de RUN: rojo suave si el dígito verificador
// no coincide; con nota nula vuelve al color de fila (cebra).
function _pintarRUT(sh, row, col, nota) {
  var zebra = _UI.zebraBg[row % 2]
  if (nota) sh.getRange(row, col).setBackground('#FEE2E2').setFontColor('#B91C1C')
  else sh.getRange(row, col).setBackground(zebra).setFontColor('#000000')
}

// Pinta todas las columnas con dropdown de una vez (al formatear). Solo se
// colorean las opciones conocidas; el resto conserva su fondo actual.
function _pintarOpciones(sh, lr, lc) {
  if (lr < 4) return
  var cols = Object.keys(PAC_VALIDACIONES).map(Number).filter(function(c) {
    return c <= lc && _CHECKBOX_COLS.indexOf(c) < 0 && c !== COL.SECTOR && c !== COL.VITAL
  })
  for (var i = 0; i < cols.length; i++) {
    var col = cols[i]
    var dataRows = lr - 3
    var rng = sh.getRange(4, col, dataRows, 1)
    var vals = rng.getValues()
    var actualBgs = rng.getBackgrounds()
    var actualFgs = rng.getFontColors()
    var bgs = [], fgs = [], fws = []
    for (var r = 0; r < vals.length; r++) {
      var vOpc = String(vals[r][0] || '').trim()
      var cc = _colorCeldaOpcion(vOpc, col)
      bgs.push([cc ? cc.bg : actualBgs[r][0]])
      fgs.push([cc ? cc.fg : actualFgs[r][0]])
      fws.push([(cc && PAC_MULTISELECT.indexOf(col) >= 0 && vOpc.indexOf(',') >= 0) ? 'bold' : 'normal'])
    }
    rng.setBackgrounds(bgs)
    rng.setFontColors(fgs)
    rng.setFontWeights(fws)
  }
}

// Versión por lote de _pintarCeldaOpcion para UNA fila completa (usada al crear
// un paciente nuevo, evita ~100 llamadas por fila). Lee los estilos actuales
// una sola vez y reescribe solo las celdas de dropdown que cambian.
function _pintarOpcionesFila(sh, row, lc) {
  if (row < 4 || lc < 1) return
  var rng = sh.getRange(row, 1, 1, lc)
  var vals = rng.getValues()[0]
  var bgs = rng.getBackgrounds()[0]
  var fgs = rng.getFontColors()[0]
  var wts = rng.getFontWeights()[0]
  var zebra = _UI.zebraBg[row % 2]
  for (var _pof = 0; _pof < lc; _pof++) {
    var col = _pof + 1
    if (_CHECKBOX_COLS.indexOf(col) >= 0) continue
    if (col === COL.SECTOR || col === COL.VITAL) continue
    var vO = String(vals[_pof] == null ? '' : vals[_pof]).trim()
    if (!vO) { bgs[_pof] = zebra; wts[_pof] = 'normal'; continue }
    var cc2 = _colorCeldaOpcion(vO, col)
    if (!cc2) continue
    bgs[_pof] = cc2.bg
    fgs[_pof] = cc2.fg
    wts[_pof] = (PAC_MULTISELECT.indexOf(col) >= 0 && vO.indexOf(',') >= 0) ? 'bold' : 'normal'
  }
  rng.setBackgrounds([bgs]).setFontColors([fgs]).setFontWeights([wts])
}

// ─── DETECTOR DE FECHAS MAL ESCRITAS / FUTURAS ───────────────────────────────

// Marca celdas de columnas de fecha con texto que no es fecha (ej. "PENDIENTE",
// "ENTREGADO", "X") y fechas futuras en columnas de control (año mal escrito).
// Las celdas correctas no se tocan. soloFila: solo esa fila (usado en onEdit).
function _pintarFechasInvalidas(sh, lr, lc, soloFila) {
  if (lr < 4) return
  var r0 = soloFila || 4
  var nr = lr - r0 + 1
  var hoy = new Date()
  hoy.setHours(0, 0, 0, 0)
  var hoyT = hoy.getTime()
  var NO_FUTURA = {}
  for (var nfi = 0; nfi < _CONTROL_FECHAS.length; nfi++) {
    var nfc = _CONTROL_FECHAS[nfi][1]
    if (nfc !== COL.F_PROXIMA_CURACION) NO_FUTURA[nfc] = true
  }
  // Fechas de nacimiento, egreso y fallecimiento tampoco pueden ser futuras.
  NO_FUTURA[COL.F_NACIMIENTO] = true
  NO_FUTURA[COL.F_EGRESO] = true
  NO_FUTURA[COL.F_FALLECIMIENTO] = true
  for (var i = 0; i < _FECHAS_VA.length; i++) {
    var col = _FECHAS_VA[i]
    if (col > lc) continue
    var rng = sh.getRange(r0, col, nr, 1)
    var vals = rng.getValues()
    var bgs = rng.getBackgrounds()
    var fgs = rng.getFontColors()
    var wts = rng.getFontWeights()
    var nb = [], nf = [], nw = [], notas = [], toca = false
    for (var r = 0; r < nr; r++) {
      var v = vals[r][0]
      var txt = v == null ? '' : String(v).trim()
      var t = txt.toUpperCase()
      if (txt === '' || t === 'N/A' || t === 'NA') {
        nb.push(null); nf.push(null); nw.push(null); notas.push('')
        continue
      }
      var f = _parseDate(v)
      if (!f && !_parseDate(txt.replace(/-/g, '/'))) {
        toca = true
        nb.push('#FEE2E2'); nf.push('#B91C1C'); nw.push('bold')
        notas.push('⚠️ Este texto no es una fecha: escribe dd/mm/aaaa o N/A')
        continue
      }
      if (NO_FUTURA[col] && _parseDate(v).getTime() > hoyT) {
        toca = true
        nb.push('#FEF3C7'); nf.push('#A16207'); nw.push('normal')
        notas.push('⚠️ Fecha futura: debe ser una fecha ya realizada (¿año mal escrito?)')
        continue
      }
      nb.push(null); nf.push(null); nw.push(null); notas.push('')
    }
    if (!toca) continue
    for (var r2 = 0; r2 < nr; r2++) {
      if (nb[r2] !== null) { bgs[r2][0] = nb[r2]; fgs[r2][0] = nf[r2]; wts[r2][0] = nw[r2] }
    }
    rng.setBackgrounds(bgs).setFontColors(fgs).setFontWeights(wts)
    rng.setNotes([notas])
  }
}

// Mantiene la hoja oculta _Opciones: una columna por dropdown, cada opción con
// fondo/letra coloreados. Devuelve { col: { letra, vals } } por columna.
function _sincronizarOpciones(ss) {
  var hoja = ss.getSheetByName('_Opciones')
  if (!hoja) {
    hoja = ss.insertSheet('_Opciones')
    try { hoja.hideSheet() } catch (eHid) {}
    try { hoja.setTabColor('#DDDDDD') } catch (eTab) {}
  }
  hoja.clear()

  var cols = Object.keys(PAC_VALIDACIONES).map(Number).sort(function(a, b) { return a - b })
  var refs = {}
  for (var i = 0; i < cols.length; i++) {
    var c = cols[i]
    var lista = PAC_VALIDACIONES[c]
    if (!lista || !lista.length || lista.length > 50) continue
    var letra = colToLetter(i + 1)
    var filas = lista.length + 1
    var rng = hoja.getRange(1, i + 1, filas, 1)

    var datos = [['Col ' + c + (typeof _COLUMNAS !== 'undefined' && _COLUMNAS[c] ? ' · ' + _COLUMNAS[c].name : '')]]
    for (var o = 0; o < lista.length; o++) datos.push([String(lista[o])])
    rng.setValues(datos)

    var bgs = [[_UI.hdrBg]], fgs = [['#ffffff']]
    for (var o2 = 0; o2 < lista.length; o2++) {
      var cc = _colorOpcion(c, lista[o2], o2)
      bgs.push([cc[0]])
      fgs.push([cc[1]])
    }
    rng.setBackgrounds(bgs).setFontColors(fgs).setFontSize(10).setVerticalAlignment('middle')
    hoja.getRange(1, i + 1).setFontWeight('bold')

    refs[c] = { letra: letra, vals: lista.slice(0) }
  }
  return refs
}

function _aplicarValidaciones(sh, lr, lc) {
  var cols = Object.keys(PAC_VALIDACIONES).map(Number).filter(function(c) {
    return c <= lc
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

  var refs = null
  try { refs = _sincronizarOpciones(sh.getParent()) } catch (eOpt) {}

  if (lr >= 4) {
    var rowCount = lr - 3
    for (var gi = 0; gi < groups.length; gi++) {
      var g = groups[gi]
      var rngSrc = refs ? refs[g[0]] : null
      var dv
      if (rngSrc && rngSrc.vals.length <= 50) {
        var rr = sh.getParent().getRange("'_Opciones'!" + rngSrc.letra + '2:' + rngSrc.letra + (rngSrc.vals.length + 1))
        dv = SpreadsheetApp.newDataValidation().requireValueInRange(rr).setAllowInvalid(true).build()
      } else {
        dv = SpreadsheetApp.newDataValidation()
          .requireValueInList(PAC_VALIDACIONES[g[0]], true).setAllowInvalid(true).build()
      }
      sh.getRange(4, g[0], rowCount, g.length).setDataValidation(dv)
    }
  }
}

// ─── CORRECCIÓN DE DROPDOWNS (al formatear) ─────────────────────────────────

// Devuelve cuántas celdas cambiaron. Nunca borra contenido.

var _DROP_VARIANTES = {
  'NA': 'N/A', 'N.A': 'N/A', 'NA.': 'N/A', 'NA/': 'N/A',
  'PROB DOMICILIO': 'PROB. DOMICILIO', 'PROB.DOMICILIO': 'PROB. DOMICILIO',
  'PROB DOM': 'PROB. DOMICILIO', 'PROB': 'PROB. DOMICILIO',
  'SOBRE PESO': 'SOBREPESO', 'SOBREPESO.': 'SOBREPESO',
  'BAJO PESO.': 'BAJO PESO', 'SIN SOBRECARGA.': 'SIN SOBRECARGA',
  'SOBRECARGA LEVE.': 'SOBRECARGA LEVE', 'SOBRECARGA INTENSA.': 'SOBRECARGA INTENSA',
  'CUIDADORA REMUNERADA.': 'CUIDADORA REMUNERADA', 'CUID REMUNERADA': 'CUIDADORA REMUNERADA',
  'EN ESPERA DE LISTA': 'EN ESPERA', 'EN ESPERA.': 'EN ESPERA',
  'NO INGRESO': 'NO INGRESA', 'NO INGRESA.': 'NO INGRESA',
  'SILLA DE RUEDAS': 'SILLA RUEDAS', 'SILLA DE RUEDA': 'SILLA RUEDAS',
  'PEND.': 'PENDIENTE',
  'COLCHON ANTI ESCARAS': 'COLCHON ANTIESCARAS', 'COLCHON ANTI ESCARA': 'COLCHON ANTIESCARAS',
  'COJIN ANTIESCARAS': 'COJIN + COLCHON', 'COJIN COLCHON': 'COJIN + COLCHON',
  'RECHAZADO': 'RECHAZA',
}

function _normalizarValorDropdown(col, v) {
  if (v == null || typeof v !== 'string') return v
  var s = v.trim().replace(/\s+/g, ' ')
  if (!s) return ''
  var lista = PAC_VALIDACIONES[col]
  if (!lista || !lista.length) return s === v ? v : s
  var up = _quitarAcentos(s).toUpperCase()
  for (var i = 0; i < lista.length; i++) {
    var it = String(lista[i])
    if (_quitarAcentos(it).toUpperCase() === up) return it
  }

  var upLimpio = up.replace(/[.,;:\-)]+$/g, '')
  if (_DROP_VARIANTES[upLimpio]) return _DROP_VARIANTES[upLimpio]
  var hits = 0, cand = null
  for (var j = 0; j < lista.length; j++) {
    var it2 = String(lista[j])
    var up2 = _quitarAcentos(it2).toUpperCase()
    if (upLimpio.length >= 3 && up2.indexOf(upLimpio) === 0) { hits++; if (!cand) cand = it2 }
  }
  if (hits === 1) return cand
  return s
}

function _corregirDropdownsPacientes(sh, lr) {
  var lc = sh.getLastColumn()
  if (lr < 4 || lc < 1) return 0
  var cols = Object.keys(PAC_VALIDACIONES).map(Number).filter(function(c) { return c <= lc })
  var corregidas = 0
  for (var ci = 0; ci < cols.length; ci++) {
    var col = cols[ci]
    var data = sh.getRange(4, col, lr - 3, 1).getValues()
    var dirty = false
    for (var r = 0; r < data.length; r++) {
      var v = data[r][0]
      if (v == null || typeof v !== 'string' || String(v).trim() === '') continue
      var n
      if (col === 6) n = _normalizarVitalEstado(v)
      else if (col === 7 || col === 17) n = _normalizarSexo(v) || String(v).trim().replace(/\s+/g, ' ')
      else if (PAC_MULTISELECT.indexOf(col) >= 0 && String(v).indexOf(',') >= 0) {
        var _parts = String(v).split(',')
        var _normParts = []
        for (var _pp = 0; _pp < _parts.length; _pp++) {
          _normParts.push(_normalizarValorDropdown(col, _parts[_pp]))
        }
        n = _normParts.join(', ')
      }
      else n = _normalizarValorDropdown(col, v)
      if (n !== v) { data[r][0] = n; dirty = true; corregidas++ }
    }
    if (dirty) sh.getRange(4, col, lr - 3, 1).setValues(data)
  }
  return corregidas
}

// ─── FORMATO CONDICIONAL ──────────────────────────────────────────────────

function _refrescarFormatoCondicional(sh, lr, lc) {
  if (lr < 4) return
  _aplicarValidaciones(sh, lr, lc)

  function _rCF(f, rng, o) {
    var r = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f)
    if (o.bg) r.setBackground(o.bg)
    if (o.fg) r.setFontColor(o.fg)
    if (o.bold !== undefined) r.setBold(o.bold)
    if (o.strike !== undefined) r.setStrikethrough(o.strike)
    return r.setRanges([rng]).build()
  }
  function _tCF(t, rng, o) {
    var r = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t)
    if (o.bg) r.setBackground(o.bg)
    if (o.fg) r.setFontColor(o.fg)
    if (o.bold !== undefined) r.setBold(o.bold)
    return r.setRanges([rng]).build()
  }

  var rules = []
  var vitalRng = sh.getRange(4, 1, lr - 3, lc)
  var _vKeys = ['FALLECIDO', 'SUSPENDIDO', 'EGRESO', 'EGRESO POR ALTA', 'ALTA', 'TRASLADO']
  for (var _vk = 0; _vk < _vKeys.length; _vk++) {
    var _kv = _vKeys[_vk], _vs = _VITAL_ROW_COLORS[_kv] || {}
    rules.push(_rCF('=UPPER($F4)="' + _kv + '"', vitalRng, { fg: _vs.fg || '#000000', strike: !!_vs.strike }))
  }
  rules.push(_rCF('=UPPER($F4)="VIGENTE"', vitalRng, { fg: '#000000', bold: false, strike: false }))

  var sectorRng = sh.getRange(4, COL.SECTOR, lr - 3, 1)
  var _sKeys = ['VERDE', 'AMARILLO', 'NARANJO', 'PENDIENTE']
  for (var _sk = 0; _sk < _sKeys.length; _sk++) {
    var _scM = _SECTOR_COLORS[_sKeys[_sk]]
    if (_scM) rules.push(_rCF('=UPPER($B4)="' + _sKeys[_sk] + '"', sectorRng, { bg: _scM[0], fg: _scM[1] }))
  }

  var prioridadRng = sh.getRange(4, COL.PRIORIDAD, lr - 3, 1)
  var _pf = _ESTADO_FECHA_COLORS
  rules.push(_tCF('URGENTE', prioridadRng, { bg: _pf['VENCIDO'][0], fg: _pf['VENCIDO'][1], bold: true }))
  rules.push(_tCF('POR REVISAR', prioridadRng, { bg: _pf['POR VENCER'][0], fg: _pf['POR VENCER'][1], bold: true }))
  rules.push(_tCF('AL DIA', prioridadRng, { bg: _pf['AL DIA'][0], fg: _pf['AL DIA'][1] }))
  rules.push(_tCF('N/A', prioridadRng, { bg: _pf['N/A'][0], fg: _pf['N/A'][1] }))

  function _porCol(colores, col, rng) {
    for (var k in colores) {
      if (colores.hasOwnProperty(k)) rules.push(_rCF('=UPPER($' + colToLetter(col) + '4)="' + k + '"', rng, { bg: colores[k][0], fg: colores[k][1] }))
    }
  }

  var _sexoColores = { 'F': ['#FCE7F3', '#BE185D'], 'M': ['#E0F2FE', '#0369A1'] }
  for (var _sx = 0; _sx < 2; _sx++) {
    var _sxc = _sx === 0 ? COL.SEXO : COL.SEXO_CUIDADOR
    if (_sxc <= lc) _porCol(_sexoColores, _sxc, sh.getRange(4, _sxc, lr - 3, 1))
  }

  var _estadoColores = {
    'VIGENTE': _ESTADO_COLORS.VIGENTE, 'FALLECIDO': _ESTADO_COLORS.FALLECIDO,
    'EGRESO': _ESTADO_COLORS.EGRESO, 'EGRESO POR ALTA': _ESTADO_COLORS['EGRESO POR ALTA'],
    'SUSPENDIDO': _ESTADO_COLORS.SUSPENDIDO, 'ALTA': _ESTADO_COLORS.ALTA,
    'TRASLADO': _ESTADO_COLORS.TRASLADO, 'PENDIENTE': _ESTADO_COLORS.PENDIENTE,
  }
  _porCol(_estadoColores, COL.VITAL, sh.getRange(4, COL.VITAL, lr - 3, 1))

  var _vacColores = { 'SI': ['#DCFCE7', '#15803D'], 'R': ['#DCFCE7', '#15803D'], 'NO': ['#FEE2E2', '#B91C1C'], 'P': ['#FEF3C7', '#B45309'], 'N/A': ['#F1F5F9', '#94A3B8'] }
  for (var _vci = 0; _vci < _VACUNA_COLS.length; _vci++) {
    var _vcc = _VACUNA_COLS[_vci]
    if (_vcc <= lc) _porCol(_vacColores, _vcc, sh.getRange(4, _vcc, lr - 3, 1))
  }

  if (COL.ESTADO_NUTRICIONAL <= lc) _porCol({
    'NORMAL': ['#DCFCE7', '#15803D'], 'SOBREPESO': ['#FFEDD5', '#C2410C'],
    'OBESIDAD': ['#FEE2E2', '#B91C1C'], 'BAJO PESO': ['#FEF3C7', '#B45309'],
    'N/A': ['#EEF1F5', '#64748B'],
  }, COL.ESTADO_NUTRICIONAL, sh.getRange(4, COL.ESTADO_NUTRICIONAL, lr - 3, 1))

  if (COL.RESULTADO_ZARIT <= lc) _porCol({
    'SIN SOBRECARGA': ['#DCFCE7', '#15803D'], 'SOBRECARGA LEVE': ['#FFEDD5', '#C2410C'],
    'SOBRECARGA INTENSA': ['#FEE2E2', '#B91C1C'], 'N/A': ['#EEF1F5', '#64748B'],
  }, COL.RESULTADO_ZARIT, sh.getRange(4, COL.RESULTADO_ZARIT, lr - 3, 1))

  for (var _di = 0; _di < _FECHAS_VA.length; _di++) {
    var _dc = _FECHAS_VA[_di]
    if (_dc <= lc) rules.push(_tCF('N/A', sh.getRange(4, _dc, lr - 3, 1), { bg: _pf['N/A'][0], fg: _pf['N/A'][1] }))
  }

  try {
    rules.push(_tCF('N/A', sh.getRange(4, 1, lr - 3, lc), { bg: _pf['N/A'][0], fg: _pf['N/A'][1] }))
  } catch(eNA) {}

  // ── MEJORAS VISUALES AUTOMÁTICAS (se recalculan solas, sin menú) ─────────
  // RUN duplicado: mismo RUN en dos pacientes (más de una fila con el mismo valor)
  if (COL.RUN <= lc) {
    rules.push(_rCF('=AND($H4<>"",COUNTIF($H$4:$H,$H4)>1)',
      sh.getRange(4, COL.RUN, lr - 3, 1), { bg: '#FEE2E2', fg: '#B91C1C', bold: true }))
  }
  // Observaciones con avisos ⚠️ (fechas inválidas, atenciones mal registradas, etc.)
  if (COL.OBSERVACIONES <= lc) {
    rules.push(_rCF('=ISNUMBER(SEARCH("' + '\u26a0\ufe0f' + '",$' + colToLetter(COL.OBSERVACIONES) + '4))',
      sh.getRange(4, COL.OBSERVACIONES, lr - 3, 1), { bg: '#FEF3C7', fg: '#A16207' }))
  }

  try { sh.setConditionalFormatRules(rules) } catch(e) {}
}

// ─── SECTOR Y ESTADO (colorear por valor) ─────────────────────────────────

function _aplicarColorPorValor(sh, lr, col, coloresMap) {
  if (!sh) sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PAC)
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

// ─── COLOREAR VIGENCIA POR FECHA (un color por fecha) ──────────────────────

function colorearFechas(sh, lr) {
  if (!sh) sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PAC)
  if (!sh) return
  if (!lr) lr = sh.getLastRow()
  if (lr < 4) return
  var params = leerParametros()
  var diasAviso = params['DIAS_AVISO'] || 0
  var lc = sh.getLastColumn()
  if (lc < 1) return
  var rows = lr - 3

  var vitalCol = sh.getRange(4, COL.VITAL, rows, 1).getValues()

  for (var fi = 0; fi < _FECHAS_COLOR.length; fi++) {
    var def = _FECHAS_COLOR[fi]
    var fc = def[0]
    if (fc > lc) continue
    var meses = _mesesControl(params, def[1])
    var rng = sh.getRange(4, fc, rows, 1)
    var src = rng.getValues()
    var bgs = [], fgs = [], fmts = [], wts = []
    for (var r = 0; r < rows; r++) {
      var fallecido = String(vitalCol[r][0] || '').trim() === 'FALLECIDO'
      var st = fallecido ? 'N/A' : _estadoFecha(src[r][0], meses, diasAviso)
      var c = _ESTADO_FECHA_COLORS[st] || _ESTADO_FECHA_COLORS['N/A']
      bgs.push([c[0]])
      fgs.push([c[1]])
      fmts.push([st === 'N/A' ? '@' : 'dd/mm/yyyy'])
      wts.push([st === 'VENCIDO' ? 'bold' : 'normal'])
    }
    try {
      rng.setBackgrounds(bgs).setFontColors(fgs).setFontWeights(wts).setNumberFormats(fmts)
    } catch(e) {
      rng.setBackgrounds(bgs).setFontColors(fgs).setNumberFormats(fmts)
    }
  }
}

// ─── LIMPIAR FORMATO EN COLUMNAS DE TEXTO LIBRE ──────────────────────────

function _FORMATEADAS() {
  var s = {}
  for (var i = 0; i < _FECHAS_COLOR.length; i++) s[_FECHAS_COLOR[i][0]] = true
  s[2] = true
  s[6] = true
  s[10] = true
  s[19] = true
  s[22] = true
  s[66] = true
  s[110] = true
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
  var total = 0

  for (var si = 0; si < ss.getSheets().length; si++) {
    var sh = ss.getSheets()[si]
    var rng = sh.getDataRange()
    var formulas = rng.getFormulas()
    var cambio = false

    for (var r = 0; r < formulas.length; r++) {
      for (var c = 0; c < formulas[r].length; c++) {
        var f = formulas[r][c]
        if (f === '' || f.indexOf('Pacientes') < 0) continue
        var nf = f
        nf = nf.replace(/'Pacientes'!/g, "'" + HOJA_PAC + "'!")
        nf = nf.replace(/(?<![A-Z0-9_])Pacientes!/g, HOJA_PAC + '!')
        if (nf !== f) {
          formulas[r][c] = nf
          cambio = true
          total++
        }
      }
    }

    if (cambio) rng.setFormulas(formulas)
  }

  ss.toast(total + ' fórmulas reparadas en Dashboard', 'PADDS', 4)
  try { recalcularTodo() } catch(e) {}
}

// ─── CREAR HOJA PARÁMETROS ─────────────────────────────────────────────────

function crearParametros() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Creando hoja Parámetros…', 'PADDS', 1)
  var old = ss.getSheetByName('Parámetros')
  var valoresPrevios = {}
  if (old) {
    try {
      var dPrev = old.getDataRange().getValues()
      for (var _pr = 0; _pr < dPrev.length; _pr++) {
        var _pn = String(dPrev[_pr][0] || '').trim().toUpperCase()
        if (!_pn) continue
        if (_pn.indexOf('PARÁMETRO') === 0 || _pn.indexOf('PARAMETRO') === 0) continue
        if (_pn.indexOf('DE VIGENCIA') >= 0 || _pn.indexOf('PADDS') >= 0 || _pn.indexOf('💡') >= 0 || _pn.indexOf('Los cambios') >= 0) continue
        valoresPrevios[_pn] = { valor: dPrev[_pr][1], unidad: dPrev[_pr][2] }
      }
    } catch(ePV) {}
  }
  if (old) ss.deleteSheet(old)

  var sh = ss.insertSheet('Parámetros')
  sh.setTabColor('#8C959D')
  sh.setColumnWidth(1, 340)
  sh.setColumnWidth(2, 100)
  sh.setColumnWidth(3, 110)
  sh.setColumnWidth(4, 560)

  var NAVY = _UI.hdrBg, MID = _UI.accent, ACCENT = _UI.accent2
  var LIGHT = '#EEF2F7', WHITE = '#ffffff', BORDER = _UI.borderMed
  var LIGHT_GREEN = '#DCFCE7', LIGHT_BLUE = '#E0F2FE', LIGHT_ORANGE = '#FFEDD5'
  var LIGHT_PURPLE = '#F3E8FF', LIGHT_TEAL = '#CCFBF1', LIGHT_PINK = '#FCE7F3'
  var BS = SpreadsheetApp.BorderStyle.SOLID

  function _secColor(i) {
    var palette = [LIGHT_GREEN, LIGHT_BLUE, LIGHT_ORANGE, LIGHT_PURPLE, LIGHT_TEAL, LIGHT_PINK]
    return palette[i % palette.length]
  }

  sh.getRange(1, 1, 1, 4).merge()
  sh.getRange(1, 1)
    .setValue('⚙  PARÁMETROS DE VIGENCIA  —  PADDS 2026')
    .setFontFamily(_UI.font).setFontSize(22).setFontWeight('bold')
    .setFontColor(WHITE).setBackground(NAVY)
    .setVerticalAlignment('middle').setHorizontalAlignment('left')
  sh.setRowHeight(1, 52)

  sh.getRange(2, 1, 1, 4).merge()
  sh.getRange(2, 1)
    .setValue('Plazos de vigencia por prestación: valor en la columna VALOR y unidad (MESES, DÍAS o N/A = desactivado) en UNIDAD. Los cambios se aplican con "Recalcular estados" o al editar una fecha.')
    .setFontFamily(_UI.font).setFontSize(11).setFontColor(MID)
    .setBackground(LIGHT).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  sh.setRowHeight(2, 34)

  var GRUPOS = [
    {
      nombre: '🩺  CONTROLES MÉDICOS Y EXÁMENES',
      color: '#0F766E',
      items: [
        ['EXAMENES USUARIO',            6, 'Plazo entre tomas de exámenes del usuario.'],
        ['EXAMENES CUIDADOR',          12, 'Plazo entre exámenes del cuidador.'],
        ['CONTROL MEDICO',              6, 'Plazo entre controles médicos.'],
        ['PIC-1C',                     12, 'Vigencia del PIC (1 componente).'],
        ['PIC-2C',                     12, 'Vigencia del PIC (2 componentes).'],
        ['RECETAS CONTROLADAS',         3, 'Vigencia de recetas controladas.'],
      ],
    },
    {
      nombre: '💓  CONTROLES CARDIOVASCULARES Y ESPECIALIDADES',
      color: '#B91C1C',
      items: [
        ['CCV MEDICO',                  6, 'Plazo entre CCV de médico.'],
        ['CSCV ENFERMERIA',             6, 'Plazo entre CSCV de enfermería.'],
        ['PODOLOGO',                    6, 'Plazo entre atenciones de podología.'],
        ['NUTRICIONISTA',               6, 'Plazo entre controles nutricionales.'],
        ['FONOAUDIOLOGA',              12, 'Plazo entre atenciones de fonoaudiología.'],
        ['CONTROL KINESICO',            6, 'Plazo entre controles kinésicos.'],
        ['ODONTOLOGIA',                12, 'Plazo entre atenciones odontológicas.'],
      ],
    },
    {
      nombre: '📋  EMPA / EMPAM Y PREVENCIÓN',
      color: '#15803D',
      items: [
        ['EMPA/EMPAM CUIDADOR',        12, 'Vigencia del examen preventivo del cuidador.'],
        ['EMPA/EMPAM USUARIO',         12, 'Vigencia del examen preventivo del usuario.'],
        ['CCV VIGENTE CUIDADOR',       12, 'Vigencia de la CCV del cuidador.'],
      ],
    },
    {
      nombre: '🧠  SOCIAL / PSICOLÓGICO',
      color: '#7E22CE',
      items: [
        ['ZARIT',                      12, 'Plazo entre aplicaciones de la escala Zarit.'],
        ['CONSULTA TRABAJADORA SOCIAL',12, 'Plazo entre consultas de trabajadora social.'],
        ['CONSULTA PSICOLOGA',         12, 'Plazo entre consultas psicológicas.'],
      ],
    },
    {
      nombre: '🩹  SONDA FOLEY Y CURACIONES',
      color: '#BE185D',
      items: [
        ['SONDA FOLEY',                 3, 'MESES', 'Plazo entre cambios de sonda.'],
        ['CURACIONES',                  1, 'MESES', 'Plazo entre curaciones de heridas.'],
        ['PROXIMA CURACION',            0, 'MESES', '0 = vence al pasar la fecha indicada.'],
      ],
    },
    {
      nombre: '📚  CAPACITACIONES, INMUNIZACIÓN Y AVISOS',
      color: '#C2410C',
      items: [
        ['CAPACITACIONES',              6, 'MESES', 'Vigencia de capacitaciones del cuidador.'],
        ['INMUNIZACION',               12, 'MESES', 'Plazo de referencia de vacunas.'],
        ['PAÑALES',                     2, 'MESES', 'Plazo entre entregas de pañales.'],
        ['Días aviso',                 15, '', 'Días antes de vencer para marcar POR VENCER.'],
      ],
    },
  ]

  var row = 3
  var BS2 = SpreadsheetApp.BorderStyle.SOLID_MEDIUM

  for (var g = 0; g < GRUPOS.length; g++) {
    var grupo = GRUPOS[g]
    var secBg = _secColor(g)

    sh.getRange(row, 1, 1, 4).merge()
    sh.getRange(row, 1)
      .setValue(grupo.nombre)
      .setFontFamily(_UI.font).setFontSize(12).setFontWeight('bold')
      .setFontColor(WHITE).setBackground(grupo.color)
      .setVerticalAlignment('middle').setHorizontalAlignment('left')
      .setBorder(true, true, true, true, true, true, BORDER, BS)
    sh.setRowHeight(row, 28)
    row++

    sh.getRange(row, 1).setValue('PARÁMETRO')
      .setFontFamily(_UI.font).setFontSize(9).setFontWeight('bold')
      .setFontColor(grupo.color).setBackground('#F1F5F9')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, BORDER, BS2)
    sh.getRange(row, 2).setValue('VALOR')
      .setFontFamily(_UI.font).setFontSize(9).setFontWeight('bold')
      .setFontColor(grupo.color).setBackground('#F1F5F9')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, BORDER, BS2)
    sh.getRange(row, 3).setValue('UNIDAD')
      .setFontFamily(_UI.font).setFontSize(9).setFontWeight('bold')
      .setFontColor(grupo.color).setBackground('#F1F5F9')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, BORDER, BS2)
    sh.getRange(row, 4).setValue('DESCRIPCIÓN')
      .setFontFamily(_UI.font).setFontSize(9).setFontWeight('bold')
      .setFontColor(grupo.color).setBackground('#F1F5F9')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, BORDER, BS2)
    sh.setRowHeight(row, 22)
    row++

    var _dvUnidad = SpreadsheetApp.newDataValidation()
      .requireValueInList(['MESES', 'DIAS', 'N/A'], true).setAllowInvalid(true).build()

    for (var i = 0; i < grupo.items.length; i++) {
      var item = grupo.items[i]
      var bg = i % 2 === 0 ? WHITE : secBg
      var unidad = (item.length >= 4 && String(item[2]).trim())
        ? String(item[2]).trim()
        : 'MESES'
      var descripcion = String(item.length >= 4 ? item[3] : item[2] || '').trim()
      var esDiasAviso = String(item[0]).toUpperCase().indexOf('DIAS') === 0
      var claveParam = String(item[0]).trim().toUpperCase()
      var prev = valoresPrevios[claveParam]
      var _unidadesValidas = { 'MESES': true, 'DIAS': true, 'DÍAS': true, 'N/A': true }
      var _unidadUso = unidad
      if (prev && String(prev.unidad || '').trim()) {
        var _uPrev = String(prev.unidad).trim().toUpperCase()
        if (_uPrev === 'DÍAS') _uPrev = 'DIAS'
        if (_unidadesValidas[_uPrev]) _unidadUso = _uPrev
      }
      var _valorUso = (prev && String(prev.valor) !== '' && prev.valor !== undefined && prev.valor !== null) ? prev.valor : item[1]

      sh.getRange(row, 1).setValue(item[0])
        .setFontFamily(_UI.font).setFontSize(11).setFontWeight('bold')
        .setFontColor(NAVY).setBackground(bg)
        .setVerticalAlignment('middle')
        .setBorder(false, true, false, true, false, false, BORDER, BS)
        .setNote(descripcion)

      sh.getRange(row, 2).setValue(_valorUso)
        .setFontFamily(_UI.font).setFontSize(14).setFontWeight('bold')
        .setFontColor(grupo.color).setBackground(WHITE)
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setNumberFormat('0')
        .setBorder(true, true, true, true, false, false, grupo.color, BS2)

      sh.getRange(row, 3).setValue(esDiasAviso ? '' : _unidadUso)
        .setFontFamily(_UI.font).setFontSize(11).setFontWeight('bold')
        .setFontColor(grupo.color).setBackground(WHITE)
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true, true, true, true, false, false, grupo.color, BS2)
        .setNote(esDiasAviso ? '' : 'Unidad del plazo: MESES · DIAS · N/A (N/A = vigencia desactivada: la fecha no se colorea ni alerta).')
      if (!esDiasAviso) sh.getRange(row, 3).setDataValidation(_dvUnidad)

      sh.getRange(row, 4).setValue(descripcion)
        .setFontFamily(_UI.font).setFontSize(10).setFontColor('#424242')
        .setBackground(bg).setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setBorder(false, true, false, true, false, false, BORDER, BS)

      sh.setRowHeight(row, 40)
      row++
    }

    sh.setRowHeight(row, 6)
    sh.getRange(row, 1, 1, 4).setBackground('#F1F5F9').setBorder(false, false, false, false, false, false)
    row++
  }

  sh.getRange(row, 1, 1, 4).merge()
  sh.getRange(row, 1)
    .setValue('💡  Los cambios se aplican automáticamente al ejecutar "Recalcular estados" o al editar una fecha en ' + HOJA_PAC + '.')
    .setFontFamily(_UI.font).setFontSize(10).setFontColor(MID)
    .setBackground(LIGHT).setVerticalAlignment('middle').setHorizontalAlignment('center')
  sh.setRowHeight(row, 28)
  row++

  sh.getRange(row, 1, 1, 4).merge()
  sh.getRange(row, 1)
    .setValue('PADDS 2026 — Centro de Salud Familiar San Joaquín')
    .setFontFamily(_UI.font).setFontSize(9).setFontColor('#b0bec5')
    .setBackground(LIGHT).setVerticalAlignment('middle').setHorizontalAlignment('center')
  sh.setRowHeight(row, 22)

  _unmergeQueCruzaFila(sh, 1)
  try { sh.setFrozenRows(1) } catch(eFz) {}

  var maxRows = sh.getMaxRows()
  var neededRows = row + 1
  if (maxRows > neededRows) sh.deleteRows(neededRows, maxRows - neededRows + 1)

  _paramCache = null
  try { CacheService.getScriptCache().remove(_PARAM_CACHE_KEY) } catch(e) {}

  ss.toast('Hoja Parámetros creada/actualizada', 'PADDS', 4)
}

// ─── CREAR HOJA INSTRUCCIONES ──────────────────────────────────────────────

function crearInstrucciones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var old = ss.getSheetByName('Referencia Columnas')
  if (old) ss.deleteSheet(old)

  var sh = ss.insertSheet('Referencia Columnas')
  sh.setTabColor('#A3927F')
  sh.setColumnWidth(1, 50)
  sh.setColumnWidth(2, 220)
  sh.setColumnWidth(3, 340)
  sh.setColumnWidth(4, 320)
  sh.setColumnWidth(5, 200)

  var DARK = _UI.hdrBg, MID = _UI.accent, LIGHT = '#E0F2FE', WHITE = '#ffffff'
  var BS = SpreadsheetApp.BorderStyle.SOLID

  sh.getRange(1, 1, 1, 5).merge()
  sh.getRange(1, 1).setValue('PADDS 2026  —  Referencia de Columnas')
    .setFontFamily(_UI.font).setFontSize(20).setFontWeight('bold')
    .setFontColor(WHITE).setBackground(DARK)
    .setVerticalAlignment('middle')
  sh.setRowHeight(1, 44)

  sh.getRange(2, 1, 1, 5).merge()
  sh.getRange(2, 1).setValue('Pase el mouse sobre los encabezados en Pacientes (fila 3) para ver la descripción · 🩺 Pacientes → 🛠️ Datos → "🎨 Formatear hoja" aplica el diseño completo')
    .setFontFamily(_UI.font).setFontSize(11).setFontStyle('italic')
    .setFontColor('#E0F2FE').setBackground(MID)
    .setVerticalAlignment('middle')
  sh.setRowHeight(2, 24)

  var HD = ['N°', 'CAMPO', 'DESCRIPCIÓN', 'VALORES', 'AUTOMATIZACIÓN']
  for (var c = 0; c < HD.length; c++) {
    sh.getRange(3, c + 1).setValue(HD[c])
      .setFontFamily(_UI.font).setFontSize(10).setFontWeight('bold')
      .setFontColor(WHITE).setBackground(MID)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true)
  }
  sh.setRowHeight(3, 24)

  var row = 4
  for (var si = 0; si < PAC_SECCIONES.length; si++) {
    var sec = PAC_SECCIONES[si]
    var colBg = sec.bg2 || sec.bg

    sh.getRange(row, 1, 1, 5).merge()
    sh.getRange(row, 1).setValue(sec.nombre)
      .setFontFamily(_UI.font).setFontSize(11).setFontWeight('bold')
      .setFontColor(WHITE).setBackground(colBg)
      .setVerticalAlignment('middle')
    sh.setRowHeight(row, 22)
    row++

    for (var c = sec.ini; c <= Math.min(sec.fin, _COLUMNAS.length); c++) {
      var col = _COLUMNAS[c]
      if (!col) continue
      var alt = (c - sec.ini) % 2 === 0
      var bg = alt ? WHITE : '#f5f7fa'

      sh.getRange(row, 1).setValue(col.n)
        .setFontFamily('Courier New').setFontSize(9).setFontColor('#888')
        .setBackground(bg).setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true, _UI.border, BS)

      sh.getRange(row, 2).setValue(col.name)
        .setFontFamily(_UI.font).setFontSize(10).setFontWeight('bold')
        .setFontColor('#212121').setBackground(bg).setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true, _UI.border, BS)

      sh.getRange(row, 3).setValue(col.desc)
        .setFontFamily(_UI.font).setFontSize(9).setFontColor('#424242')
        .setBackground(bg).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setBorder(true, true, true, true, true, true, _UI.border, BS)

      sh.getRange(row, 4).setValue(col.vals)
        .setFontFamily(_UI.font).setFontSize(9).setFontColor('#555')
        .setBackground(bg).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setBorder(true, true, true, true, true, true, _UI.border, BS)

      sh.getRange(row, 5).setValue(col.auto)
        .setFontFamily(_UI.font).setFontSize(9).setFontStyle('italic')
        .setFontColor(_UI.accent).setBackground(bg).setVerticalAlignment('middle')
        .setHorizontalAlignment('center')
        .setBorder(true, true, true, true, true, true, _UI.border, BS)

      sh.setRowHeight(row, 20)
      row++
    }
    row++
  }

  sh.getRange(row, 1, 1, 5).merge()
  sh.getRange(row, 1).setValue('Los colores de cada sección coinciden con los de la fila de encabezados en Pacientes. Tooltips disponibles al pasar el mouse sobre fila 3.')
    .setFontFamily(_UI.font).setFontSize(10).setFontColor('#757575').setBackground('#F1F5F9')
    .setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  sh.setRowHeight(row, 24)

  var _mr = sh.getMaxRows()
  if (_mr > row + 3) sh.deleteRows(row + 3, _mr - row - 2)

  ss.toast('Guía de columnas creada (' + (_COLUMNAS._count || _COLUMNAS.length) + ' columnas)', 'PADDS', 4)
}

// ─── CREAR DASHBOARD ──────────────────────────────────────────────────────

function crearDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var pac = ss.getSheetByName(HOJA_PAC)
  if (!pac) {   ss.toast('No se encontró la hoja ' + HOJA_PAC + '. Revisa que exista.', 'Pacientes', 4); return }
  ss.toast('Generando Dashboard…', 'PADDS', 1)

  var old = ss.getSheetByName('Dashboard')
  if (old) ss.deleteSheet(old)

  var sh = ss.insertSheet('Dashboard')
  sh.setTabColor(_UI.tabBW)
  sh.setHiddenGridlines(true)

  var P = "'" + HOJA_PAC + "'!"
  var lc = pac.getLastColumn()
  var SEP = getFormulaSep()
  var BS = SpreadsheetApp.BorderStyle.SOLID
  var BSM = SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  var GREEN = _UI.hdrBg, GREEN_L = '#EEF2F7', GREEN_M = _UI.accent2
  var W = '#ffffff', T = '#212121', TG = '#757575'
  var SEXO = colToLetter(7), EDAD = colToLetter(10)
  var PRIO = colToLetter(COL.PRIORIDAD), DEP = colToLetter(26)
  var EMPC = colToLetter(22), EMPU = colToLetter(66)

  var CARD_BORDER = _UI.accentL

  sh.setColumnWidth(1, 240)
  for (var _ci = 2; _ci <= 8; _ci++) sh.setColumnWidth(_ci, 130)
  sh.setColumnWidth(9, 100)

  var _paramsDash = leerParametros()

  function sec(t, r) {
    sh.getRange(r, 1, 1, 9).merge()
    sh.getRange(r, 1).setValue(t).setFontFamily(_UI.font).setFontSize(13).setFontWeight('bold')
      .setFontColor(W).setBackground(GREEN).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, W, BSM)
    sh.setRowHeight(r, 30)
    sh.getRange(r + 1, 1).setBackground(GREEN_L)
    sh.setRowHeight(r + 1, 4)
  }
  function cardLabel(r, c, w, t, fc) {
    sh.getRange(r, c, 1, w).merge()
      .setValue(t).setFontFamily(_UI.font).setFontSize(10).setFontWeight('bold')
      .setFontColor(fc || T).setBackground(_UI.chipCard).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, false, true, true, false, CARD_BORDER, BS)
  }
  function cardVal(r, c, w, f, fc) {
    if (f) sh.getRange(r, c).setFormula(f)
    var rng = sh.getRange(r, c, 1, w).merge()
    rng.setFontFamily(_UI.font).setFontSize(22).setFontWeight('bold')
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
    .setFontFamily(_UI.font).setFontSize(22).setFontWeight('bold')
    .setFontColor(W).setBackground(GREEN).setHorizontalAlignment('center').setVerticalAlignment('middle')
  sh.setRowHeight(R, 52)
  sh.getRange(R, 1).setBorder(true, true, true, true, true, true, '#1E293B', BSM)
  R++
  sh.getRange(R, 1, 1, 9).merge()
  sh.getRange(R, 1).setFormula('="Actualizado: "&TEXT(NOW()' + SEP + '"dd/MM/yyyy HH:mm")&"  •  Datos en vivo desde la hoja ' + HOJA_PAC + '"')
    .setFontFamily(_UI.font).setFontSize(9).setFontStyle('italic').setFontColor('#ffffff')
    .setBackground(GREEN_M).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(false, true, true, true, true, true, W, BS)
  sh.setRowHeight(R, 24)
  R += 2

  sec('RESUMEN GENERAL', R); R += 2
  var RC = R
  var resumen = [
    [1, 2, 'Total Pacientes', '=COUNTA(' + P + 'C4:C)', GREEN],
    [3, 2, 'Vigentes',        cf(P + 'F4:F', 'VIGENTE'), '#15803D'],
    [5, 2, 'Fallecidos',      cf(P + 'F4:F', 'FALLECIDO'), '#B91C1C'],
    [7, 2, 'Egresados',       cf(P + 'F4:F', 'EGRESO*'), '#C2410C'],
  ]
  resumen.forEach(function(x) { cardLabel(RC, x[0], x[1], x[2], x[3]); cardVal(RC + 1, x[0], x[1], x[3], x[3]) })
  sh.setRowHeight(RC, 22); sh.setRowHeight(RC + 1, 48)
  var _totalValCell = colToLetter(1) + (RC + 1)
  R = RC + 4

  sec('PACIENTES ONCOLOGICOS  —  prevalencia', R); R += 2
  var RO = R
  var oncRng = P + colToLetter(COL.ONCOLOGICO) + '4:' + colToLetter(COL.ONCOLOGICO)
  var oncTotalF = 'COUNTIF(' + oncRng + SEP + '"SI")'
  var oncol = [
    [1, 2, 'Oncológicos (SI)', '=' + oncTotalF, '#0369A1'],
    [3, 2, 'Hombres',          '=COUNTIFS(' + oncRng + SEP + '"SI"' + SEP + P + SEXO + '4:' + SEXO + SEP + '"M")', '#0F766E'],
    [5, 2, 'Mujeres',          '=COUNTIFS(' + oncRng + SEP + '"SI"' + SEP + P + SEXO + '4:' + SEXO + SEP + '"F")', '#C2410C'],
    [7, 2, '% del total',      '=IFERROR(' + colToLetter(1) + (RO + 1) + '/' + _totalValCell + SEP + '0)', '#7E22CE'],
  ]
  oncol.forEach(function(x) { cardLabel(RO, x[0], x[1], x[2], x[4]); cardVal(RO + 1, x[0], x[1], x[3], x[4]) })
  sh.getRange(RO + 1, 7, 1, 2).setNumberFormat('0%')
  sh.getRange(RO + 1, 7, 1, 2).setNote('Porcentaje de pacientes oncológicos sobre el total de pacientes.')
  sh.setRowHeight(RO, 20); sh.setRowHeight(RO + 1, 44)
  R = RO + 4

  sec('DEMOGRAFIA  —  edad y sexo', R); R += 2
  var RC = R
  var demo = [
    [1, 2, 'Hombres',            cf(P + SEXO + '4:' + SEXO, 'M'), '#0369A1'],
    [3, 2, 'Mujeres',            cf(P + SEXO + '4:' + SEXO, 'F'), '#B91C1C'],
    [5, 2, 'Edad Pendiente',     '=IFERROR(' + _totalValCell + '-COUNTIF(' + P + EDAD + '4:' + EDAD + SEP + '">0")' + SEP + '0)', '#C2410C'],
    [7, 2, 'Edad Promedio',      '=ROUND(AVERAGE(' + P + EDAD + '4:' + EDAD + ')' + SEP + ' 0)', GREEN],
    [9, 1, 'Electrodependiente', cf(P + colToLetter(31) + '4:' + colToLetter(31), 'TRUE'), '#7E22CE'],
  ]
  demo.forEach(function(x) { cardLabel(RC, x[0], x[1], x[2], x[4]); cardVal(RC + 1, x[0], x[1], x[3], x[4]) })
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
    ['URGENTE',     cf(P + PRIO + '4:' + PRIO, 'URGENTE'),     '#B91C1C', '#FEE2E2'],
    ['POR REVISAR', cf(P + PRIO + '4:' + PRIO, 'POR REVISAR'), '#C2410C', '#FFEDD5'],
    ['AL DIA',      cf(P + PRIO + '4:' + PRIO, 'AL DIA'),      '#15803D', '#DCFCE7'],
    ['N/A',          cf(P + PRIO + '4:' + PRIO, 'N/A'),         TG,        '#F1F5F9'],
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
    ['VERDE',     cf(P + 'B4:B', 'VERDE'),     '#15803D', '#DCFCE7'],
    ['AMARILLO',  cf(P + 'B4:B', 'AMARILLO'),  '#B45309', '#FEF3C7'],
    ['NARANJO',   cf(P + 'B4:B', 'NARANJO'),   '#C2410C', '#FFEDD5'],
    ['PENDIENTE', cf(P + 'B4:B', 'PENDIENTE'), TG,        '#F1F5F9'],
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
    ['SEVERO',        cf(P + DEP + '4:' + DEP, 'SEVERO'),        '#B91C1C', '#FEE2E2'],
    ['MODERADO',      cf(P + DEP + '4:' + DEP, 'MODERADO'),      '#C2410C', '#FFEDD5'],
    ['LEVE',          cf(P + DEP + '4:' + DEP, 'LEVE'),          '#B45309', '#FEF3C7'],
    ['INDEPENDIENTE', cf(P + DEP + '4:' + DEP, 'INDEPENDIENTE'), '#15803D', '#DCFCE7'],
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
    [1, 2, 'Cuidador EMPA',  cf(P + EMPC + '4:' + EMPC, 'EMPA'),  '#0369A1'],
    [3, 2, 'Cuidador EMPAM', cf(P + EMPC + '4:' + EMPC, 'EMPAM'), '#7E22CE'],
    [5, 2, 'Usuario EMPA',   cf(P + EMPU + '4:' + EMPU, 'EMPA'),  '#0369A1'],
    [7, 2, 'Usuario EMPAM',  cf(P + EMPU + '4:' + EMPU, 'EMPAM'), '#7E22CE'],
  ]
  empa.forEach(function(x) { cardLabel(RE, x[0], x[1], x[2], x[4]); cardVal(RE + 1, x[0], x[1], x[3], x[4]) })
  sh.setRowHeight(RE, 20); sh.setRowHeight(RE + 1, 44)
  R = RE + 4

  sec('VACUNACIÓN  —  cobertura', R); R += 2
  var vacRow = R
  var vacHeaders = ['Vacuna', 'Vacunados', 'Faltante', '%', 'P (prog.)', 'NO / N/A']
  for (var _vhi = 0; _vhi < vacHeaders.length; _vhi++) {
    sh.getRange(vacRow, _vhi + 1).setValue(vacHeaders[_vhi])
      .setFontFamily(_UI.font).setFontSize(9).setFontWeight('bold')
      .setFontColor(W).setBackground(GREEN).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, GREEN, BS)
  }
  sh.getRange(vacRow, 1).setNote('Vacunados cuenta: SI y R (refuerzo). Faltante = total de pacientes menos vacunados. P = programada · NO / N/A = sin cobertura.')
  sh.setRowHeight(vacRow, 24)

  var vacas = [
    ['INFLUENZA USUARIO', 80],
    ['NEUMO23 USUARIO', 81],
    ['INFLUENZA CUIDADOR', 82],
    ['NEUMO23 CUIDADOR', 83],
  ]
  for (var va = 0; va < vacas.length; va++) {
    var vrow = vacRow + 1 + va
    var vcl = colToLetter(vacas[va][1])
    var vbg = va % 2 === 0 ? W : '#F8FAFC'
    var vr = P + vcl + '4:' + vcl
    var vf = 'COUNTIF(' + vr + SEP + '">0")+COUNTIF(' + vr + SEP + '"SI")+COUNTIF(' + vr + SEP + '"R")'
    sh.getRange(vrow, 1).setValue(vacas[va][0])
      .setFontFamily(_UI.font).setFontSize(10).setFontColor(T).setBackground(vbg).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.getRange(vrow, 2).setFormula('=' + vf)
      .setFontFamily(_UI.font).setFontSize(10).setFontColor(GREEN).setBackground(vbg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.getRange(vrow, 3).setFormula('=' + _totalValCell + '-(' + vf + ')')
      .setNumberFormat('0')
      .setFontFamily(_UI.font).setFontSize(10).setFontColor(TG).setBackground(vbg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.getRange(vrow, 4).setFormula('=IFERROR((' + vf + ')/' + _totalValCell + SEP + '0)')
      .setNumberFormat('0%')
      .setFontFamily(_UI.font).setFontSize(10).setFontWeight('bold').setFontColor(GREEN).setBackground(vbg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.getRange(vrow, 5).setFormula('=COUNTIF(' + vr + SEP + '"P")')
      .setFontFamily(_UI.font).setFontSize(10).setFontColor('#C2410C').setBackground(vbg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.getRange(vrow, 6).setFormula('=COUNTIF(' + vr + SEP + '"NO")+COUNTIF(' + vr + SEP + '"N/A")')
      .setFontFamily(_UI.font).setFontSize(10).setFontColor(TG).setBackground(vbg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.setRowHeight(vrow, 22)
  }
  R = vacRow + 1 + vacas.length + 2

  sec('PATOLOGIAS CRONICAS  —  prevalencia', R); R += 2
  var pHeaders = pac.getRange(3, 32, 1, 17).getValues()[0]
  var pathRow = R

  var pathHeaders = ['Patologia', 'Pacientes', 'Faltante', '%']
  for (var _phi = 0; _phi < pathHeaders.length; _phi++) {
    sh.getRange(pathRow, _phi + 1).setValue(pathHeaders[_phi])
      .setFontFamily(_UI.font).setFontSize(9).setFontWeight('bold')
      .setFontColor(W).setBackground(GREEN).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, GREEN, BS)
  }
  sh.getRange(pathRow, 1).setNote('Prevalencia: pacientes con la patología marcada (casilla) sobre el total. Faltante = total de pacientes menos los que tienen la patología.')
  sh.setRowHeight(pathRow, 24)

  for (var p = 0; p < pHeaders.length; p++) {
    var pr = pathRow + 1 + p
    var cn = 32 + p
    var cl = colToLetter(cn)
    var nm = String(pHeaders[p] || '').trim() || 'Patologia ' + cn
    var bg = p % 2 === 0 ? W : '#F8FAFC'

    var countFormula = ''
    if (cn === 48) countFormula = 'COUNTA(' + P + cl + '4:' + cl + ')'
    else countFormula = 'COUNTIF(' + P + cl + '4:' + cl + SEP + 'TRUE)'

    sh.getRange(pr, 1).setValue(nm)
      .setFontFamily(_UI.font).setFontSize(10).setFontColor(T).setBackground(bg).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.getRange(pr, 2).setFormula('=' + countFormula)
      .setFontFamily(_UI.font).setFontSize(10).setFontColor(GREEN).setBackground(bg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.getRange(pr, 3).setFormula('=' + _totalValCell + '-(' + countFormula + ')')
      .setNumberFormat('0')
      .setFontFamily(_UI.font).setFontSize(10).setFontColor(TG).setBackground(bg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.getRange(pr, 4).setFormula('=IFERROR(' + countFormula + '/' + _totalValCell + SEP + '0)')
      .setNumberFormat('0%')
      .setFontFamily(_UI.font).setFontSize(10).setFontWeight('bold').setFontColor(GREEN).setBackground(bg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.setRowHeight(pr, 22)
  }
  R = pathRow + 1 + pHeaders.length + 2

  sec('VIGENCIAS POR AREA', R); R += 2
  var vigHeaders = ['Control / Area', 'AL DIA', 'POR\nVENCER', 'VENCIDO', 'PENDTE', 'N/A', 'TOTAL', '% AL DIA', 'BARRAS']
  var vigBgs = [GREEN, '#15803D', '#C2410C', '#B91C1C', '#B45309', TG, GREEN, '#15803D', '#1E293B']
  vigHeaders.forEach(function(h, i) {
    sh.getRange(R, i + 1).setValue(h)
      .setFontFamily(_UI.font).setFontSize(9).setFontWeight('bold')
      .setFontColor(W).setBackground(vigBgs[i] || GREEN_M).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, GREEN, BS)
  })
  sh.setRowHeight(R, 30)
  R++
  var vigRows = 0
  for (var v = 0; v < _CONTROL_FECHAS.length; v++) {
    var def = _CONTROL_FECHAS[v]
    var fc = def[1]
    if (fc > lc) continue
    var vr = R + vigRows
    var nm2 = def[0]
    var bg2 = vigRows % 2 === 0 ? '#FFFFFF' : '#F8FAFC'
    var mesesM = _mesesControl(_paramsDash, def[2])

    sh.getRange(vr, 1).setValue(nm2)
      .setFontFamily(_UI.font).setFontSize(9).setFontWeight('bold').setFontColor(T).setBackground(bg2).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)

    sh.getRange(vr, 2, 1, 5).setValues([[0, 0, 0, 0, 0]])
      .setFontFamily(_UI.font).setFontSize(9).setFontWeight('bold').setFontColor(T).setBackground(bg2)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.getRange(vr, 7).setFormula('=SUM(B' + vr + ':F' + vr + ')')
      .setFontFamily(_UI.font).setFontSize(9).setFontWeight('bold').setFontColor(GREEN).setBackground(bg2)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.getRange(vr, 8).setFormula('=IF(G' + vr + '=0' + SEP + ' ""' + SEP + ' ROUND(B' + vr + '/G' + vr + '*100' + SEP + ' 0))')
      .setFontFamily(_UI.font).setFontSize(9).setFontWeight('bold').setFontColor('#15803D').setBackground(bg2)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.getRange(vr, 9).setFormula('=SPARKLINE(B' + vr + ':F' + vr + ')')
      .setBackground(bg2).setVerticalAlignment('middle').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.setRowHeight(vr, 22)
    vigRows++
  }
  R += vigRows + 2

  sec('ALERTAS RESUMEN', R); R += 2
  var ar = R
  var alertas = [
    ['VENCIDOS',   '0', '#B91C1C', '#FEE2E2'],
    ['POR VENCER', '0', '#C2410C', '#FFEDD5'],
    ['PENDIENTES', _CONTROL_FECHAS.filter(function(d) { return d[1] <= lc }).map(function(d) {
      return 'SUMPRODUCT(--(LEN(TRIM(' + P + colToLetter(d[1]) + '4:' + colToLetter(d[1]) + '))=0))'
    }).join('+'), '#B45309', '#FEF3C7'],
  ]
  alertas.forEach(function(x, i) {
    var ac = 1 + i * 3
    sh.getRange(ar, ac, 1, 3).merge()
    sh.getRange(ar, ac).setValue(x[0])
      .setFontFamily(_UI.font).setFontSize(13).setFontWeight('bold')
      .setFontColor(x[2]).setBackground(x[3]).setVerticalAlignment('middle').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true, x[2], BS)
    sh.getRange(ar + 1, ac, 1, 3).merge()
    sh.getRange(ar + 1, ac).setFormula('=' + x[1])
      .setFontFamily(_UI.font).setFontSize(26).setFontWeight('bold')
      .setFontColor(x[2]).setBackground(W).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, x[2], BS)
    sh.setRowHeight(ar, 24)
    sh.setRowHeight(ar + 1, 52)
  })
  R = ar + 4

  sec('GRÁFICOS — distribución visual', R); R += 2
  var gT = R

  var gs = [['SECTOR', 'Pacientes']].concat([
    ['VERDE',     cf(P + 'B4:B', 'VERDE')],
    ['AMARILLO',  cf(P + 'B4:B', 'AMARILLO')],
    ['NARANJO',   cf(P + 'B4:B', 'NARANJO')],
    ['PENDIENTE', cf(P + 'B4:B', 'PENDIENTE')],
  ])
  var gx = [['SEXO', 'Pacientes']].concat([
    ['M', cf(P + SEXO + '4:' + SEXO, 'M')],
    ['F', cf(P + SEXO + '4:' + SEXO, 'F')],
  ])

  var _tinte = {
  VERDE: '#15803D', AMARILLO: '#B45309', NARANJO: '#C2410C', PENDIENTE: '#64748B',
  M: '#0369A1', F: '#BE185D',
  }
  function miniTabla(tbl, col) {
    for (var tr = 0; tr < tbl.length; tr++) {
      var rr = gT + tr
      var esHdr = tr === 0
      var bg = esHdr ? GREEN : (tr % 2 === 0 ? W : '#F8FAFC')
      var lb = tbl[tr][0]
      var fc = esHdr ? W : (_tinte[lb] || T)
      sh.getRange(rr, col).setValue(lb)
        .setFontFamily(_UI.font).setFontSize(esHdr ? 9 : 10).setFontWeight('bold')
        .setFontColor(fc).setBackground(esHdr ? GREEN : (_tinte[lb] ? '#ffffff' : bg))
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
      var c2 = sh.getRange(rr, col + 1)
      if (esHdr) c2.setValue('N')
      else c2.setFormula(tbl[tr][1])
      c2.setFontFamily(_UI.font).setFontSize(esHdr ? 9 : 10).setFontWeight('bold')
        .setFontColor(esHdr ? W : GREEN).setBackground(bg)
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
      sh.setRowHeight(rr, esHdr ? 18 : 20)
    }
  }
  miniTabla(gs, 1)
  miniTabla(gx, 4)

  var gA = gT + Math.max(gs.length, gx.length) + 2
  var nChart = 0
  if (_insertarGrafico(sh, Charts.ChartType.COLUMN, sh.getRange(gT, 1, gs.length, 2), gA, 1, 320, 250,
      { title: 'Distribución por SECTOR', colors: ['#15803D'] })) nChart++
  if (_insertarGrafico(sh, Charts.ChartType.PIE, sh.getRange(gT, 4, gx.length, 2), gA, 6, 320, 250,
      { pieHole: 0.45, title: 'Sexo', colors: ['#0369A1', '#B91C1C'] })) nChart++
  if (!nChart) ss.toast('No se pudieron insertar los gráficos (revisa permisos)', 'Dashboard', 6)
  R = gA + 14

  sh.getRange(R, 1, 1, 9).merge()
  sh.getRange(R, 1).setValue('Leyenda de clasificacion de controles')
    .setFontFamily(_UI.font).setFontSize(12).setFontWeight('bold').setFontColor(W).setBackground(_UI.hdrBg)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, W, BSM)
  sh.setRowHeight(R, 28)
  R += 2
  var _leyColors = { bg: ['#DCFCE7', '#FEF3C7', '#FEE2E2', '#FEF3C7', '#F1F5F9'], fg: ['#15803D', '#C2410C', '#B91C1C', '#B45309', TG] }
  var stLbl = [
    ['AL DIA',       'Fecha dentro del plazo configurado en Parametros (color verde en la celda)'],
    ['POR VENCER',   'Proximo a vencer — segun DIAS_AVISO en Parametros (color ambar)'],
    ['VENCIDO',      'Fecha supero el plazo maximo de vigencia (color rojo)'],
    ['PENDIENTE',    'Sin fecha registrada o formato no valido (color amarillo)'],
    ['N/A',          'No aplica / fecha no vigente (color gris)'],
  ]
  for (var _li = 0; _li < stLbl.length; _li++) {
    var s = stLbl[_li]
    sh.getRange(R + _li, 1).setValue(s[0])
      .setFontFamily(_UI.font).setFontSize(10).setFontWeight('bold')
      .setFontColor(_leyColors.fg[_li]).setBackground(_leyColors.bg[_li]).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.getRange(R + _li, 2, 1, 8).merge().setValue(s[1])
      .setFontFamily(_UI.font).setFontSize(10).setFontColor('#333').setBackground(_leyColors.bg[_li])
      .setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.setRowHeight(R + _li, 22)
  }
  R += stLbl.length + 2

  sh.getRange(R, 1).setValue('Controles por fecha unica (el color de la celda indica la vigencia):')
    .setFontFamily(_UI.font).setFontSize(10).setFontWeight('bold').setFontColor(W).setBackground('#334155')
    .setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, GREEN, BS)
  sh.getRange(R, 2, 1, 8).merge()
  sh.setRowHeight(R, 24)
  R++
  for (var _ai2 = 0; _ai2 < _CONTROL_FECHAS.length; _ai2++) {
    var c = _CONTROL_FECHAS[_ai2]
    var bgA2 = _ai2 % 2 === 0 ? '#f5f7fa' : '#eef1f5'
    sh.getRange(R + _ai2, 1).setValue(c[0])
      .setFontFamily(_UI.font).setFontSize(9).setFontColor(T).setBackground(bgA2).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    var clF = colToLetter(c[1])
    sh.getRange(R + _ai2, 2, 1, 8).merge().setValue('Fecha: columna ' + clF + '  →  vigencia por color')
      .setFontFamily('Consolas,monospace').setFontSize(9).setFontColor('#555').setBackground(bgA2)
      .setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
    sh.setRowHeight(R + _ai2, 20)
  }
  R += _CONTROL_FECHAS.length + 2

  sh.getRange(R, 1, 1, 9).merge()
  sh.getRange(R, 1).setValue('Prioridad General (col ' + COL.PRIORIDAD + '): VENCIDO → URGENTE | POR VENCER → POR REVISAR | AL DIA → AL DIA | sin datos → N/A')
    .setFontFamily(_UI.font).setFontSize(9).setFontStyle('italic').setFontColor('#555').setBackground('#F1F5F9')
    .setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#E2E8F0', BS)
  sh.setRowHeight(R, 24)
  R++

  var _mr2 = sh.getMaxRows()
  if (_mr2 > R + 2) sh.deleteRows(R + 2, _mr2 - R - 1)

  actualizarDashboard()
  ss.toast('Dashboard creado y vigencias actualizadas', 'PADDS', 4)
}

// ─── GRÁFICOS EMBEBIDOS DEL DASHBOARD ────────────────────────────────────────

function _insertarGrafico(sh, tipo, rng, fila, col, w, h, opts) {
  try {
    var b = sh.newChart()
      .setChartType(tipo)
      .addRange(rng)
      .setPosition(fila, col, 0, 0)
      .setOption('width', w).setOption('height', h)
      .setOption('fontName', _UI.font)
      .setOption('chartArea', { width: '78%', height: '72%' })
      .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
      .setOption('titleTextStyle', { fontSize: 12, bold: true })
    if (opts) {
      for (var k in opts) {
        if (Object.prototype.hasOwnProperty.call(opts, k)) b.setOption(k, opts[k])
      }
    }
    sh.insertChart(b.build())
    return true
  } catch (e) { return false }
}

// ─── ACTUALIZAR VIGENCIAS Y ALERTAS DEL DASHBOARD (valores por script) ──────

function actualizarDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var BS = SpreadsheetApp.BorderStyle.SOLID
  var pac = ss.getSheetByName(HOJA_PAC)
  var sh = ss.getSheetByName('Dashboard')
  if (!pac || !sh) { ss.toast('Faltan: hoja ' + HOJA_PAC + ' o Dashboard', 'PADDS', 4); return }

  var lr = pac.getLastRow()
  var lc = pac.getLastColumn()
  if (lr < 4) { ss.toast('No hay datos en ' + HOJA_PAC + ' para actualizar el Dashboard', 'Dashboard', 4); return }
  var data = pac.getRange(4, 1, lr - 3, lc).getValues()

  var shVals = sh.getDataRange().getValues()
  var vigH = -1
  for (var i = 0; i < shVals.length; i++) {
    if (String(shVals[i][0] || '').trim() === 'Control / Area') { vigH = i; break }
  }
  if (vigH < 0) { ss.toast('No encontré la tabla VIGENCIAS POR AREA. Ejecuta primero 📊 Crear Dashboard.', 'Dashboard', 5); return }

  var params = leerParametros()
  var diasAv = params['DIAS_AVISO'] || 15
  var mesesDe = function(key) { var m = Number(params[key]); return m > 0 ? m : 6 }

  var totV = 0, totPv = 0
  var filaVig = vigH + 2
  for (var v = 0; v < _CONTROL_FECHAS.length; v++) {
    var def = _CONTROL_FECHAS[v]
    var col = def[1]
    if (col > lc) continue
    var mesesM = mesesDe(def[2])
    var cAl = 0, cPv = 0, cV = 0, cPd = 0, cNa = 0
    for (var r = 0; r < data.length; r++) {
      var st = _estadoFecha(data[r][col - 1], mesesM, diasAv)
      if (st === 'AL DIA') cAl++
      else if (st === 'POR VENCER') cPv++
      else if (st === 'VENCIDO') cV++
      else if (st === 'N/A') cNa++
      else cPd++
    }
    var vr = filaVig + v
    sh.getRange(vr, 2, 1, 5).setValues([[cAl, cPv, cV, cPd, cNa]])
    totV += cV
    totPv += cPv
  }

  var secAl = -1
  for (var i3 = 0; i3 < shVals.length; i3++) {
    if (String(shVals[i3][0] || '').trim() === 'ALERTAS RESUMEN') { secAl = i3; break }
  }
  if (secAl >= 0) {
    sh.getRange(secAl + 4, 1).setValue(totV)
    sh.getRange(secAl + 4, 4).setValue(totPv)

    var conV = 0, conPv = 0, sinCtrl = 0
    for (var r = 0; r < data.length; r++) {
      var has = false, hasV = false, hasPv = false
      for (var v = 0; v < _CONTROL_FECHAS.length; v++) {
        var colB = _CONTROL_FECHAS[v][1]
        if (colB > lc) continue
        var st2 = _estadoFecha(data[r][colB - 1], mesesDe(_CONTROL_FECHAS[v][2]), diasAv)
        if (st2 === 'VENCIDO') { hasV = true; has = true }
        else if (st2 === 'POR VENCER') { hasPv = true; has = true }
        else if (st2 !== 'PENDIENTE' && st2 !== 'N/A') has = true
      }
      if (hasV) conV++
      if (hasPv) conPv++
      if (!has) sinCtrl++
    }
    sh.getRange(secAl + 5, 1, 1, 9).merge()
    sh.getRange(secAl + 5, 1)
      .setValue('Con ≥1 control VENCIDO: ' + conV + '   ·   Con ≥1 POR VENCER: ' + conPv + '   ·   Sin ningún control registrado: ' + sinCtrl + '   (' + data.length + ' pacientes)')
      .setFontFamily(_UI.font).setFontSize(9).setFontWeight('bold').setFontColor('#92400E')
      .setBackground('#FEF3C7').setVerticalAlignment('middle').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true, '#FCD34D', BS)
    sh.setRowHeight(secAl + 5, 22)
  }

  ss.toast('Vigencias y alertas actualizadas', 'PADDS', 3)
}

// ─── RESALTAR FILA ACTIVA (onSelectionChange) ──────────────────────────────

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

    var sheets = [HOJA_PAC, HOJA_FORM]
    for (var si = 0; si < sheets.length; si++) {
      var sh = ss.getSheetByName(sheets[si])
      if (!sh) continue

      // fórmula de la regla (getBooleanCriteria) para no dejar duplicados que

      var existing = []
      try {
        existing = sh.getConditionalFormatRules()
      } catch(e2) {
        existing = []
      }
      var keep = []
      for (var ri = 0; ri < existing.length; ri++) {
        var hasResalte = false
        try {
          var bc = existing[ri].getBooleanCriteria()
          if (bc && bc.criteriaValues &&
              String(bc.criteriaValues[0] || '').indexOf('_Resalte') !== -1) hasResalte = true
        } catch(e3) {}
        if (!hasResalte) keep.push(existing[ri])
      }

      var mr = sh.getMaxRows()
      var mc = sh.getLastColumn() || sh.getMaxColumns()
      var range = sh.getRange(1, 1, mr, mc)
      var formula = '=AND(ROW()>=4, ROW()=INDIRECT("_Resalte!B1"), INDIRECT("_Resalte!A1")="' + sheets[si] + '")'
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground('#CCFBF1')
        .setRanges([range])
        .build()
      keep.push(rule)
      sh.setConditionalFormatRules(keep)
    }
    try { ss.toast('Resaltado de fila activa configurado en ' + sheets.length + ' hoja(s)', 'PADDS', 3) } catch(eT) {}
  } catch(e) {
    try { ss.toast('Error al configurar resaltado: ' + e.message, 'PADDS', 5) } catch(e5) {}
  }
}
