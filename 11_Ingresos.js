// ─── INGRESOS: CREAR HOJA + FORMATO + PASO A PACIENTES ──────────────────────

// ⚠️ REGLAS DE SEGURIDAD (obligatorias, no eliminar):
// 1. ESTE CÓDIGO NUNCA ELIMINA HOJAS. Nunca. Ni esta ni ninguna otra.

// Nunca borra ni altera el contenido. Agrega UNA fila nueva al final

// Pacientes NO se duplica y la fila NO se elimina (queda con nota).

// y la fila se ELIMINA de INGRESOS (se elimina la fila, nunca una

// formatea sola vía onEdit. Así las filas nuevas nunca quedan sin formato.

// ─── DETECCIÓN DE ESTRUCTURA (no asume columnas fijas) ──────────────────────

function _ingDetectarColumnas(sh) {
  var lc = sh.getLastColumn()
  if (lc < 2) return null
  var maxRow = Math.min(10, sh.getMaxRows())
  var data = sh.getRange(1, 1, maxRow, lc).getValues()
  var best = null
  var bestScore = 0

  for (var r = 0; r < data.length; r++) {
    var score = 0
    var cols = {
      run: 0, nombre: 0, apPaterno: 0, apMaterno: 0,
      estado: 0, accion: 0, fechaSolicitud: 0,
      direccion: 0, telefono: 0, derivado: 0, antecedentes: 0
    }
    for (var c = 0; c < data[r].length; c++) {
      var h = String(data[r][c] || '').toUpperCase()
      if (!h) continue
      if (/RUT|RUN/.test(h) && !cols.run) { cols.run = c + 1; score++ }
      if (/NOMBRE/.test(h) && !cols.nombre) { cols.nombre = c + 1; score++ }
      if (/APELLIDO/.test(h)) {
        if (!cols.apPaterno) { cols.apPaterno = c + 1; score++ }
        else if (!cols.apMaterno) { cols.apMaterno = c + 1; score++ }
      }
      if (/^ESTADO/.test(h) && !cols.estado) { cols.estado = c + 1; score++ }
      if (/OBSERVACION|OBS/.test(h) && !cols.accion) { cols.accion = c + 1; score++ }
      if (/FECHA|F\./.test(h) && /SOLICITUD|REGISTRO/.test(h) && !cols.fechaSolicitud) { cols.fechaSolicitud = c + 1; score++ }
      if (/DIRECCION/.test(h) && !cols.direccion) { cols.direccion = c + 1 }
      if (/TELEFONO|FONO/.test(h) && !cols.telefono) { cols.telefono = c + 1 }
      if (/DERIVADO/.test(h) && !cols.derivado) { cols.derivado = c + 1 }
      if (/ANTECEDENTE/.test(h) && !cols.antecedentes) { cols.antecedentes = c + 1 }
    }
    if (score > bestScore && cols.run && cols.nombre && (cols.accion || cols.estado)) {
      bestScore = score
      best = { headerRow: r + 1, cols: cols }
    }
  }
  return best
}

function _ingVacio(v) {
  return v == null || String(v).trim() === ''
}

function _ingEsIngresa(v) {
  var s = String(v || '').trim().toUpperCase()
  return s === 'INGRESA' || s === 'INGRESO' || s === 'INGRESADO'
}

function _ingColorAccion(v) {
  var s = String(v || '').trim().toUpperCase()
  if (s === 'INGRESA' || s === 'INGRESO' || s === 'INGRESADO') return '#DCFCE7'
  if (s === 'PENDIENTE') return '#FEF3C7'
  if (s.indexOf('NO INGRESA') === 0) return '#FEE2E2'
  return ''
}

function _ingRepetidasMismoRut(sh, d, runN, exceptRows) {
  var res = []
  var hr = d.headerRow
  var lr = sh.getLastRow()
  if (lr <= hr || !d.cols.run) return res
  var exc = {}
  for (var x = 0; x < exceptRows.length; x++) exc[exceptRows[x]] = true
  var runs = sh.getRange(hr + 1, d.cols.run, lr - hr, 1).getValues()
  for (var i = 0; i < runs.length; i++) {
    var r = hr + 1 + i
    if (exc[r]) continue
    var rRaw = String(runs[i][0] || '').trim()
    if (!/[0-9]/.test(rRaw)) continue
    if (formatearRUT(rRaw).toUpperCase() !== String(runN).toUpperCase()) continue
    res.push(r)
  }
  return res
}

// ─── FORMATO MEJORADO (compartido por crear y mejorar la hoja) ──────────────

var _ING_PREFILAS = 500

function _ingAplicarFormato(sh, d) {
  var hr = d.headerRow
  var lr = sh.getLastRow()
  var lc = sh.getLastColumn()
  var cols = d.cols

  var lcF = 0
  for (var c = 1; c <= lc; c++) {
    if (!_ingVacio(sh.getRange(hr, c).getValue())) lcF = c
  }
  if (!lcF) lcF = lc
  d._lcF = lcF

  // Panel superior (título + estadísticas en vivo) solo con el layout nuevo (cabecera en fila 4)
  if (hr === 4) _ingConstruirPanel(sh, d)

  sh.getRange(hr, 1, 1, lcF)
    .setBackground(_UI.hdrBg).setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontFamily(_UI.font).setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setBorder(false, true, true, true, true, false, '#1E293B', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  sh.setRowHeight(hr, 32)
  // Congelar un rango que corta una combinación lanza "Debes seleccionar todas
  // las celdas…": descombinar primero cualquier intervalo que cruce esta fila.
  _unmergeQueCruzaFila(sh, hr)
  try { sh.setFrozenRows(hr) } catch (eFz) {}

  var desde = hr + 1
  var hasta = Math.min(sh.getMaxRows(), hr + _ING_PREFILAS)
  var total = hasta - desde + 1
  if (total < 1) return

  var bgArr = []
  for (var r = desde; r <= hasta; r++) {
    var bg = (r - hr) % 2 === 1 ? _UI.zebraBg[0] : _UI.zebraBg[1]
    var rowArr = []
    for (var c2 = 0; c2 < lcF; c2++) rowArr.push(bg)
    bgArr.push(rowArr)
  }
  sh.getRange(desde, 1, total, lcF).setBackgrounds(bgArr)

  sh.getRange(desde, 1, total, lcF)
    .setFontFamily(_UI.font).setFontSize(10)
    .setVerticalAlignment('middle').setFontColor('#202124')
  sh.setRowHeights(desde, total, _UI.rowH)

  if (cols.run) sh.getRange(desde, cols.run, total, 1).setHorizontalAlignment('center')
  if (cols.accion) sh.getRange(desde, cols.accion, total, 1).setHorizontalAlignment('center')
  if (cols.estado) sh.getRange(desde, cols.estado, total, 1).setHorizontalAlignment('center')
  if (cols.fechaSolicitud) sh.getRange(desde, cols.fechaSolicitud, total, 1).setHorizontalAlignment('center')

  if (cols.run) sh.getRange(desde, cols.run, total, 1).setNumberFormat('@').setFontFamily('Courier New')
  if (cols.fechaSolicitud) sh.getRange(desde, cols.fechaSolicitud, total, 1).setNumberFormat('dd/MM/yyyy')
  if (cols.antecedentes) sh.getRange(desde, cols.antecedentes, total, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)

  if (cols.accion) {
    var t1 = false
    try { t1 = sh.getRange(desde, cols.accion).getDataValidation() != null } catch (e1) {}
    if (!t1) {
      sh.getRange(desde, cols.accion, total, 1).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['INGRESA', 'NO INGRESA', 'PENDIENTE'], true).setAllowInvalid(true).build())
    }
  }
  if (cols.estado) {
    var t2 = false
    try { t2 = sh.getRange(desde, cols.estado).getDataValidation() != null } catch (e2) {}
    if (!t2) {
      sh.getRange(desde, cols.estado, total, 1).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['GESTIONADOS', 'PENDIENTES', 'EN ESPERA'], true).setAllowInvalid(true).build())
    }
  }

  sh.getRange(desde, 1, total, lcF)
    .setBorder(true, true, true, true, true, true, _UI.border, SpreadsheetApp.BorderStyle.SOLID)

  _ingAplicarCF(sh, d, desde, total)
  _ingAplicarFiltro(sh, hr, lcF)

  var w = {}
  if (cols.run) w[cols.run] = 110
  if (cols.nombre) w[cols.nombre] = 200
  if (cols.apPaterno) w[cols.apPaterno] = 170
  if (cols.apMaterno) w[cols.apMaterno] = 170
  if (cols.estado) w[cols.estado] = 110
  if (cols.accion) w[cols.accion] = 140
  if (cols.fechaSolicitud) w[cols.fechaSolicitud] = 120
  if (cols.direccion) w[cols.direccion] = 260
  if (cols.telefono) w[cols.telefono] = 120
  if (cols.derivado) w[cols.derivado] = 140
  if (cols.antecedentes) w[cols.antecedentes] = 360
  for (var cw in w) {
    if (w.hasOwnProperty(cw)) sh.setColumnWidth(Number(cw), w[cw])
  }

  sh.setTabColor(_UI.tabBW)
}

// ─── PANEL SUPERIOR: TÍTULO + TARJETAS DE ESTADÍSTICAS EN VIVO ──────────────

var _ING_NOTA =
  'INGRESOS · LISTA DE ADMISIÓN\n\n' +
  '· OBSERVACION: INGRESA (enviar a Pacientes) · NO INGRESA · PENDIENTE\n' +
  '· ESTADO: GESTIONADOS · PENDIENTES · EN ESPERA\n' +
  '· Al marcar INGRESA con la confirmación activada (📥 Ingresos → ⚙️ Confirmar "INGRESA"):\n' +
  '  se verifica el RUT (sin duplicados) y la fila se elimina de INGRESOS.\n' +
  '· Filtros: usa las flechas de la fila de encabezados.\n' +
  '· Nunca se borran hojas.'

function _ingConstruirPanel(sh, d) {
  var lcF = d._lcF || 11
  var hr = d.headerRow
  var rD = hr + 1
  var SEP = getFormulaSep()
  var cl = colToLetter

  // Limpiar combinaciones previas del área del panel (filas 1-3). Se descombinan
  // UNA POR UNA (rango exacto) para evitar el error "Debes seleccionar todas las
  // celdas de un intervalo combinado" al tocar solo parte de una combinación.
  var mergesPrev = []
  try { mergesPrev = sh.getMergedRanges() } catch (eM0) {}
  for (var mi = 0; mi < mergesPrev.length; mi++) {
    var mm = mergesPrev[mi]
    if (mm.getRow() <= 3) {
      try { mm.unmerge() } catch (eMU) {}
    }
  }

  var t
  try { t = sh.getRange(1, 1, 1, lcF).merge() } catch (eM) { t = sh.getRange(1, 1) }
  t.setBackground(_UI.hdrBg)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  var txtTit = 'INGRESOS · LISTA DE ADMISIÓN\nRegistro de admisión y derivación a Pacientes · PADDS'
  var nlTit = txtTit.indexOf('\n')
  var okRT = false
  try {
    var rtTit = SpreadsheetApp.newRichTextValue().setText(txtTit)
      .setTextStyle(0, nlTit, SpreadsheetApp.newTextStyle().setForegroundColor('#FFFFFF')
        .setFontSize(18).setBold(true).setFontFamily(_UI.font).build())
      .setTextStyle(nlTit + 1, txtTit.length, SpreadsheetApp.newTextStyle().setForegroundColor(_UI.hdrSub)
        .setFontSize(10).setFontFamily(_UI.font).build())
      .build()
    t.setRichTextValue(rtTit)
    okRT = true
  } catch (eRT) {}
  if (!okRT) {
    t.setValue('INGRESOS · LISTA DE ADMISIÓN')
      .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(15).setFontFamily(_UI.font)
  }
  t.setBorder(false, false, true, false, false, false, '#1E293B', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  t.setNote(_ING_NOTA)
  sh.setRowHeight(1, 42)

  var cards = []
  if (d.cols.run) {
    cards.push({ label: 'TOTAL SOLICITUDES', col1: 1, col2: 3, color: _UI.hdrBg, tint: '#EEF2F7',
      f: 'COUNTA(' + cl(d.cols.run) + rD + ':' + cl(d.cols.run) + ')' })
  }
  if (d.cols.estado) {
    cards.push({ label: 'PENDIENTES', col1: 4, col2: 5, color: '#B45309', tint: '#FEF3C7',
      f: 'COUNTIF(' + cl(d.cols.estado) + rD + ':' + cl(d.cols.estado) + SEP + '"PENDIENTES")' })
    cards.push({ label: 'EN ESPERA', col1: 6, col2: 7, color: '#0E7490', tint: '#E0F2FE',
      f: 'COUNTIF(' + cl(d.cols.estado) + rD + ':' + cl(d.cols.estado) + SEP + '"EN ESPERA")' })
    cards.push({ label: 'GESTIONADOS', col1: 8, col2: 9, color: '#15803D', tint: '#DCFCE7',
      f: 'COUNTIF(' + cl(d.cols.estado) + rD + ':' + cl(d.cols.estado) + SEP + '"GESTIONADOS")' })
  }
  if (d.cols.accion) {
    cards.push({ label: 'POR INGRESAR', col1: 10, col2: 11, color: '#B91C1C', tint: '#FEE2E2',
      f: 'COUNTIF(' + cl(d.cols.accion) + rD + ':' + cl(d.cols.accion) + SEP + '"INGRESA")' })
  }

  for (var i = 0; i < cards.length; i++) {
    var ca = cards[i]
    var c2 = Math.min(ca.col2, lcF)
    if (ca.col1 > lcF || c2 < ca.col1) continue
    try {
      var nCols = c2 - ca.col1 + 1
      var rLab = sh.getRange(2, ca.col1, 1, nCols).merge()
      var labTxt = '● ' + ca.label
      var rtLab = SpreadsheetApp.newRichTextValue().setText(labTxt)
        .setTextStyle(0, 1, SpreadsheetApp.newTextStyle().setForegroundColor(ca.color).setBold(true)
          .setFontSize(8).setFontFamily(_UI.font).build())
        .setTextStyle(2, labTxt.length, SpreadsheetApp.newTextStyle().setForegroundColor('#64748B').setBold(true)
          .setFontSize(8).setFontFamily(_UI.font).build())
        .build()
      rLab.setRichTextValue(rtLab)
      rLab.setBackground(ca.tint).setHorizontalAlignment('center').setVerticalAlignment('middle')
      var rNum = sh.getRange(3, ca.col1, 1, nCols).merge()
      var aNum = sh.getRange(3, ca.col1)
      try { aNum.setFormula(ca.f) } catch (eF) { aNum.setValue(0) }
      aNum.setNumberFormat('0')
        .setFontFamily(_UI.font).setFontSize(20).setFontWeight('bold')
        .setFontColor(ca.color).setBackground('#FFFFFF')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
      sh.getRange(2, ca.col1, 2, nCols)
        .setBorder(true, true, true, true, false, false, ca.color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    } catch (eCard) {}
  }
  sh.setRowHeight(2, 20)
  sh.setRowHeight(3, 36)
}

// ─── FORMATO CONDICIONAL (colores vivos, se actualizan solos) ───────────────

function _ingAplicarCF(sh, d, desde, total) {
  var reglas = sh.getConditionalFormatRules()
  var fin = desde + total

  function limpiar(col) {
    if (!col) return
    var out = []
    for (var i = 0; i < reglas.length; i++) {
      var rngs = reglas[i].getRanges()
      var toca = false
      for (var j = 0; j < rngs.length; j++) {
        var r = rngs[j]
        if (r.getColumn() === col && r.getLastRow() >= desde && r.getRow() <= fin) { toca = true; break }
      }
      if (!toca) out.push(reglas[i])
    }
    reglas = out
  }

  if (d.cols.accion) {
    limpiar(d.cols.accion)
    var rA = sh.getRange(desde, d.cols.accion, total, 1)
    reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('INGRESA')
      .setBackground('#DCFCE7').setFontColor('#15803D').setBold(true).setRanges([rA]).build())
    reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('PENDIENTE')
      .setBackground('#FEF3C7').setFontColor('#A16207').setRanges([rA]).build())
    reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('NO INGRESA')
      .setBackground('#FEE2E2').setFontColor('#B91C1C').setRanges([rA]).build())
  }
  if (d.cols.estado) {
    limpiar(d.cols.estado)
    var rE = sh.getRange(desde, d.cols.estado, total, 1)
    reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('GESTIONADOS')
      .setBackground('#DCFCE7').setFontColor('#15803D').setRanges([rE]).build())
    reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('PENDIENTES')
      .setBackground('#FEF3C7').setFontColor('#A16207').setRanges([rE]).build())
    reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('EN ESPERA')
      .setBackground('#E0F2FE').setFontColor('#0369A1').setRanges([rE]).build())
  }
  sh.setConditionalFormatRules(reglas)
}

// ─── AUTOFILTRO (flechas en la fila de encabezados) ─────────────────────────

function _ingAplicarFiltro(sh, hr, lcF) {
  try {
    if (sh.getFilter()) return
    var fin = Math.max(sh.getLastRow(), hr + 1)
    sh.getRange(hr, 1, fin - hr + 1, lcF).createFilter()
  } catch (e) {}
}

// ─── SEGURIDAD: las filas sobre la cabecera no contienen datos ──────────────

function _ingFilasMetaSeguras(sh, d) {
  var hr = d.headerRow
  if (hr < 2) return false
  var lc = sh.getLastColumn()
  var metaTexto = ''
  for (var r = 1; r < hr; r++) {
    var vals = sh.getRange(r, 1, 1, lc).getValues()[0]
    for (var c = 0; c < vals.length; c++) {
      var v = vals[c]
      if (v == null || String(v).trim() === '') continue
      var s = String(v)
      metaTexto += s
      if (/[0-9]{1,8}[-]\s?[0-9kK]/.test(s)) return false
      if (/^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$/.test(s.trim())) return false
    }
  }
  // Solo migrar si las filas sobre la cabecera son el banner/título conocido del sistema
  return /INGRESO|ADMISI|OBSERVACION/.test(metaTexto.toUpperCase())
}

// ─── FORMATO DE UNA FILA EDITADA (llamado por onEdit) ───────────────────────

// nunca queden sin formato aunque estén fuera del rango pre-formateado.

function _ingFormatearFila(sh, row, numRows, d) {
  var lcF = d._lcF || 11
  var hr = d.headerRow
  var cols = d.cols
  var maxFilas = Math.min(numRows, 100)

  for (var i = 0; i < maxFilas; i++) {
    var r = row + i
    var bg = (r - hr) % 2 === 1 ? _UI.zebraBg[0] : _UI.zebraBg[1]
    var rng = sh.getRange(r, 1, 1, lcF)
    rng.setBackground(bg)
      .setFontFamily(_UI.font).setFontSize(10)
      .setVerticalAlignment('middle').setFontColor('#202124')
    if (cols.run) { sh.getRange(r, cols.run).setHorizontalAlignment('center'); sh.getRange(r, cols.run).setNumberFormat('@') }
    if (cols.estado) sh.getRange(r, cols.estado).setHorizontalAlignment('center')
    if (cols.fechaSolicitud) { sh.getRange(r, cols.fechaSolicitud).setHorizontalAlignment('center'); sh.getRange(r, cols.fechaSolicitud).setNumberFormat('dd/MM/yyyy') }
    if (cols.accion) {
      sh.getRange(r, cols.accion).setHorizontalAlignment('center')
      var bgAcc = _ingColorAccion(sh.getRange(r, cols.accion).getValue())
      if (bgAcc) sh.getRange(r, cols.accion).setBackground(bgAcc)
    }
    sh.getRange(r, 1, 1, lcF)
      .setBorder(true, true, true, true, true, true, _UI.border, SpreadsheetApp.BorderStyle.SOLID)
  }
}

// ─── CREAR HOJA INGRESOS (solo si no existe; nunca borra la existente) ───────

var _ING_HEADERS = ['ESTADO', 'FECHA SOLICITUD INGRESO', 'OBSERVACION', 'APELLIDO PATERNO',
  'APELLIDO MATERNO', 'NOMBRE', 'RUN', 'DIRECCION', 'TELEFONO', 'DERIVADO POR', 'ANTECEDENTES']

var _ING_VERSION = '5.5'

function _crearHojaIngresosSinConfirmar(ss) {
  var sh = ss.getSheetByName('Ingresos')
  if (sh) {
    ss.toast('La hoja INGRESOS ya existe: no se borra ni se recrea. Se aplica el nuevo diseño.', 'INGRESOS', 6)
    corregirYMejorarIngresos()
    return
  }
  sh = ss.insertSheet('Ingresos')
  var nCols = _ING_HEADERS.length

  if (sh.getMaxColumns() > nCols) sh.deleteColumns(nCols + 1, sh.getMaxColumns() - nCols)

  // Fila 1: título · Filas 2-3: panel de estadísticas (las completa _ingConstruirPanel)
  sh.getRange(1, 1, 1, nCols).merge()
  sh.getRange(1, 1).setValue('INGRESOS · LISTA DE ADMISIÓN')
  sh.getRange(1, 1).setBackground(_UI.hdrBg).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(15)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
  sh.setRowHeight(1, 36)

  sh.getRange(4, 1, 1, nCols).setValues([_ING_HEADERS])

  var d = _ingDetectarColumnas(sh)
  if (d) {
    _ingAplicarFormato(sh, d)
  } else {

    sh.setFrozenRows(4)
  sh.setTabColor(_UI.tabBW)
  }

  ss.toast('Hoja INGRESOS creada con el nuevo diseño (panel de estadísticas, filtros y colores automáticos)', 'INGRESOS', 6)
}

// ─── CORREGIR Y MEJORAR FORMATO (menú) ───────────────────────────────────────

// de texto). 2) Aplica el formato mejorado. Nunca borra ni elimina nada.

function _ingCorregirDatos(sh, d) {
  var hr = d.headerRow
  var lr = sh.getLastRow()
  if (lr <= hr) return 0
  var cols = d.cols
  var lc = sh.getLastColumn()

  var data = sh.getRange(hr + 1, 1, lr - hr, lc).getValues()
  var corregidas = 0

  var upCols = {}
  if (cols.nombre) upCols[cols.nombre] = true
  if (cols.apPaterno) upCols[cols.apPaterno] = true
  if (cols.apMaterno) upCols[cols.apMaterno] = true
  if (cols.accion) upCols[cols.accion] = true
  if (cols.estado) upCols[cols.estado] = true
  if (cols.derivado) upCols[cols.derivado] = true

  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < data[r].length; c++) {
      var v = data[r][c]
      if (v == null) continue
      var col = c + 1
      var nuevo = null

      if (typeof v === 'string') {

        var s = v.trim().replace(/\s+/g, ' ')
        if (col === cols.run && s !== '') {

          var f = formatearRUT(s)
          if (f !== s) nuevo = f
        } else if (upCols[col] && s !== '') {
          var u = s.toUpperCase()
          if (u !== s) nuevo = u
        } else if (col === cols.fechaSolicitud && s !== '') {

          var fd = _parseDate(s)
          if (fd) nuevo = fd
        } else if (s !== v) {
          nuevo = s
        }
      }

      if (nuevo !== null && nuevo !== v) {
        data[r][c] = nuevo
        corregidas++
      }
    }
  }

  if (corregidas > 0) {
    sh.getRange(hr + 1, 1, data.length, lc).setValues(data)
  }
  return corregidas
}

function corregirYMejorarIngresos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  try {
    _corregirIngresosImpl(ss)
    ss.toast('INGRESOS corregido · diseño v' + _ING_VERSION, 'INGRESOS', 6)
  } catch (e) {
    ss.toast('Error INGRESOS (v' + _ING_VERSION + '): ' + e.message, 'INGRESOS', 10)
    try {
      var uiE = SpreadsheetApp.getUi()
      uiE.alert('Error en INGRESOS (v' + _ING_VERSION + ')',
        'Mensaje: ' + e.message + '\n\nDetalle técnico (stack):\n' + (e.stack || '(sin stack)'),
        uiE.ButtonSet.OK)
    } catch (e2) {}
  }
}

function _corregirIngresosImpl(ss) {
  var sh = ss.getSheetByName('Ingresos')
  if (!sh) {
    _crearHojaIngresosSinConfirmar(ss)
    return
  }

  var d = _ingDetectarColumnas(sh)
  if (!d) {
    ss.toast('No encontré fila de encabezados (RUN/RUT · NOMBRE · OBSERVACION). No hice cambios. Envíame una captura de tu hoja para ajustar.', 'INGRESOS', 8)
    return
  }

  // 1) Corregir datos (solo normalización, nunca borra)
  var corregidas = _ingCorregirDatos(sh, d)

  // 2) Migrar layout antiguo (cabecera en fila 3) al nuevo (cabecera en fila 4,
  //    dejando espacio para el panel de estadísticas). Solo si las filas sobre
  //    la cabecera no contienen datos (seguridad).
  var migro = false
  if (d.headerRow === 3 && _ingFilasMetaSeguras(sh, d)) {
    try {
      sh.insertRowAfter(1)
      var d2 = _ingDetectarColumnas(sh)
      if (d2) { d = d2; migro = true }
    } catch (eMig) { /* sin panel: solo se formatea la tabla */ }
  }

  if (!d) return
  _ingAplicarFormato(sh, d)

  if (corregidas > 0) {
    ss.toast(corregidas + ' celda(s) corregida(s) y nuevo diseño aplicado' + (migro ? ' (panel de estadísticas agregado)' : ''), 'INGRESOS', 6)
    SpreadsheetApp.getUi().alert('Corrección y formato',
      corregidas + ' celda(s) fueron corregidas:\n· RUT: puntos, espacios o caracteres sueltos al final (ej. 1629792-5| → 1629792-5)\n· Espacios y mayúsculas en nombres/estados\n· Fechas escritas como texto → fecha real\n\nEl nuevo diseño quedó aplicado: título, panel de estadísticas en vivo, filtros y colores automáticos.',
      SpreadsheetApp.getUi().ButtonSet.OK)
  } else {
    ss.toast('Sin errores que corregir · nuevo diseño aplicado' + (migro ? ' (panel de estadísticas agregado)' : ''), 'INGRESOS', 6)
  }
}

// ─── OBSERVACION = INGRESA → ENVIAR A PACIENTES ──────────────────────────────
// Llamada por onEdit (04_Eventos.gs). Independiente del flujo del formulario.

function _ingresarDesdeLista(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Ingresos')
  var pac = ss.getSheetByName(HOJA_PAC)
  if (!sh || !pac) return 'no_estructura'

  var d = _ingDetectarColumnas(sh)
  if (!d || !d.cols.run || !d.cols.nombre) return 'no_estructura'
  var cols = d.cols
  var celAccion = sh.getRange(row, cols.accion || cols.estado || 1)

  var run = String(sh.getRange(row, cols.run).getValue() || '').trim()
  if (!/[0-9]/.test(run)) {
    celAccion.setNote('RUN no válido (sin dígitos): no se envió a Pacientes')
    ss.toast('RUN no válido en la fila ' + row + ': no se envió a Pacientes', 'INGRESOS', 5)
    return 'sin_rut'
  }
  var runN = formatearRUT(run)
  if (!/^\d{7,8}-[0-9K]$/.test(runN) || !_validarDigitoRUT(runN)) {
    celAccion.setNote('RUN inválido (dígito verificador incorrecto o incompleto): no se envió a Pacientes')
    ss.toast('RUN inválido en la fila ' + row + ' (dígito verificador): no se envió a Pacientes', 'INGRESOS', 5)
    return 'run_invalido'
  }

  var filaExistente = _buscarFilaPaciente(pac, runN)
  if (filaExistente > 0) {
    celAccion.setNote('Ya existe en Pacientes (fila ' + filaExistente + ') — no se duplicó')
    ss.toast('El RUT ya está en Pacientes (fila ' + filaExistente + '): no se duplicó', 'INGRESOS', 6)
    return 'duplicado'
  }

  // 2) Crear paciente nuevo (agregar fila al final de Pacientes, nunca borrar)
  try { _borrarFilasVacias(pac, 4) } catch (eC) {}
  var lr = pac.getLastRow()
  var fila = lr + 1
  var id = 1
  if (lr > 3) {
    var ids = pac.getRange(4, 1, lr - 3, 1).getValues()
    for (var i = 0; i < ids.length; i++) {
      var v = Number(ids[i][0])
      if (!isNaN(v) && v >= id) id = v + 1
    }
  }

  var nCols = _COLUMNAS._count || 112
  if (pac.getMaxColumns() < nCols) {
    celAccion.setNote('Pacientes no tiene las ' + nCols + ' columnas esperadas: usa el menú de corrección de la hoja Pacientes')
    ss.toast('Pacientes tiene ' + pac.getMaxColumns() + ' columnas (se esperan ' + nCols + '): no se envió a Pacientes', 'INGRESOS', 5)
    return 'no_estructura'
  }

  if (fila > pac.getMaxRows()) {
    pac.insertRowsAfter(pac.getMaxRows(), fila - pac.getMaxRows())
  }
  var newRow = []
  for (var c = 1; c <= nCols; c++) newRow.push('')
  pac.getRange(fila, 1, 1, nCols).setValues([newRow])

  var nombre = String(sh.getRange(row, cols.nombre).getValue() || '').trim().toUpperCase()
  var apP = cols.apPaterno ? String(sh.getRange(row, cols.apPaterno).getValue() || '').trim().toUpperCase() : ''
  var apM = cols.apMaterno ? String(sh.getRange(row, cols.apMaterno).getValue() || '').trim().toUpperCase() : ''
  var dir = cols.direccion ? String(sh.getRange(row, cols.direccion).getValue() || '').trim() : ''
  var tel = cols.telefono ? String(sh.getRange(row, cols.telefono).getValue() || '').trim() : ''
  var deriv = cols.derivado ? String(sh.getRange(row, cols.derivado).getValue() || '').trim() : ''
  var antec = cols.antecedentes ? String(sh.getRange(row, cols.antecedentes).getValue() || '').trim() : ''

  pac.getRange(fila, COL.ID).setValue(id)
  pac.getRange(fila, COL.VITAL).setValue('VIGENTE')
  pac.getRange(fila, COL.RUN).setValue(runN)
  if (nombre) pac.getRange(fila, COL.NOMBRE).setValue(nombre)
  if (apP) pac.getRange(fila, COL.APELLIDO).setValue(apP)
  if (apM) pac.getRange(fila, COL.APELLIDO2).setValue(apM)
  if (dir) pac.getRange(fila, COL.DIRECCION).setValue(dir.toUpperCase())
  if (tel) pac.getRange(fila, COL.TELEFONO).setValue(tel)

  pac.getRange(fila, COL.F_INGRESO_PADI).setValue(new Date())

  if (antec) pac.getRange(fila, COL.MORBILIDAD).setValue(antec.toUpperCase())

  var notaObs = '[INGRESO desde hoja INGRESOS] ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy')
  if (deriv) notaObs += '\n  Derivado por: ' + deriv
  pac.getRange(fila, COL.OBSERVACIONES).setValue(notaObs)

  try { pac.getRange(fila, COL.EDITOR).setValue(Session.getActiveUser().getEmail()) } catch (e) {}

  pac.getRange(fila, 1, 1, nCols)
    .setFontColor('#000000').setFontWeight('normal').setFontSize(9).setVerticalAlignment('middle')

  try { _actualizarEstadosFila(fila) } catch (e) {}

  celAccion.setNote(
    'Enviado a Pacientes · fila ' + fila + ' · ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'))

  SpreadsheetApp.flush()

  try {
    var _nPre = _compactarPacientes(pac) + _limpiarFilasVaciasLoop(pac, 4)
    _log(ss, 'Pacientes', '_ingresarDesdeLista', 'ok',
      'fila=' + fila + ' lr=' + lr + ' maxRows=' + pac.getMaxRows() + ' vaciasEliminadas=' + _nPre)
  } catch (eV) {}

  ss.toast('Ingreso confirmado → Pacientes fila ' + fila + ' (ID ' + id + ')', 'INGRESOS', 6)
  return 'enviado'
}

// ─── CONFIRMACIÓN AL PONER INGRESA (disparador instalable) ───────────────────

function configurarTriggerIngresos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  var triggers = ScriptApp.getProjectTriggers()
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEditIngresos') {
      ui.alert('Ya está activo', 'La confirmación de INGRESA ya está activa.', ui.ButtonSet.OK)
      return
    }
  }

  var resp = ui.alert('Activar confirmación de INGRESA',
    'Al activar, cada vez que escribas INGRESA en OBSERVACION aparecerá la ventana "¿Enviar a la hoja Pacientes?".\nAl confirmar: se verifica el RUT (sin duplicados) y la fila se elimina de INGRESOS.\n\n¿Activar?',
    ui.ButtonSet.YES_NO)
  if (resp !== ui.Button.YES) return

  ScriptApp.newTrigger('onEditIngresos').forSpreadsheet(ss).onEdit().create()
  ui.alert('Activado', 'Confirmación activada. Ya puedes escribir INGRESA en OBSERVACION.', ui.ButtonSet.OK)
}

function onEditIngresos(e) {
  try {
    _onEditIngresosImpl(e)
  } catch (errIE) {
    console.error('onEditIngresos: ' + errIE.message)
    _log(e && e.source, 'Ingresos', 'onEditIngresos', 'ERROR', errIE.message)
  }
}

function _onEditIngresosImpl(e) {
  if (!e) return
  var sh = e.range.getSheet()
  if (sh.getName() !== 'Ingresos') return

  var row = e.range.getRow()
  var numRows = e.range.getNumRows()
  if (numRows !== 1 || e.value == null) return

  var d = _ingDetectarColumnas(sh)
  var colAcc = d && (d.cols.accion || d.cols.estado)
  if (!d || row <= d.headerRow) return

  // Mayúsculas automáticas al escribir nombres/direcciones (como en Pacientes).
  var colEdit = e.range.getColumn()
  var upCols = {}
  if (d.cols.nombre) upCols[d.cols.nombre] = true
  if (d.cols.apPaterno) upCols[d.cols.apPaterno] = true
  if (d.cols.apMaterno) upCols[d.cols.apMaterno] = true
  if (d.cols.direccion) upCols[d.cols.direccion] = true
  if (d.cols.derivado) upCols[d.cols.derivado] = true
  if (upCols[colEdit]) {
    var nuevoUp = _mayusNombre(e.value)
    if (nuevoUp !== String(e.value).trim()) sh.getRange(row, colEdit).setValue(nuevoUp)
    return
  }

  if (!colAcc || colEdit !== colAcc) return
  if (!_ingEsIngresa(e.value)) return

  if (row > sh.getLastRow()) return
  var valActual = String(sh.getRange(row, colAcc).getValue() || '').trim().toUpperCase()
  if (valActual !== String(e.value).trim().toUpperCase()) return

  var nombre = d.cols.nombre ? String(sh.getRange(row, d.cols.nombre).getValue() || '').trim() : ''
  var runRaw = d.cols.run ? String(sh.getRange(row, d.cols.run).getValue() || '').trim() : ''
  var etiqueta = (nombre ? nombre.toUpperCase() : '') + (runRaw ? ' · ' + runRaw : '')

  var ui = SpreadsheetApp.getUi()
  var resp = ui.alert('Enviar a hoja Pacientes',
    '¿Enviar a la hoja Pacientes?\n' +
    (etiqueta ? 'Paciente: ' + etiqueta + '\n' : '') +
    'Al confirmar:\n· Se verificará el RUT (no se duplican pacientes).\n· La fila se ELIMINARÁ de INGRESOS.\n\n¿Continuar?',
    ui.ButtonSet.YES_NO)
  if (resp !== ui.Button.YES) return

  runRaw = String(sh.getRange(row, d.cols.run).getValue() || '').trim()
  if (!/[0-9]/.test(runRaw)) {
    ui.alert('RUN no válido',
      'La fila no tiene un RUN válido (sin dígitos).\nNo se envió ni se eliminó la fila.', ui.ButtonSet.OK)
    return
  }
  var runN = formatearRUT(runRaw)
  var pac = e.source.getSheetByName(HOJA_PAC)
  var existente = pac ? _buscarFilaPaciente(pac, runN) : -1
  if (existente > 0) {
    sh.getRange(row, colAcc).setNote('Ya existe en Pacientes (fila ' + existente + ') — no se envió ni se eliminó la fila')
    ui.alert('Paciente duplicado',
      'El RUT ' + runN + ' ya existe en Pacientes (fila ' + existente + ').\nNo se duplicó: la fila de INGRESOS se mantiene.', ui.ButtonSet.OK)
    return
  }

  var res = _ingresarDesdeLista(row)
  if (res !== 'enviado') {
    ui.alert('No se pudo enviar', 'No se eliminó la fila.', ui.ButtonSet.OK)
    return
  }

  // Eliminar SOLO la fila (nunca una hoja) y marcar para que el onEdit

  try {
    CacheService.getScriptCache().put('ING_DEL_' + row, '1', 30)
    sh.deleteRow(row)
  } catch (eDel) {
    ui.alert('Enviado (aviso)',
      'El paciente se agregó a Pacientes, pero no se pudo eliminar la fila (' + eDel.message + '). Revisa manualmente.', ui.ButtonSet.OK)
    return
  }

  try {
    if (sh.getLastRow() >= d.headerRow + 1) _ingReBandear(sh, d, row)
  } catch (eBand) {}

  // (solo filas, nunca hojas)
  var elimPend = 0
  try {
    var reps = _ingRepetidasMismoRut(sh, d, runN, [])
    if (reps.length > 0) {
      var respP = ui.alert('Filas repetidas del mismo paciente',
        'El paciente ya fue ingresado en Pacientes y hay ' + reps.length +
        ' fila(s) más con el mismo RUN en INGRESOS (repetidas, pendientes o marcadas NO INGRESA).\n\n¿Eliminarlas también?',
        ui.ButtonSet.YES_NO)
      if (respP === ui.Button.YES) {
        reps.sort(function(a, b) { return b - a })
        var minPend = reps[reps.length - 1]
        for (var pk = 0; pk < reps.length; pk++) {
          try { sh.deleteRow(reps[pk]); elimPend++ } catch (ePD) {}
        }
        try {
          if (sh.getLastRow() >= d.headerRow + 1) _ingReBandear(sh, d, minPend)
        } catch (ePB) {}
      }
    }
  } catch (ePend) {}

  ui.alert('Enviado',
    'El paciente se agregó a Pacientes y su fila fue eliminada de INGRESOS.' +
    (elimPend > 0 ? '\n' + elimPend + ' fila(s) repetida(s) del mismo paciente también fueron eliminadas.' : ''),
    ui.ButtonSet.OK)
}

function _ingReBandear(sh, d, desdeFila) {
  var hr = d.headerRow
  var lcF = d._lcF || 11
  var lr = sh.getLastRow()
  if (lr < desdeFila || lr < hr + 1) return

  var bgArr = []
  for (var r = desdeFila; r <= lr; r++) {
    var bg = (r - hr) % 2 === 1 ? _UI.zebraBg[0] : _UI.zebraBg[1]
    var rowArr = []
    for (var c = 0; c < lcF; c++) rowArr.push(bg)
    bgArr.push(rowArr)
  }
  sh.getRange(desdeFila, 1, lr - desdeFila + 1, lcF).setBackgrounds(bgArr)

  if (d.cols.accion) {
    var vals = sh.getRange(desdeFila, d.cols.accion, lr - desdeFila + 1, 1).getValues()
    var bgsAcc = sh.getRange(desdeFila, d.cols.accion, lr - desdeFila + 1, 1).getBackgrounds()
    for (var i = 0; i < vals.length; i++) {
      var bgA = _ingColorAccion(vals[i][0])
      if (bgA) bgsAcc[i][0] = bgA
    }
    sh.getRange(desdeFila, d.cols.accion, lr - desdeFila + 1, 1).setBackgrounds(bgsAcc)
  }
}

// ─── ENVÍO MASIVO: todas las filas con OBSERVACION = INGRESA ─────────────────

function enviarIngresasAPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Ingresos')
  var pac = ss.getSheetByName(HOJA_PAC)
  if (!sh || !pac) { ss.toast('Faltan las hojas INGRESOS o Pacientes', 'PADDS', 5); return }

  var d = _ingDetectarColumnas(sh)
  var colAcc = d && (d.cols.accion || d.cols.estado)
  if (!d || !colAcc || !d.cols.run) {
    ss.toast('No pude detectar la estructura de INGRESOS', 'INGRESOS', 6)
    return
  }

  var hr = d.headerRow
  var lr = sh.getLastRow()
  if (lr <= hr) { ss.toast('No hay datos que enviar', 'INGRESOS', 5); return }

  var vals = sh.getRange(hr + 1, colAcc, lr - hr, 1).getValues()
  var candidatas = []
  for (var i = 0; i < vals.length; i++) {
    if (_ingEsIngresa(vals[i][0])) candidatas.push(hr + 1 + i)
  }
  if (!candidatas.length) { ss.toast('No hay filas con OBSERVACION = INGRESA', 'INGRESOS', 5); return }

  var ui = SpreadsheetApp.getUi()
  var resp = ui.alert('Enviar a Pacientes',
    'Se enviarán ' + candidatas.length + ' fila(s) con OBSERVACION = INGRESA a Pacientes.\n' +
    'Las que ya existan no se duplican.\nLas filas enviadas se ELIMINARÁN de INGRESOS (se preguntará de nuevo antes de eliminar).\n\n¿Continuar?',
    ui.ButtonSet.YES_NO)
  if (resp !== ui.Button.YES) return

  var enviadas = 0, dup = 0, sinRut = 0
  var filasEnviadas = []
  var rutsEnviados = []
  for (var j = 0; j < candidatas.length; j++) {
    var res = _ingresarDesdeLista(candidatas[j])
    if (res === 'enviado') {
      enviadas++; filasEnviadas.push(candidatas[j])
      var rRaw = String(sh.getRange(candidatas[j], d.cols.run).getValue() || '').trim()
      if (rRaw && /[0-9]/.test(rRaw)) rutsEnviados.push(formatearRUT(rRaw).toUpperCase())
    }
    else if (res === 'duplicado') dup++
    else if (res === 'sin_rut') sinRut++
  }

  var msg = 'Enviadas: ' + enviadas + ' · ya existían: ' + dup
  if (sinRut) msg += ' · sin RUT válido: ' + sinRut
  ss.toast(msg, 'INGRESOS', 8)

  if (enviadas > 0) {
    var resp2 = ui.alert('Eliminar filas',
      enviadas + ' paciente(s) enviado(s). ¿Eliminar esas filas de INGRESOS?',
      ui.ButtonSet.YES_NO)
    if (resp2 === ui.Button.YES) {

      filasEnviadas.sort(function(a, b) { return b - a })
      for (var k = 0; k < filasEnviadas.length; k++) {
        try { sh.deleteRow(filasEnviadas[k]) } catch (eDel) {}
      }
      try {
        var d2 = _ingDetectarColumnas(sh)
        if (d2) _ingReBandear(sh, d2, filasEnviadas[filasEnviadas.length - 1])
      } catch (eBand) {}
      msg += ' · filas eliminadas: ' + enviadas
    }
  }

  // nunca hojas). Se excluyen las filas enviadas de esta misma pasada.
  var elimPend = 0
  try {
    var pendsTodas = []
    for (var pr = 0; pr < rutsEnviados.length; pr++) {
      var pend = _ingRepetidasMismoRut(sh, d, rutsEnviados[pr], filasEnviadas)
      for (var pi = 0; pi < pend.length; pi++) {
        if (pendsTodas.indexOf(pend[pi]) < 0) pendsTodas.push(pend[pi])
      }
    }
    if (pendsTodas.length > 0) {
      var respP = ui.alert('Filas repetidas del mismo paciente',
        'Los pacientes enviados tienen ' + pendsTodas.length +
        ' fila(s) más con el mismo RUN en INGRESOS (repetidas, pendientes o marcadas NO INGRESA).\n\n¿Eliminarlas también?',
        ui.ButtonSet.YES_NO)
      if (respP === ui.Button.YES) {
        pendsTodas.sort(function(a, b) { return b - a })
        var minPend = pendsTodas[pendsTodas.length - 1]
        for (var pk2 = 0; pk2 < pendsTodas.length; pk2++) {
          try { sh.deleteRow(pendsTodas[pk2]); elimPend++ } catch (ePD2) {}
        }
        try {
          var d3 = _ingDetectarColumnas(sh)
          if (d3) _ingReBandear(sh, d3, minPend)
        } catch (ePB2) {}
      }
    }
  } catch (ePend2) {}
  if (elimPend > 0) msg += ' · pendientes eliminadas: ' + elimPend

  ui.alert('Resultado', msg, ui.ButtonSet.OK)
}
