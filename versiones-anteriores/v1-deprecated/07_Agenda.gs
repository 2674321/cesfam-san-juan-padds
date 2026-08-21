// ─── AGENDA: GENERACIÓN + UTILIDADES ─────────────────────────────────────────

// ─── GENERAR SEMANA ──────────────────────────────────────────────────────────

function generarSemana() {
  var h = new Date()
  var d = h.getDay() === 0 ? -6 : 1 - h.getDay()
  crearSemana(new Date(h.getFullYear(), h.getMonth(), h.getDate() + d))
}

function generarTresSemanas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA)
  if (!sh) sh = ss.insertSheet(HOJA)

  var h = new Date()
  var d = h.getDay() === 0 ? -6 : 1 - h.getDay()
  var l = new Date(h.getFullYear(), h.getMonth(), h.getDate() + d)

  var maxCol = GI[2] + PC
  var lr = sh.getLastRow()
  var d2 = lr > 0 ? sh.getRange(1, 1, lr, maxCol).getValues() : []
  var freeRow = 1
  while (true) {
    if (freeRow > d2.length) break
    var rIdx = freeRow - 1
    if (rIdx < d2.length) {
      var ci0 = GI[0] - 1
      if (!d2[rIdx][ci0] || String(d2[rIdx][ci0]).trim() === '') break
      if ((!d2[rIdx][GI[1]-1] || String(d2[rIdx][GI[1]-1]).trim() === '') &&
          (!d2[rIdx][GI[2]-1] || String(d2[rIdx][GI[2]-1]).trim() === '')) break
    }
    freeRow += _calcularRS() + GF
  }

  for (var i = 0; i < 3; i++) {
    var f = new Date(l.getFullYear(), l.getMonth(), l.getDate() + i * 7)
    limpiarBloque(sh, freeRow, GI[i])
    ponerSemana(sh, f, freeRow, GI[i])
  }
  ss.toast('3 semanas creadas en Agenda', '', 4)
}

function generarSemanaEspecifica() {
  var ui = SpreadsheetApp.getUi()
  var r = ui.prompt('Fecha del LUNES', 'DD/MM/AAAA:', ui.ButtonSet.OK_CANCEL)
  if (r.getSelectedButton() !== ui.Button.OK) return
  var p = r.getResponseText().trim().split('/')
  if (p.length !== 3) { ui.alert('Formato inválido. Usa DD/MM/AAAA.'); return }
  var f = new Date(+p[2], +p[1]-1, +p[0])
  if (isNaN(f)) { ui.alert('Fecha inválida.'); return }
  if (f.getDay() !== 1) { ui.alert('La fecha debe ser un LUNES.'); return }
  crearSemana(f)
}

function crearSemana(lunes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA)
  if (!sh) sh = ss.insertSheet(HOJA)
  var pos = siguienteBloque(sh)
  ponerSemana(sh, lunes, pos.row, pos.col)
  ss.toast('Semana del ' + fmtFecha(lunes) + ' creada', '', 3)
}

function siguienteBloque(sh) {
  var maxCol = GI[2] + PC
  var lr = sh.getLastRow()
  if (lr < 1) return { row: 1, col: GI[0] }
  var d = sh.getRange(1, 1, lr, maxCol).getValues()
  var row = 0
  while (row < d.length) {
    for (var bi = 0; bi < 3; bi++) {
      var ci = GI[bi] - 1
      if (ci >= d[row].length || !d[row][ci] || String(d[row][ci]).trim() === '') {
        return { row: row + 1, col: GI[bi] }
      }
    }
    row += _calcularRS() + GF
  }
  return { row: row + 1, col: GI[0] }
}

function _getWeekTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var tpl = ss.getSheetByName('_PlantillaSemana')
  var expectedRS = _calcularRS()
  if (tpl) {
    if (tpl.getMaxRows() === expectedRS) return tpl
    ss.deleteSheet(tpl)
  }
  tpl = ss.insertSheet('_PlantillaSemana')
  tpl.hideSheet()

  var C = PC + 1
  var RS = _calcularRS()
  var vals = [], bgs = []
  for (var _vi = 0; _vi < RS; _vi++) {
    var _rowV = [], _rowB = []
    for (var _ci = 0; _ci < C; _ci++) { _rowV.push(''); _rowB.push('#ffffff') }
    vals.push(_rowV); bgs.push(_rowB)
  }

  var fixLunes = new Date(2026, 0, 5)
  var fe2 = new Date(fixLunes.getFullYear(), fixLunes.getMonth(), fixLunes.getDate() + 4)
  vals[0][0] = 'Semana ' + fmtFecha(fixLunes) + ' — ' + fmtFecha(fe2)
  for (var c = 0; c < 5; c++) bgs[0][c] = '#1a3c5e'
  vals[0][5] = '✕'; bgs[0][5] = '#e53935'
  vals[0][6] = '📦'; bgs[0][6] = '#f39c12'

  var _dCols = ['#e3f2fd','#e8f5e9','#fff8e1','#fce4ec','#f3e5f5']
  var rowPos = 1

  for (var d = 0; d < 5; d++) {
    var slots = SLOTS[d] || SLOTS_DEFAULT
    var b = rowPos
    var fd = new Date(fixLunes.getFullYear(), fixLunes.getMonth(), fixLunes.getDate() + d)

    // Day header
    vals[b][0] = DIAS[d] + '  ' + fmtFecha(fd)
    for (var c = 0; c < 6; c++) bgs[b][c] = _dCols[d]

    // Sub-header
    for (var c = 0; c < 6; c++) { vals[b+1][c] = SUBS[c]; bgs[b+1][c] = S.sub.bg[c] }

    // Slots
    for (var i = 0; i < slots.length; i++) {
      var slot = slots[i]
      var r = b + 2 + i

      if (slot[1] === 'COLACION') {
        vals[r][0] = 'C O L A C I O N'
        for (var c = 0; c < 6; c++) bgs[r][c] = S.slots.colacion.bg
      } else {
        vals[r][0] = slot[0]; vals[r][1] = slot[1]
        var st = S.slots[slot[1]] || { bg: '#fff', fg: '#333' }
        bgs[r][1] = st.bg
        var alt = i % 2 === 0 ? '#f8f9fb' : '#ffffff'
        for (var c = 2; c < 6; c++) bgs[r][c] = alt
      }
    }

    rowPos += 3 + slots.length
  }

  var RS = _calcularRS()
  var block = tpl.getRange(1, 1, RS, C)
  block.setValues(vals)
  block.setBackgrounds(bgs)
  block.setFontFamily('Arial').setFontSize(11).setHorizontalAlignment('left').setVerticalAlignment('middle')

  tpl.getRange(1, 1, 1, 5).merge()
  var pr = tpl.getRange(1, 1, 1, 5)
  pr.setFontSize(11).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, false)
  tpl.setRowHeight(1, 28)

  tpl.getRange(1, 6).setFontSize(14).setFontWeight('bold').setBackground('#e53935')
    .setFontColor('#fff').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true)
    .setNote('Eliminar: selecciona la celda y presiona Delete')
  tpl.getRange(1, 7).setFontSize(14).setBackground('#f39c12').setFontColor('#fff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true)
    .setNote('Limpiar: selecciona la celda y presiona Delete')

  var valTipos = SpreadsheetApp.newDataValidation()
    .requireValueInList(['VDI', 'abreviadas', 'REGISTRO'], true).setAllowInvalid(true).build()

  rowPos = 1
  for (var d = 0; d < 5; d++) {
    var slots = SLOTS[d] || SLOTS_DEFAULT
    var b = rowPos

    // Day header merge + style
    tpl.getRange(b, 1, 1, 6).merge()
    tpl.getRange(b, 1, 1, 6).setFontSize(12).setFontWeight('bold')
      .setFontColor('#1a3c5e').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false).setWrap(true)
    tpl.setRowHeight(b, 30)

    // Sub-header
    tpl.getRange(b + 1, 1, 1, 6).setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setBorder(true, true, true, true, true, true)
    tpl.setRowHeight(b + 1, 22)

    // Individual slot rows
    for (var i = 0; i < slots.length; i++) {
      var slot = slots[i]
      var r = b + 2 + i
      var slotType = slot[1]

      if (slotType === 'COLACION') {
        tpl.getRange(r, 1, 1, 6).merge()
        tpl.getRange(r, 1, 1, 6).setFontWeight('bold').setFontSize(10)
          .setFontColor(S.slots.colacion.fg).setHorizontalAlignment('center')
          .setVerticalAlignment('middle').setBorder(true, true, true, true, true, true).setWrap(true)
        tpl.setRowHeight(r, 28)
        tpl.getRange(r, 2).clearDataValidations()
      } else {
        tpl.getRange(r, 2).setFontWeight('bold').setFontSize(10)
          .setDataValidation(valTipos)
        tpl.getRange(r, 2).setFontColor(S.slots[slotType] ? S.slots[slotType].fg : '#333')
        if (i < Math.floor(slots.length / 2)) tpl.getRange(r, 2).setFontStyle('italic')
        tpl.getRange(r, 1, 1, 6).setBorder(true, true, true, true, false, false)
        tpl.setRowHeight(r, 26)
      }
    }

    // Separator row
    tpl.setRowHeight(b + 2 + slots.length, 6)
    tpl.getRange(b + 2 + slots.length, 1, 1, 6)
      .setBorder(false, false, true, false, false, false).setBackground('#e0e0e0')

    rowPos += 3 + slots.length
  }

  for (var i = 0; i < 6; i++) tpl.setColumnWidth(i + 1, ANCHOS[i])
  if (7 <= tpl.getMaxColumns()) tpl.setColumnWidth(7, 40)

  // Tooltips for sub-header columns
  var SUB_TOOLTIPS = [
    'Horario del bloque (formato HH:MM). Editar si se requiere ajustar.',
    'Tipo de atención: VDI (1h), abreviadas (15 min) o REGISTRO. Seleccionar de la lista.',
    'Nombre completo del paciente agendado.',
    'RUN del paciente (sin puntos, con guion).',
    'Dirección del paciente.',
    'Observaciones o notas de la atención.'
  ]
  rowPos = 1
  for (var d = 0; d < 5; d++) {
    var b = rowPos
    for (var c = 0; c < 6; c++) {
      tpl.getRange(b + 1, c + 1).setNote(SUB_TOOLTIPS[c])
    }
    var slots = SLOTS[d] || SLOTS_DEFAULT
    rowPos += 3 + slots.length
  }

  // Format hora/RUT columns via RangeList
  var horaRngs = [], rutRngs = []
  rowPos = 1
  for (var d = 0; d < 5; d++) {
    var slots = SLOTS[d] || SLOTS_DEFAULT
    var b = rowPos
    if (slots.length > 0) {
      horaRngs.push('A' + (b+2) + ':A' + (b+1+slots.length))
      rutRngs.push('D' + (b+2) + ':D' + (b+1+slots.length))
    }
    rowPos += 3 + slots.length
  }
  if (horaRngs.length) {
    tpl.getRangeList(horaRngs).setFontFamily('Consolas, monospace').setFontSize(10)
      .setFontWeight('bold').setHorizontalAlignment('center')
    tpl.getRangeList(rutRngs).setNumberFormat('@')
  }

  return tpl
}

function ponerSemana(sh, lunes, row, col) {
  var C = PC + 1
  var tpl = _getWeekTemplate()

  // Copy template to target position
  var RS = _calcularRS()
  var srcRange = tpl.getRange(1, 1, RS, C)
  var dstRange = sh.getRange(row, col, RS, C)
  srcRange.copyTo(dstRange)

  // Update date-specific content
  var fe = new Date(lunes)
  var fe2 = new Date(lunes.getFullYear(), lunes.getMonth(), lunes.getDate() + 4)
  var semTxt = 'Semana ' + fmtFecha(fe) + ' — ' + fmtFecha(fe2)

  var updates = []
  updates.push({ r: row, c: col, v: semTxt })
  var dayRow = row + 1
  for (var d = 0; d < 5; d++) {
    var fd = new Date(lunes.getFullYear(), lunes.getMonth(), lunes.getDate() + d)
    updates.push({ r: dayRow, c: col, v: DIAS[d] + '  ' + fmtFecha(fd) })
    var slots = SLOTS[d] || SLOTS_DEFAULT
    dayRow += 3 + slots.length
  }
  for (var u = 0; u < updates.length; u++) {
    sh.getRange(updates[u].r, updates[u].c).setValue(updates[u].v)
  }

  for (var i = 0; i < 6; i++) sh.setColumnWidth(col + i, ANCHOS[i])
  if (col + 6 <= sh.getMaxColumns()) sh.setColumnWidth(col + 6, 40)
}

// ─── ELIMINAR ────────────────────────────────────────────────────────────────

function eliminarSemana() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA)
  if (!sh) return
  var a = sh.getActiveRange()
  if (!a) return
  var row = a.getRow(), col = a.getColumn()
  var bc = -1
  for (var i = 0; i < 3; i++) {
    if (col >= GI[i] && col < GI[i] + PC) { bc = i; break }
  }
  if (bc < 0) { SpreadsheetApp.getUi().alert('Selecciona primero una celda dentro de la agenda.'); return }

  var inicio = _findBloqueInicio(sh, row, GI[bc])
  if (inicio < 0) { SpreadsheetApp.getUi().alert('No se pudo encontrar el bloque de semana.'); return }

  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Eliminar semana',
    '¿Eliminar toda esta semana? Se borrarán todos los datos.',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return

  limpiarBloque(sh, inicio, GI[bc])
  ss.toast('Semana eliminada', '', 2)
}

// ─── LIMPIAR ─────────────────────────────────────────────────────────────────

function limpiarSemanasPasadas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA)
  if (!sh) return
  var maxCol = GI[2] + 1
  var lr = sh.getLastRow()
  if (lr < 1) { ss.toast('Agenda vacía, no hay semanas que limpiar', '', 2); return }
  var d = sh.getRange(1, 1, lr, maxCol).getValues()
  var hoy = new Date()
  hoy.setHours(0, 0, 0, 0)

  var lunesHoy
  {
    var h2 = new Date(hoy)
    lunesHoy = new Date(h2.getFullYear(), h2.getMonth(), h2.getDate() - (h2.getDay() === 0 ? 6 : h2.getDay() - 1))
  }

  var aLimpiar = []
  for (var r = 0; r < d.length; r++) {
    for (var bi = 0; bi < 3; bi++) {
      var ci = GI[bi] - 1
      if (ci >= d[r].length) continue
      var a = String(d[r][ci] || '').trim()
      if ((a.indexOf('Semana ') === 0 || a.indexOf('👤') === 0 || a.indexOf('Profesional') === 0) && r + 1 < d.length) {
        var rawDateStr = String(d[r + 1][ci] || '').trim()
        // Extract DD/MM/AAAA from "LUNES  DD/MM/AAAA" or similar
        var dateMatch = rawDateStr.match(/(\d{2})\/(\d{2})\/(\d{4})/)
        var fe = dateMatch ? new Date(+dateMatch[3], +dateMatch[2] - 1, +dateMatch[1]) : null
        if (fe && fe < lunesHoy) aLimpiar.push({ fila: r + 1, col: GI[bi] })
      }
    }
  }

  if (aLimpiar.length === 0) {
    ss.toast('No hay semanas anteriores a hoy para limpiar', '', 3)
    return
  }

  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Limpiar ' + aLimpiar.length + ' semana(s)',
    '¿Limpiar ' + aLimpiar.length + ' semana(s) anteriores a ' + fmtFecha(lunesHoy) + '?',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return

  ss.toast('Limpiando ' + aLimpiar.length + ' semanas…', '', 1)
  for (var _ai = 0; _ai < aLimpiar.length; _ai++) {
    limpiarBloque(sh, aLimpiar[_ai].fila, aLimpiar[_ai].col)
  }
  ss.toast(aLimpiar.length + ' semanas anteriores limpiadas', '', 4)
}

// ─── NAVEGACIÓN ──────────────────────────────────────────────────────────────

function irAHoy() {
  navegarA(fmtFecha(new Date()))
}

function irAFecha() {
  var ui = SpreadsheetApp.getUi()
  var r = ui.prompt('Buscar fecha', 'DD/MM/AAAA:', ui.ButtonSet.OK_CANCEL)
  if (r.getSelectedButton() !== ui.Button.OK) return
  navegarA(r.getResponseText().trim())
}

function navegarA(s) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA)
  if (!sh) { SpreadsheetApp.getUi().alert('Primero debes crear una semana en la Agenda.'); return }
  var maxCol = GI[2] + 1
  var d = sh.getRange(1, 1, sh.getLastRow(), maxCol).getValues()
  for (var r = 0; r < d.length; r++) {
    for (var ci = 0; ci < 3; ci++) {
      var c = GI[ci] - 1
      if (c < d[r].length && d[r][c] && _fmtCeldaParaComparar(d[r][c]).indexOf(s) === 0) {
        sh.setActiveRange(sh.getRange(r + 1, GI[ci]))
        ss.toast(s, '', 2); return
      }
    }
  }
  for (var r = 0; r < d.length; r++) {
    for (var ci = 0; ci < 3; ci++) {
      var c = GI[ci] - 1
      if (c < d[r].length && d[r][c] && _fmtCeldaParaComparar(d[r][c]).indexOf(s) !== -1) {
        sh.setActiveRange(sh.getRange(r + 1, GI[ci]))
        ss.toast(s, '', 2); return
      }
    }
  }
  SpreadsheetApp.getUi().alert('No encontrado', s + ' no encontrado en la agenda.')
}

function buscarPaciente() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()
  var r = ui.prompt('🔎 Buscar paciente', 'Nombre o parte:', ui.ButtonSet.OK_CANCEL)
  if (r.getSelectedButton() !== ui.Button.OK) return
  var q = r.getResponseText().trim().toLowerCase()
  if (!q) return
  var sh = ss.getSheetByName(HOJA)
  if (!sh) return
  var maxCol = GI[2] + PC + 1
  var d = sh.getRange(1, 1, sh.getLastRow(), maxCol).getValues()
  var resultados = []
  for (var ri = 0; ri < d.length; ri++) {
    for (var bi = 0; bi < 3; bi++) {
      var ci = GI[bi] + 1
      if (ci < d[ri].length) {
        var nom = String(d[ri][ci] || '').trim()
        if (nom.toLowerCase().indexOf(q) !== -1) {
          resultados.push({ row: ri + 1, col: ci + 1, nombre: nom, bi: bi })
        }
      }
    }
  }
  if (resultados.length === 0) { ui.alert('Sin resultados', 'No se encontró "' + q + '" en la agenda.'); return }
  if (resultados.length === 1) {
    sh.setActiveRange(sh.getRange(resultados[0].row, resultados[0].col))
    sh.getRange(resultados[0].row, resultados[0].col - 2, 1, 6)
      .setBorder(true, true, true, true, true, true)
    ss.toast(resultados[0].nombre, '', 4); return
  }
  var lista = resultados.map(function(x, i) { return (i + 1) + '. ' + x.nombre + ' (fila ' + x.row + ')' }).join('\n')
  var sel = ui.prompt(resultados.length + ' resultados:\n' + lista + '\n\nN° para ir:',
    '1-' + resultados.length, ui.ButtonSet.OK_CANCEL)
  if (sel.getSelectedButton() !== ui.Button.OK) return
  var idx = parseInt(sel.getResponseText().trim()) - 1
  if (isNaN(idx) || idx < 0 || idx >= resultados.length) { ui.alert('Número inválido', 'Ingresa un número entre 1 y ' + resultados.length + '.'); return }
  sh.setActiveRange(sh.getRange(resultados[idx].row, resultados[idx].col))
  sh.getRange(resultados[idx].row, resultados[idx].col - 2, 1, 6)
    .setBorder(true, true, true, true, true, true)
  ss.toast(resultados[idx].nombre, '', 4)
}

// ─── RESUMEN ─────────────────────────────────────────────────────────────────

function resumen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Generando resumen…', '', 1)
  var sh = ss.getSheetByName(HOJA)
  if (!sh) return
  var maxCol = GI[2] + PC
  var lr = sh.getLastRow()
  if (lr < 1) return
  var d = sh.getRange(1, 1, lr, maxCol).getValues()
  var sem = 0, vdi = 0, abre = 0, lle = 0, col = 0
  var dcPorGI = [0, 0, 0]
  for (var _ri = 0; _ri < d.length; _ri++) {
    var row = d[_ri]
    for (var _gii = 0; _gii < GI.length; _gii++) {
      var ci = GI[_gii]
      if (ci < row.length) {
        var v = String(row[ci] || '').trim()
        if (v.indexOf('Semana') === 0) sem++
      }
      if (ci + 1 < row.length) {
        var t = String(row[ci + 1] || '').trim().toLowerCase()
        if (t === 'vdi') vdi++
        else if (t === 'abreviadas') abre++
        else if (t.indexOf('colaci') !== -1) col++
      }
      var ocCol = ci + 2
      if (ocCol < row.length && row[ocCol] && String(row[ocCol]).trim()) lle++
      // Días con datos: detectar fecha en col ci, ver si hay nombre en ci+2
      var c = ci - 1
      if (c < row.length && /^\d{2}\/\d{2}\/\d{4}/.test(String(row[c] || ''))) {
        if (ocCol < row.length && row[ocCol] && String(row[ocCol]).trim()) dcPorGI[_gii]++
      }
    }
  }
  var dc = dcPorGI[0] + dcPorGI[1] + dcPorGI[2]
  var oc = sem > 0 ? Math.round(lle / (sem * 11) * 100) : 0
  var _sep = ''
  for (var _si = 0; _si < 20; _si++) _sep += '\u2500'
  SpreadsheetApp.getUi().alert(
'📊 RESUMEN AGENDA\n' + _sep +
'\n📅 Semanas: ' + sem +
'\n✏️ Días con datos: ' + dc +
'\n🏠 VDI: ' + vdi +
'\n⚡ Abreviadas: ' + abre +
'\n🍽 Colaciones: ' + col +
'\n📝 Ocupados: ' + lle +
'\n📈 Ocupación: ' + oc + '%')
}


