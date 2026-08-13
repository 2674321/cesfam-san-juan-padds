// ════════════════════════════════════════════════════════
// 04_Eventos.gs │ disparadores onEdit/onSelectionChange + pipeline de fila

// ─────────────────────────────────────────────────────────
// ─── EVENTOS: onEdit + procesamiento de filas ──────────────────────────────

var _deleting = false
var _processingEdit = false

// El toggle multiselect escribe el valor programáticamente y el eco de esa
// edición llega como OTRA invocación de onEdit (las variables globales NO
// persisten entre invocaciones). Para ignorar el eco se usa CacheService:
// _toggleMultiselect arma la clave antes del setValue y aquí se consume.
function _esEchoMultiselect(col, row, valor) {
  try {
    var c = CacheService.getScriptCache()
    if (!c) return false
    var k = '_skip_ms_' + col + '_' + row
    var guardado = c.get(k)
    if (guardado !== null && String(valor || '').trim() === guardado) {
      c.remove(k)
      return true
    }
  } catch (e) {}
  return false
}

function onSelectionChange(e) {
  if (!e) return
  var sh = e.range.getSheet()
  var name = sh.getName()
  if (name !== HOJA_PAC && name !== HOJA_FORM) return
  var row = e.range.getRow()
  if (name === HOJA_PAC && row < 4) return
  if (name === HOJA_FORM && row < 6) return

  var cache = CacheService.getScriptCache()
  var last = cache.get('_hlr')
  if (last && Number(last) === row) return
  cache.put('_hlr', String(row), 10)

  try {
    var helper = e.source.getSheetByName('_Resalte')
    if (!helper) return
    helper.getRange(1, 1, 1, 2).setValues([[name, row]])
  } catch(_e) {}
}

function onEdit(e) {
  if (!e) return
  if (_processingEdit) return
  _processingEdit = true
  var _ss = e.source

  try {
    var sh = e.range.getSheet()
    var shName = sh.getName()
    var row = e.range.getRow()
    var c1 = e.range.getColumn()
    var c2 = e.range.getLastColumn()
    var numRows = e.range.getNumRows()

    if (shName === HOJA_FORM && c1 === 3 && row >= 6) {
      var val = String(e.value || '').trim()
      if (_ESTADO_CSS[val]) {
        _colorEstado(sh, row, val, sh.getLastColumn())
      }
      return
    }

    if (shName === HOJA_FORM && row >= 6 && c1 === 7 && c2 === 7 && e.value !== null && e.value !== undefined) {
      var _rn = String(e.value).replace(/[. ]/g, '').toUpperCase()
      if (_rn && _rn !== String(e.value)) {
        e.range.setValue(_rn)
        return
      }
      var _notaForm = _notaRUN(_rn)
      e.range.setNote(_notaForm)
      var _bgForm = _FORM_SEC_TINTS[7] || '#ffffff'
      if (row % 2 === 0) _bgForm = _ajustarHex(_bgForm, -8, -8, -8)
      e.range.setBackground(_notaForm ? '#FEE2E2' : _bgForm)
        .setFontColor(_notaForm ? '#B91C1C' : '#212121')
      return
    }

    if (shName === 'Parámetros' && row >= 4) {
      _paramCache = null
      try { CacheService.getScriptCache().remove('_pac_params') } catch(e) {}
      return
    }

    if (shName === HOJA_PAC && row >= 4 && c1 === c2 && PAC_VALIDACIONES[c1]) {
      if (PAC_MULTISELECT.indexOf(c1) >= 0) {
        _pintarCeldaOpcion(sh, row, c1)
        if (_esEchoMultiselect(c1, row, e.value)) return
        var _tog = _toggleMultiselect(sh, row, c1, e.value)
        _pintarCeldaOpcion(sh, row, c1)
        if (_tog) return
      } else {
        _pintarCeldaOpcion(sh, row, c1)
      }
    }

    // 📥 Ingresos → "Activar confirmación de INGRESA" (ver 11_Ingresos.gs).
    // Nunca se borran hojas.
    if (shName === 'Ingresos') {
      var _ingD = _ingDetectarColumnas(sh)
      if (_ingD && row > _ingD.headerRow) {
        try {
          var _ingCacheM = CacheService.getScriptCache()
          if (_ingCacheM && _ingCacheM.get('ING_DEL_' + row)) { return }
        } catch (eCache) {}
        // las filas nuevas nunca quedan sin formato
        try { _ingFormatearFila(sh, row, numRows, _ingD) } catch (errFmt) {
          console.error('onEdit(INGRESOS fmt): ' + errFmt.message)
        }
      }
      return
    }

    if (shName === HOJA_PAC && row === 2 && c1 === 2) {
      aplicarFiltroBusqueda(String(e.value || '').trim())
      return
    }

    if (shName === HOJA_PAC && row === 2 && c1 === 6) {
      aplicarFiltroSecciones(sh)
      return
    }

    if (shName === HOJA_PAC && row >= 4) {
      var _hasRelevant = false
      for (var cc = c1; cc <= c2; cc++) { if (_RELEVANT_COLS[cc]) { _hasRelevant = true; break } }
      if (!_hasRelevant) {

        try {
          var _lcAge = sh.getLastColumn()
          for (var _ar = row; _ar < row + numRows; _ar++) {
            if (_ar < 4) continue
            var _rowAge = sh.getRange(_ar, 1, 1, _lcAge).getValues()[0]
            var _nacA = _parseDate(_rowAge[8])
            if (_nacA) {
              var _edadA = _calcularEdad(_nacA)
              if (String(_rowAge[COL.EDAD_USUARIO - 1]) !== String(_edadA)) {
                sh.getRange(_ar, COL.EDAD_USUARIO).setValue(_edadA)
              }
            }
          }
        } catch (_eAge) {}
        return
      }

      var lock = LockService.getScriptLock()
      try { if (!lock.tryLock(5000)) return } catch(ee) { return }

      try {
        var _params = leerParametros()
        var _diasAviso = _params['DIAS_AVISO'] || 0
        var _lc = sh.getLastColumn()
        var _editor = e.user ? e.user.getEmail() : Session.getActiveUser().getEmail()
        var _M = _precalcularMeses(_params)

        for (var r = row; r < row + numRows; r++) {
          if (r < 4) continue
          _processPacientesRow(r, sh, c1, c2, _lc, _params, _diasAviso, _editor, _M)
        }
      } finally {
        lock.releaseLock()
      }
      return
    }

    if (shName !== HOJA && !_esHojaAgenda(sh)) return

    var colAg = e.range.getColumn()
    var valAgenda = String(e.value || '').trim()
    var oldVal = String(e.oldValue || '').trim()

    if (!_deleting && valAgenda === '' && (oldVal === '✕' || oldVal === '📦')) {
      var delIdx = oldVal === '✕' ? 0 : 1
      for (var i = 0; i < 3; i++) {
        if (colAg === GI[i] + [5, 6][delIdx]) {
          _deleting = true
          try { e.range.setValue(oldVal) } catch(ee) {}

          var inicio = _findBloqueInicio(sh, row, GI[i])
          if (inicio < 0) { _deleting = false; _ss.toast('No se encontró la cabecera de la semana', 'Agenda', 3); return }

          var ui = SpreadsheetApp.getUi()
          if (ui.alert('Limpiar semana', '¿Limpiar toda esta semana?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
            limpiarBloque(sh, inicio, GI[i])
            _ss.toast('Semana limpiada', 'Agenda', 2)
          }
          _deleting = false
          return
        }
      }
    }

    if (colAg % (PC + GC) === 3 && valAgenda.length > 0) {
      for (var i = 0; i < 3; i++) {
        if (colAg >= GI[i] && colAg < GI[i] + PC) {
          sh.getRange(row, GI[i], 1, 6).setBorder(true, true, true, true, true, true)
        }
      }
    }
  } catch(ee) {
    _ss.toast('Error en onEdit: ' + ee.message, 'PADDS', 5)
  } finally {
    _processingEdit = false
  }
}

// ─── MULTISELECCIÓN (columnas en PAC_MULTISELECT, ej. AYUDAS TECNICAS y
// RIESGO PIE DM) ─────────────────────────────────────────────────────────────
// Cada vez que se elige una opción del dropdown se AGREGA a la lista separada
// por comas ("SILLA RUEDAS, CAMA"). Elegir un valor ya presente NO cambia nada
// (nunca elimina): así ni el eco de la edición ni una re-elección borran datos.
// N/A es excluyente: elegirlo reemplaza la lista; elegir otra opción lo quita.

// Lógica pura del multiselect: calcula el nuevo valor de la celda.
// Devuelve null si el valor no cambió (opción desconocida o ya presente).
function _calcularMultiselect(actual, valor, lista) {
  var nuevo = String(valor == null ? '' : valor).trim()
  if (!nuevo || !lista || !lista.length) return null

  var clave = _quitarAcentos(nuevo).replace(/\s+/g, ' ').toUpperCase()
  var item = null
  for (var i = 0; i < lista.length; i++) {
    if (_quitarAcentos(String(lista[i])).replace(/\s+/g, ' ').toUpperCase() === clave) {
      item = String(lista[i])
      break
    }
  }
  if (!item) return null

  var actualTxt = String(actual == null ? '' : actual).trim()
  var partes = actualTxt ? actualTxt.split(',').map(function(s) { return s.trim() }).filter(function(s) { return s }) : []
  var idx = -1
  for (var p = 0; p < partes.length; p++) {
    if (_quitarAcentos(partes[p]).replace(/\s+/g, ' ').toUpperCase() === clave) { idx = p; break }
  }

  if (item === 'N/A') {
    if (idx >= 0) return null
    partes = ['N/A']
  } else {
    if (idx >= 0) return null
    partes = partes.filter(function(s) { return s !== 'N/A' })
    partes.push(item)
  }

  partes.sort(function(a, b) {
    var ia = lista.indexOf(a), ib = lista.indexOf(b)
    if (ia < 0) ia = lista.length
    if (ib < 0) ib = lista.length
    return ia - ib
  })

  var nuevoValor = partes.join(', ')
  return nuevoValor === actualTxt ? null : nuevoValor
}

function _toggleMultiselect(sh, row, col, valor) {
  var lista = PAC_VALIDACIONES[col]
  if (!lista) return false

  var actual = String(sh.getRange(row, col).getValue() || '').trim()
  var nuevoValor = _calcularMultiselect(actual, valor, lista)
  if (nuevoValor === null) return false

  try {
    CacheService.getScriptCache().put('_skip_ms_' + col + '_' + row, nuevoValor, 60)
  } catch (eC) {}
  try {
    sh.getRange(row, col).setValue(nuevoValor)
  } catch (eSet) {
    try { CacheService.getScriptCache().remove('_skip_ms_' + col + '_' + row) } catch (eR2) {}
    return false
  }
  return true
}

function _processPacientesRow(row, sh, c1, c2, lc, _params, _diasAviso, editor, _M) {
  try {
    var rowData = sh.getRange(row, 1, 1, lc).getValues()[0]
    var dirty = {}

    if (rowData[COL.EDITOR - 1] !== editor) { rowData[COL.EDITOR - 1] = editor; dirty[COL.EDITOR] = true }

    var COLS_RUN = [COL.RUN, COL.RUN_CUIDADOR]
    for (var ri = 0; ri < COLS_RUN.length; ri++) {
      var rc = COLS_RUN[ri]
      if (c1 <= rc && c2 >= rc) {
        var rRaw = String(rowData[rc - 1] || '').trim()
        if (rRaw) {
          var rFmt = formatearRUT(rRaw)
          if (rFmt !== rRaw) { rowData[rc - 1] = rFmt; dirty[rc] = true }

          if (rFmt.indexOf('-') > 0 && rFmt.length >= 4) {
            var _notaRun = _notaRUN(rFmt)
            sh.getRange(row, rc).setNote(_notaRun)
            _pintarRUT(sh, row, rc, _notaRun)
          } else {
            sh.getRange(row, rc).setNote(null)
          }
        } else {
          sh.getRange(row, rc).setNote(null)
          _pintarRUT(sh, row, rc, null)
        }
      }
    }

    if (c1 <= COL.TELEFONO && c2 >= COL.TELEFONO) {
      var _telRaw = String(rowData[COL.TELEFONO - 1] || '').trim()
      if (_telRaw) {
        var _telFmt = formatChilePhone(_telRaw)
        if (_telFmt !== _telRaw) { rowData[COL.TELEFONO - 1] = _telFmt; dirty[COL.TELEFONO] = true }
      }
    }

    if (c1 === c2 && TEXT_UPPER.indexOf(c1) >= 0) {
      var txtRaw = String(rowData[c1 - 1] || '').trim()
      if (txtRaw) {
        var txtUp = txtRaw.toUpperCase()
        if (txtUp !== txtRaw) { rowData[c1 - 1] = txtUp; dirty[c1] = true }
      }
    }

    var _nacRaw = rowData[8]
    var _nacDate = _parseDate(_nacRaw)
    if (_nacDate) {
      var _edadCalc = _calcularEdad(_nacDate)
      var _edadAct = rowData[COL.EDAD_USUARIO - 1]
      if (String(_edadCalc) !== String(_edadAct)) {
        rowData[COL.EDAD_USUARIO - 1] = _edadCalc
        dirty[COL.EDAD_USUARIO] = true
      }
    }

    var _sexCols = [COL.SEXO, COL.SEXO_CUIDADOR]
    for (var _sxi = 0; _sxi < _sexCols.length; _sxi++) {
      var _sxc = _sexCols[_sxi]
      if (c1 <= _sxc && c2 >= _sxc) {
        var _sxv = rowData[_sxc - 1]
        if (_sxv != null && String(_sxv).trim() !== '') {
          var _sxn = _normalizarSexo(_sxv)
          if (_sxn && String(_sxv) !== _sxn) {
            rowData[_sxc - 1] = _sxn
            dirty[_sxc] = true
          }
        }
      }
    }

    var _vital = String(rowData[COL.VITAL - 1] || '').trim()

    if (_vital !== 'FALLECIDO') {

      for (var cc = c1; cc <= c2; cc++) {
        if (cc === COL.EDAD_USUARIO) {
          var edadU = parseInt(rowData[COL.EDAD_USUARIO - 1]) || 0
          var empaU = _asignarEMPA(edadU)
          if (rowData[COL.EMPA_USUARIO - 1] !== empaU) { rowData[COL.EMPA_USUARIO - 1] = empaU; dirty[COL.EMPA_USUARIO] = true }
        } else if (cc === COL.EDAD_CUIDADOR) {
          var edadC = parseInt(rowData[COL.EDAD_CUIDADOR - 1]) || 0
          var empaC = _asignarEMPA(edadC)
          if (rowData[COL.EMPA_CUIDADOR - 1] !== empaC) { rowData[COL.EMPA_CUIDADOR - 1] = empaC; dirty[COL.EMPA_CUIDADOR] = true }
        }
      }
    }

    var dirtyCols = Object.keys(dirty).map(Number).sort(function(a,b) { return a - b })
    if (dirtyCols.length) {
      var groups = _agruparContiguos(dirtyCols)
      for (var gi = 0; gi < groups.length; gi++) {
        var g = groups[gi]
        var sub = []
        for (var si = 0; si < g[1]; si++) sub.push(rowData[g[0] + si - 1])
        sh.getRange(row, g[0], 1, g[1]).setValues([sub])
      }
    }

    if (c1 <= COL.VITAL && c2 >= COL.VITAL) {
      try {

        var _vitalRaw = String(rowData[COL.VITAL - 1] || '').trim()
        var _vitalNorm = _normalizarVitalEstado(_vitalRaw)
        if (_vitalNorm !== _vitalRaw) {
          sh.getRange(row, COL.VITAL).setValue(_vitalNorm)
        }
        var vitalNow = String(_vitalNorm || '').trim().toUpperCase()
        var _emV = _ESTADO_COLORS[vitalNow]
        sh.getRange(row, COL.VITAL)
          .setBackground(_emV ? _emV[0] : null).setFontColor(_emV ? _emV[1] : '#000000')
        _aplicarVitalFila(row, sh, lc, vitalNow)
      } catch (eVital) {
        console.error('onEdit ESTADO: ' + eVital.message)
      }
    }

    if (c1 <= 2 && c2 >= 2) {
      try {
        var _sectorRaw = String(rowData[1] || '').trim().replace(/\s+/g, ' ')
        var _sectorNorm = _sectorRaw ? _sectorRaw.toUpperCase() : _sectorRaw
        if (_sectorNorm !== _sectorRaw) sh.getRange(row, 2).setValue(_sectorNorm)
        var _sm = _SECTOR_COLORS[_sectorNorm]
        sh.getRange(row, 2).setBackground(_sm ? _sm[0] : null).setFontColor(_sm ? _sm[1] : '#000000')
      } catch (eSec) {
        console.error('onEdit SECTOR: ' + eSec.message)
      }
    }

    var _soloCols = []
    var _tocaControl = false
    for (var _ec = c1; _ec <= c2; _ec++) {
      if (_FECHAS_BY_COL[_ec] !== undefined) _soloCols.push(_ec)
      for (var _cfi = 0; _cfi < _CONTROL_FECHAS.length; _cfi++) {
        if (_CONTROL_FECHAS[_cfi][1] === _ec) { _tocaControl = true; break }
      }
    }
    if (_soloCols.length || (c1 <= COL.VITAL && c2 >= COL.VITAL)) {
      try {
        _colorearFechasFila(row, sh, lc, _params, _diasAviso, rowData,
          (c1 <= COL.VITAL && c2 >= COL.VITAL) ? null : _soloCols, _M && _M.f)
      } catch (eCol) {
        console.error('onEdit colorearFechasFila: ' + eCol.message)
      }
    }
    if (_soloCols.length) {
      try { _pintarFechasInvalidas(sh, row, lc, row) } catch (ePI) {
        console.error('onEdit fechas inválidas: ' + ePI.message)
      }
    }
    if (_tocaControl || (c1 <= COL.VITAL && c2 >= COL.VITAL)) {
      try { _recalcularPrioridad(row, sh, lc, _params, _diasAviso, rowData, _M && _M.c) } catch (ePrio) {
        console.error('onEdit recalcularPrioridad: ' + ePrio.message)
      }
    }

    for (var _cc = c1; _cc <= c2; _cc++) {
      if (_COLS_TEXTO_LIBRE.indexOf(_cc) >= 0) _limpiarFormatoCelda(sh, row, _cc)
    }
  } catch (e) {
    console.error('onEdit Pacientes error: ' + e.message)
  }
}
