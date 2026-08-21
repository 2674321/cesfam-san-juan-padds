// ─── EVENTOS: onEdit + procesamiento de filas ─────────────────────────────────

var _deleting = false
var _processingEdit = false

function onSelectionChange(e) {
  if (!e) return
  var sh = e.range.getSheet()
  var name = sh.getName()
  if (name !== 'Pacientes' && name !== 'Formulario Usuario / Profesional') return
  var row = e.range.getRow()
  if (name === 'Pacientes' && row < 4) return
  if (name === 'Formulario Usuario / Profesional' && row < 6) return

  // Debounce: skip if same row within 400ms
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

    // Formulario: color al cambiar estado en dropdown
    if (shName === 'Formulario Usuario / Profesional' && c1 === 3 && row >= 6) {
      var val = String(e.value || '').trim()
      if (_ESTADO_CSS[val]) {
        _colorEstado(sh, row, val, sh.getLastColumn())
      }
      return
    }

    // Parametros: invalidar caché
    if (shName === 'Parametros' && row >= 4) {
      _paramCache = null
      try { CacheService.getScriptCache().remove('_pac_params') } catch(e) {}
      return
    }

    // INGRESOS: formato automático de filas editadas.
    // La ventana de confirmación al poner OBSERVACION = INGRESA (verificación
    // de RUT, envío a Pacientes y eliminación de la fila) la maneja el
    // disparador instalable onEditIngresos — se activa desde el menú
    // 📥 Ingresos → "Activar confirmación de INGRESA" (ver 11_Ingresos.gs).
    // Nunca se borran hojas.
    if (shName === 'INGRESOS') {
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

    // Pacientes: búsqueda en vivo al editar B2
    if (shName === 'Pacientes' && row === 2 && c1 === 2) {
      var term = String(e.value || '').trim()
      if (term) _ss.toast('Buscando…', '', 1)
      aplicarFiltroBusqueda(term)
      return
    }

    // Pacientes: auto-actualizar
    if (shName === 'Pacientes' && row >= 4) {
      var _hasRelevant = false
      for (var cc = c1; cc <= c2; cc++) { if (_RELEVANT_COLS[cc]) { _hasRelevant = true; break } }
      if (!_hasRelevant) {
        // EDAD: refresco automático con CUALQUIER edición (también al cumplir
        // años), sin disparar el resto del procesamiento de la fila
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

        for (var r = row; r < row + numRows; r++) {
          if (r < 4) continue
          _processPacientesRow(r, sh, c1, c2, _lc, _params, _diasAviso, _editor)
        }
      } finally {
        lock.releaseLock()
      }
      return
    }

    // Agenda
    if (shName !== HOJA) return

    var colAg = e.range.getColumn()
    var valAgenda = String(e.value || '').trim()
    var oldVal = String(e.oldValue || '').trim()

    // ✕ / 📦 delete — solo si la celda tenía el emoji antes del borrado (evita falsos +
    // con OBSERVACIONES en col 6)
    if (!_deleting && valAgenda === '' && (oldVal === '✕' || oldVal === '📦')) {
      var delIdx = oldVal === '✕' ? 0 : 1
      for (var i = 0; i < 3; i++) {
        if (colAg === GI[i] + [5, 6][delIdx]) {
          _deleting = true
          try { e.range.setValue(oldVal) } catch(ee) {}

          var inicio = _findBloqueInicio(sh, row, GI[i])
          if (inicio < 0) { _deleting = false; _ss.toast('No se encontró cabecera de semana', '', 3); return }

          var ui = SpreadsheetApp.getUi()
          if (ui.alert('Limpiar semana', '¿Limpiar toda esta semana?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
            limpiarBloque(sh, inicio, GI[i])
            _ss.toast('Semana limpiada', '', 2)
          }
          _deleting = false
          return
        }
      }
    }

    // highlight fila al escribir en NOMBRE
    if (colAg % (PC + GC) === 3 && valAgenda.length > 0) {
      for (var i = 0; i < 3; i++) {
        if (colAg >= GI[i] && colAg < GI[i] + PC) {
          sh.getRange(row, GI[i], 1, 6).setBorder(true, true, true, true, true, true)
        }
      }
    }
  } catch(ee) {
    _ss.toast('Error en onEdit: ' + ee.message, '', 5)
  } finally {
    _processingEdit = false
  }
}

function _processPacientesRow(row, sh, c1, c2, lc, _params, _diasAviso, editor) {
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
          // Only validate if RUT looks complete (at least XX-X = 4 chars)
          if (rFmt.indexOf('-') > 0 && rFmt.length >= 4) {
            if (_validarDigitoRUT(rFmt)) {
              sh.getRange(row, rc).setNote(null)
            } else {
              sh.getRange(row, rc).setNote('\u26a0\ufe0f RUN inv\u00e1lido: el d\u00edgito verificador no coincide')
            }
          } else {
            sh.getRange(row, rc).setNote(null)
          }
        } else {
          sh.getRange(row, rc).setNote(null)
        }
      }
    }

    if (c1 === c2 && TEXT_UPPER.indexOf(c1) >= 0) {
      var txtRaw = String(rowData[c1 - 1] || '').trim()
      if (txtRaw) {
        var txtUp = txtRaw.toUpperCase()
        if (txtUp !== txtRaw) { rowData[c1 - 1] = txtUp; dirty[c1] = true }
      }
    }

    // EDAD automática (col 10) desde F. NACIMIENTO (col 9): se recalcula en
    // cada edición relevante (así también se corrige al cumplir años). Si no
    // hay fecha de nacimiento se respeta la edad manual
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

    // SEXO (col 7) y SEXO C. (col 17): valores normalizados F/M/PENDIENTE
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

    var _editedControlDate = false
    var _editedControlStatus = false

    if (_vital !== 'FALLECIDO') {
      for (var cc = c1; cc <= c2; cc++) {
        var map = COL_TO_CONTROL[cc]
        if (map) {
          _editedControlDate = true
          var actual = String(rowData[map[2] - 1] || '').trim()
          var pName = map[0].toUpperCase()
          var meses = _params[pName] || 12
          var nuevo = _vigenteReceta(map[0], calcStatus(_parseDate(rowData[cc - 1]), meses, actual, _diasAviso, rowData[cc - 1]))
          if (nuevo !== actual) { rowData[map[2] - 1] = nuevo; dirty[map[2]] = true }
        } else {
          var smap = COL_TO_STATUS[cc]
          if (smap) {
            _editedControlStatus = true
            var newStatus = String(rowData[cc - 1] || '').trim()
            if (newStatus === 'N/A') continue
            var fc = smap[1]
            var rawFecha = String(rowData[fc - 1] || '').trim()
            if (newStatus !== 'N/A' && rawFecha === 'N/A') {
              rowData[fc - 1] = ''
              dirty[fc] = true
            }
          } else if (cc === COL.EDAD_USUARIO) {
            var edad = parseInt(rowData[COL.EDAD_USUARIO - 1]) || 0
            var empaU = _asignarEMPA(edad)
            if (rowData[COL.EMPA_USUARIO - 1] !== empaU) { rowData[COL.EMPA_USUARIO - 1] = empaU; dirty[COL.EMPA_USUARIO] = true }
          } else if (cc === COL.EDAD_CUIDADOR) {
            var edadC = parseInt(rowData[COL.EDAD_CUIDADOR - 1]) || 0
            var empaC = _asignarEMPA(edadC)
            if (rowData[COL.EMPA_CUIDADOR - 1] !== empaC) { rowData[COL.EMPA_CUIDADOR - 1] = empaC; dirty[COL.EMPA_CUIDADOR] = true }
          }
        }
      }
    }

    // Write only changed columns (agrupadas contiguas)
    var dirtyCols = Object.keys(dirty).map(Number).sort(function(a,b) { return a - b })
    if (dirtyCols.length) {
      var groups = [[dirtyCols[0], 1]]
      for (var gi = 1; gi < dirtyCols.length; gi++) {
        var prev = groups[groups.length - 1]
        if (dirtyCols[gi] === prev[0] + prev[1]) { prev[1]++ }
        else { groups.push([dirtyCols[gi], 1]) }
      }
      for (var gi = 0; gi < groups.length; gi++) {
        var g = groups[gi]
        var sub = []
        for (var si = 0; si < g[1]; si++) sub.push(rowData[g[0] + si - 1])
        sh.getRange(row, g[0], 1, g[1]).setValues([sub])
      }
    }

    if (_editedControlDate || _editedControlStatus) _formatearFechaNA(row, sh, lc, rowData)

    if (c1 <= COL.VITAL && c2 >= COL.VITAL) {
      var vitalNow = String(rowData[COL.VITAL - 1] || '').trim().toUpperCase()
      var _eBg = {'VIGENTE':'#c8e6c9','FALLECIDO':'#f5f5f5','SUSPENDIDO':'#fff9c4','EGRESO':'#f5f5f5','ALTA':'#bbdefb','PENDIENTE':'#fff9c4'}
      var _eFg = {'VIGENTE':'#2e7d32','FALLECIDO':'#aaaaaa','SUSPENDIDO':'#888888','EGRESO':'#888888','ALTA':'#1565c0','PENDIENTE':'#f9a825'}
      sh.getRange(row, 6).setBackground(_eBg[vitalNow] || null).setFontColor(_eFg[vitalNow] || '#000000')
      sh.getRange(row, 1).setFontWeight(vitalNow === 'FALLECIDO' ? 'bold' : 'normal')
      if (vitalNow === 'FALLECIDO') {
        _setControlesNA(sh, row, lc)
        _formatearFechaNA(row, sh, lc)
      }
    }

    if (c1 <= COL.CAPACITACIONES_FIN && c2 >= COL.CAPACITACIONES_INI) {
      colorearFilaCapacitacion(row, sh, _params, rowData)
    }

    if (c1 === COL.CCV || c2 === COL.CCV) {
      colorearFilaCCV(row, sh, _params, rowData)
    }

    if (c1 === COL.ZARIT || c2 === COL.ZARIT) {
      colorearFilaZARIT(row, sh, _params, rowData)
    }

    if (c1 === 88 || c2 === 88) {
      colorearFilaTS(row, sh, _params, rowData)
    }

    if (c1 === 87 || c2 === 87) {
      colorearFilaPS(row, sh, _params, rowData)
    }

    if (c1 <= COL.EMPA_USUARIO_FECHA && c2 >= COL.EMPA_USUARIO_FECHA) {
      colorearFilaEMPAUsuario(row, sh, _params, rowData)
    }

    if (c1 === 2 || c2 === 2) {
      var _sectorVal = String(rowData[1] || '').trim().toUpperCase()
      var _sm = _SECTOR_COLORS[_sectorVal]
      sh.getRange(row, 2).setBackground(_sm ? _sm[0] : null).setFontColor(_sm ? _sm[1] : '#000000')
    }

    if (c1 === 6 || c2 === 6) {
      var _estVal = String(rowData[5] || '').trim().toUpperCase()
      var _em = _ESTADO_COLORS[_estVal]
      sh.getRange(row, 6).setBackground(_em ? _em[0] : null).setFontColor(_em ? _em[1] : '#000000')
    }

    for (var _cc = c1; _cc <= c2; _cc++) {
      if ([3,4,5,11,12,16,23,53,57,79,93,110].indexOf(_cc) >= 0) _limpiarFormatoCelda(sh, row, _cc)
    }
  } catch (e) {
    console.error('onEdit Pacientes error: ' + e.message)
  }
}

function _formatearFechaNA(row, sh, lc, data) {
  if (row < 4) return
  if (!data) data = sh.getRange(row, 1, 1, lc).getValues()[0]
  var naA1 = [], rsA1 = []
  var rowRef = row  // avoid re-get per iteration
  for (var _fi = 0; _fi < CONTROL_COLS.length; _fi++) {
    var _fc = CONTROL_COLS[_fi][1]
    if (_fc > lc) continue
    var _v = String(data[_fc - 1] || '').trim()
    if (_v === 'N/A') {
      naA1.push(colToLetter(_fc) + rowRef)
    } else {
      rsA1.push(colToLetter(_fc) + rowRef)
    }
  }
  if (naA1.length) {
    sh.getRangeList(naA1).setBackground('#f5f5f5').setFontColor('#999999').setFontWeight('normal').setNumberFormat('@')
  }
  if (rsA1.length) {
    sh.getRangeList(rsA1).setFontColor('#000000').setNumberFormat('dd/mm/yyyy')
  }
}

function _setControlesNA(sh, row, lc) {
  for (var _ni = 0; _ni < CONTROL_COLS.length; _ni++) {
    var sc = CONTROL_COLS[_ni][2]
    if (sc && sc <= lc) {
      sh.getRange(row, sc).setValue('N/A').setNumberFormat('@')
    }
  }
}

