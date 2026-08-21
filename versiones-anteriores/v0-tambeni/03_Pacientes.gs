// ─── PACIENTES: CRUD + OPERACIONES ────────────────────────────────────────────

// ─── BUSCADOR ────────────────────────────────────────────────────────────────

function buscarEnPacientes() {
  var html =
'<html><head><base target="_top"><style>' +
'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:0;padding:0;color:#202124;font-size:14px;background:#fff}' +
'.bar{position:sticky;top:0;background:#fff;padding:12px 12px 8px;border-bottom:1px solid #e0e0e0;z-index:10}' +
'.bar input{width:100%;box-sizing:border-box;padding:10px 12px;font-size:14px;border:2px solid #dadce0;border-radius:8px;outline:none}' +
'.bar input:focus{border-color:#1a73e8}' +
'.bar .info{font-size:11px;color:#5f6368;margin-top:4px;text-align:right}' +
'#res{overflow-y:auto;padding:4px 0}' +
'.it{padding:8px 12px;border-bottom:1px solid #f0f0f0;cursor:pointer}' +
'.it:hover{background:#e8f0fe}' +
'.it .nom{font-size:14px;font-weight:500;color:#1a1a1a;line-height:1.3}' +
'.it .det{font-size:12px;color:#5f6368;margin-top:2px}' +
'.it .det span{margin-right:12px}' +
'.it .rut{font-family:Consolas,monospace;font-size:12px;color:#1a73e8}' +
'.it .tag{display:inline-block;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:500}' +
'.tag-v{background:#e8f5e9;color:#2e7d32}.tag-f{background:#fce4ec;color:#c62828}.tag-e{background:#fff3e0;color:#e65100}.tag-s{background:#f3e5f5;color:#6a1b9a}.tag-p{background:#fff8e1;color:#f57f17}' +
'.vacio{padding:40px 12px;text-align:center;color:#9aa0a6;font-size:13px}' +
'</style></head><body>' +
'<div class="bar"><input type="text" id="q" placeholder="Nombre, RUN, apellido..." autofocus>' +
'<div class="info" id="cta"></div></div>' +
'<div id="res"><div class="vacio">Escribe para buscar en todas las columnas</div></div>' +
'<script>' +
'var inp=document.getElementById("q");var res=document.getElementById("res");var cta=document.getElementById("cta");' +
'var _t;' +
'function _e(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}' +
'function _tag(e){if(e==="FALLECIDO")return"tag-f";if(e==="EGRESO"||e==="SUSPENDIDO")return"tag-e";if(e==="PENDIENTE")return"tag-p";return"tag-v"}' +
'function _addItem(o){' +
'var el=document.createElement("div");el.className="it";el.dataset.fila=o.f;' +
'var n=document.createElement("div");n.className="nom";n.textContent=_e(o.n);el.appendChild(n);' +
'var d=document.createElement("div");d.className="det";' +
'var rs=document.createElement("span");rs.className="rut";rs.textContent=_e(o.ru);d.appendChild(rs);' +
'var ts=document.createElement("span");ts.className="tag "+_tag(o.e);ts.textContent=_e(o.e);d.appendChild(ts);' +
'var ss=document.createElement("span");ss.textContent=_e(o.s);d.appendChild(ss);' +
'el.appendChild(d);res.appendChild(el)' +
'}' +
'res.addEventListener("click",function(e){' +
'var t=e.target;while(t&&!t.dataset.fila)t=t.parentNode;if(t)google.script.run._navegarAFila(Number(t.dataset.fila))' +
'});' +
'inp.addEventListener("input",function(){' +
'clearTimeout(_t);' +
'_t=setTimeout(function(){' +
'var v=inp.value.trim();' +
'if(!v){res.innerHTML="<div class=\\"vacio\\">Escribe para buscar</div>";cta.textContent="";return}' +
'cta.textContent="Buscando...";' +
'google.script.run.withSuccessHandler(function(r){' +
'cta.textContent=r.length+" resultado"+(r.length===1?"":"s");' +
'res.innerHTML="";' +
'if(r.length===0){res.innerHTML="<div class=\\"vacio\\">Sin resultados</div>";return}' +
'for(var i=0;i<r.length;i++){_addItem(r[i])}' +
'})._doBuscarPacientes(v);' +
'},300);' +
'});' +
'</script></body></html>'

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(html).setTitle('Buscar Pacientes').setWidth(340)
  )
}

function _doBuscarPacientes(q) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) return []
  var lr = sh.getLastRow()
  if (lr < 4) return []
  var term = String(q).trim().toLowerCase()
  if (!term) return []

  var data = sh.getRange(4, 1, lr - 3, 12).getValues()
  var sc = [0, 2, 3, 4, 7, 11]
  var results = []
  var maxResults = 50
  for (var r = 0; r < data.length && results.length < maxResults; r++) {
    for (var ci = 0; ci < sc.length; ci++) {
      var idx = sc[ci]
      if (idx >= data[r].length) continue
      var val = String(data[r][idx] || '').toLowerCase()
      if (val.indexOf(term) !== -1) {
        results.push({
          f: r + 4,
          n: String(data[r][2] || '') + ' ' + String(data[r][3] || '') + ' ' + String(data[r][4] || ''),
          ru: String(data[r][7] || ''),
          s: String(data[r][1] || ''),
          e: String(data[r][5] || ''),
        })
        break
      }
    }
  }
  return results
}

function _navegarAFila(fila) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) return
  sh.setActiveRange(sh.getRange(fila, 1))
  ss.toast('Fila ' + fila, '', 2)
}

function aplicarFiltroBusqueda(optTerm) {
  var _cache = CacheService.getScriptCache()
  var _last = _cache.get('_fts')
  if (_last && (Date.now() - Number(_last)) < 250) return
  _cache.put('_fts', String(Date.now()), 10)

  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) return
  var term = optTerm !== undefined ? String(optTerm).trim().toLowerCase() : String(sh.getRange('B2').getValue() || '').trim().toLowerCase()
  var lr = sh.getLastRow()
  if (lr < 4) return

  var f = sh.getFilter()
  if (f) {
    f.remove()
  } else {
    sh.showRows(4, lr - 3)
  }

  if (term === '') {
    sh.getRange(2, 7, 1, 2).merge().setValue('')
    return
  }

  var data = sh.getRange(4, 1, lr - 3, 8).getValues()
  var ocultas = 0
  var ranges = []
  for (var r = 0; r < data.length; r++) {
    var match = (String(data[r][2]).toLowerCase().indexOf(term) !== -1) ||
                (String(data[r][3]).toLowerCase().indexOf(term) !== -1) ||
                (String(data[r][4]).toLowerCase().indexOf(term) !== -1) ||
                (String(data[r][5]).toLowerCase().indexOf(term) !== -1) ||
                (String(data[r][7]).toLowerCase().indexOf(term) !== -1)
    if (!match) {
      if (ranges.length > 0 && ranges[ranges.length - 1][0] + ranges[ranges.length - 1][1] === 4 + r) {
        ranges[ranges.length - 1][1]++
      } else {
        ranges.push([4 + r, 1])
      }
      ocultas++
    }
  }
  for (var i = 0; i < ranges.length; i++) {
    sh.hideRows(ranges[i][0], ranges[i][1])
  }

  var visibles = data.length - ocultas
  sh.getRange(2, 7, 1, 2).merge().setValue(visibles + ' de ' + data.length + ' pacientes')
}

// ─── AGREGAR ─────────────────────────────────────────────────────────────────

function agregarPaciente() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ss.toast('Hoja Pacientes no encontrada', '', 3); return }

  var f = sh.getFilter()
  if (f) f.remove()
  var lc = sh.getLastColumn()

  // Find last row with a numeric ID in column A
  var idCol = sh.getRange(4, 1, sh.getLastRow() - 3, 1).getValues()
  var lr = 3
  var maxId = 0
  for (var _i = idCol.length - 1; _i >= 0; _i--) {
    var _id = Number(idCol[_i][0])
    if (_id > maxId) maxId = _id
    if (_id > 0 && _i + 4 > lr) { lr = _i + 4; break }
  }
  if (lr < 4) { lr = 3; maxId = 0 }

  var nr = lr + 1
  if (nr > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1)

  if (lr >= 4) {
    sh.getRange(lr, 1, 1, lc).copyTo(sh.getRange(nr, 1, 1, lc))
  }
  sh.getRange(nr, 1, 1, lc).setStrikethrough(false).setUnderline(false)
  sh.setRowHeight(nr, 24)

  var num = maxId + 1

  var blank = []
  for (var c = 0; c < lc; c++) {
    if (c === COL.ID - 1) { blank.push(num); continue }
    if (c === COL.VITAL - 1) { blank.push('VIGENTE'); continue }
    blank.push('')
  }
  sh.getRange(nr, 1, 1, lc).setValues([blank])

  _actualizarEstadosFila(nr)

  try { sh.getRange(3, 1, nr - 2, lc).createFilter() } catch(e) {}

  sh.getRange(nr, 1).activate()
  ss.toast('Paciente #' + num + ' agregado — completa los datos en la fila', '', 4)
}

// ─── RECALCULAR ESTADOS POR FILA ───────────────────────────────────────────

function _actualizarEstadosFila(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) return
  var params = leerParametros()
  var diasAviso = params['DIAS_AVISO'] || 0
  var lc = sh.getLastColumn()
  var rowData = sh.getRange(row, 1, 1, lc).getValues()[0]
  var dirty = {}
  var vital = String(rowData[COL.VITAL - 1] || '').trim()
  if (vital !== 'FALLECIDO') {
    for (var c = 0; c < CONTROL_COLS.length; c++) {
      var map = CONTROL_COLS[c]
      var fc = map[1], ec = map[2]
      if (!ec) continue
      if (fc > lc || ec > lc) continue
      var actual = String(rowData[ec - 1] || '').trim()
      var pName = map[0].toUpperCase()
      var meses = params[pName] || 12
      var _rawFecha = rowData[fc - 1]
      var nuevo = _vigenteReceta(map[0], calcStatus(_parseDate(_rawFecha), meses, actual, diasAviso, _rawFecha))
      if (nuevo !== actual) { rowData[ec - 1] = nuevo; dirty[ec] = true }
    }
    // EDAD automática (col 10) desde F. NACIMIENTO (col 9); si no hay fecha
    // de nacimiento se respeta la edad manual
    var _nacD = _parseDate(rowData[8])
    if (_nacD) {
      var _edadAuto = _calcularEdad(_nacD)
      if (String(rowData[COL.EDAD_USUARIO - 1]) !== String(_edadAuto)) {
        rowData[COL.EDAD_USUARIO - 1] = _edadAuto
        dirty[COL.EDAD_USUARIO] = true
      }
    }
    // SEXO (col 7) y SEXO C. (col 17): valores normalizados F/M/PENDIENTE
    for (var _sxi2 = 0; _sxi2 < 2; _sxi2++) {
      var _sxc2 = _sxi2 === 0 ? COL.SEXO : COL.SEXO_CUIDADOR
      var _sxv2 = rowData[_sxc2 - 1]
      if (_sxv2 != null && String(_sxv2).trim() !== '') {
        var _sxn2 = _normalizarSexo(_sxv2)
        if (_sxn2 && String(_sxv2) !== _sxn2) {
          rowData[_sxc2 - 1] = _sxn2
          dirty[_sxc2] = true
        }
      }
    }
    var edadU = parseInt(rowData[COL.EDAD_USUARIO - 1]) || 0
    var empaU = _asignarEMPA(edadU)
    if (rowData[COL.EMPA_USUARIO - 1] !== empaU) { rowData[COL.EMPA_USUARIO - 1] = empaU; dirty[COL.EMPA_USUARIO] = true }
    var edadC = parseInt(rowData[COL.EDAD_CUIDADOR - 1]) || 0
    var empaC = _asignarEMPA(edadC)
    if (rowData[COL.EMPA_CUIDADOR - 1] !== empaC) { rowData[COL.EMPA_CUIDADOR - 1] = empaC; dirty[COL.EMPA_CUIDADOR] = true }
  }
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
    _formatearFechaNA(row, sh, lc, rowData)
  }
  if (vital === 'FALLECIDO') {
    sh.getRange(row, 1, 1, lc).setFontColor('#aaaaaa').setStrikethrough(true)
    sh.getRange(row, 1).setFontWeight('bold')
    _setControlesNA(sh, row, lc)
    _formatearFechaNA(row, sh, lc)
  }
  try { colorearFilaCapacitacion(row, sh, params, rowData) } catch(e) {}
  try { colorearFilaCCV(row, sh, params, rowData) } catch(e) {}
  try { colorearFilaZARIT(row, sh, params, rowData) } catch(e) {}
  try { colorearFilaTS(row, sh, params, rowData) } catch(e) {}
  try { colorearFilaPS(row, sh, params, rowData) } catch(e) {}
  try { colorearFilaEMPAUsuario(row, sh, params, rowData) } catch(e) {}
  try {
    var _sv = String(rowData[1] || '').trim().toUpperCase()
    var _sm = _SECTOR_COLORS[_sv]
    sh.getRange(row, 2).setBackground(_sm ? _sm[0] : null).setFontColor(_sm ? _sm[1] : '#000000')
  } catch(e) {}
  try {
    var _ev = String(rowData[5] || '').trim().toUpperCase()
    var _em = _ESTADO_COLORS[_ev]
    sh.getRange(row, 6).setBackground(_em ? _em[0] : null).setFontColor(_em ? _em[1] : '#000000')
  } catch(e) {}
}

// ─── RECALCULAR TODO (batch-read, batch-write) ─────────────────────────────

function recalcularTodo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ss.toast('Hoja Pacientes no encontrada', '', 3); return }
  var lr = sh.getLastRow()
  if (lr < 4) { ss.toast('Sin datos', '', 2); return }
  ss.toast('Recalculando estados…', '', 1)
  var lc = sh.getLastColumn()

  var params = leerParametros()
  var diasAviso = params['DIAS_AVISO'] || 0

  var needed = {}
  for (var ci = 0; ci < CONTROL_COLS.length; ci++) {
    needed[CONTROL_COLS[ci][1]] = true
    if (CONTROL_COLS[ci][2]) needed[CONTROL_COLS[ci][2]] = true
  }
  needed[COL.EDAD_USUARIO] = true
  needed[COL.EDAD_CUIDADOR] = true
  needed[COL.SEXO] = true
  needed[COL.SEXO_CUIDADOR] = true
  needed[9] = true  // F. NACIMIENTO → recalcular EDAD automáticamente
  needed[COL.EMPA_CUIDADOR] = true
  needed[COL.EMPA_USUARIO] = true
  needed[COL.PRIORIDAD] = true
  needed[COL.VITAL] = true

  var cols = Object.keys(needed).map(Number).filter(function(x) { return x <= lc }).sort(function(a,b) { return a - b })
  var col1 = cols[0], colN = cols[cols.length - 1]
  var colIdx = {}
  for (var i = 0; i < cols.length; i++) colIdx[cols[i]] = cols[i] - col1

  var data = sh.getRange(4, col1, lr - 3, colN - col1 + 1).getValues()
  var rows = data.length
  var writes = {}

  for (var ci = 0; ci < CONTROL_COLS.length; ci++) {
    var map = CONTROL_COLS[ci]
    var fc = map[1], ec = map[2]
    if (fc > lc || fc < col1 || fc > colN) continue
    var pName = map[0].toUpperCase()
    var meses = params[pName] || 12
    for (var r = 0; r < rows; r++) {
      var raw = data[r][colIdx[fc]]
      var parsed = _parseDate(raw)
      if (typeof raw === 'string' && raw.trim() && parsed) data[r][colIdx[fc]] = parsed
      if (ec && ec <= lc) {
        if (String(data[r][colIdx[COL.VITAL]] || '').trim() === 'FALLECIDO') continue
        var actual = String(data[r][colIdx[ec]] || '').trim()
        var nuevo = _vigenteReceta(map[0], calcStatus(parsed, meses, actual, diasAviso, raw))
        if (nuevo !== actual) { data[r][colIdx[ec]] = nuevo; writes[ec] = true }
      }
    }
  }

  var statusECs = []
  for (var ci = 0; ci < CONTROL_COLS.length; ci++) {
    var ec = CONTROL_COLS[ci][2]
    if (ec && ec <= lc) statusECs.push(ec)
  }
  for (var r = 0; r < rows; r++) {
    if (String(data[r][colIdx[COL.VITAL]] || '').trim() === 'FALLECIDO') continue
    // EDAD automática (col 10) desde F. NACIMIENTO (col 9)
    if (colIdx[9] !== undefined) {
      var _nacD2 = _parseDate(data[r][colIdx[9]])
      if (_nacD2) {
        var _edadAuto2 = _calcularEdad(_nacD2)
        if (String(data[r][colIdx[COL.EDAD_USUARIO]]) !== String(_edadAuto2)) {
          data[r][colIdx[COL.EDAD_USUARIO]] = _edadAuto2
          writes[COL.EDAD_USUARIO] = true
        }
      }
    }
    // SEXO (col 7) y SEXO C. (col 17): normalizar a F/M/PENDIENTE
    for (var _sxi3 = 0; _sxi3 < 2; _sxi3++) {
      var _sc3 = _sxi3 === 0 ? COL.SEXO : COL.SEXO_CUIDADOR
      if (colIdx[_sc3] !== undefined) {
        var _sv3 = data[r][colIdx[_sc3]]
        if (_sv3 != null && String(_sv3).trim() !== '') {
          var _sn3 = _normalizarSexo(_sv3)
          if (_sn3 && String(_sv3) !== _sn3) {
            data[r][colIdx[_sc3]] = _sn3
            writes[_sc3] = true
          }
        }
      }
    }
    var ecC = _asignarEMPA(parseInt(data[r][colIdx[COL.EDAD_CUIDADOR]]) || 0)
    if (data[r][colIdx[COL.EMPA_CUIDADOR]] !== ecC) { data[r][colIdx[COL.EMPA_CUIDADOR]] = ecC; writes[COL.EMPA_CUIDADOR] = true }
    var ecU = _asignarEMPA(parseInt(data[r][colIdx[COL.EDAD_USUARIO]]) || 0)
    if (data[r][colIdx[COL.EMPA_USUARIO]] !== ecU) { data[r][colIdx[COL.EMPA_USUARIO]] = ecU; writes[COL.EMPA_USUARIO] = true }
    var hasV = false, hasP = false, hasD = false
    for (var si = 0; si < statusECs.length; si++) {
      var s = String(data[r][colIdx[statusECs[si]]] || '').trim()
      if (s === 'VENCIDO') hasV = true
      else if (s === 'POR VENCER') hasP = true
      if (s && s !== 'N/A') hasD = true
    }
    var p = hasV ? 'URGENTE' : hasP ? 'POR REVISAR' : hasD ? 'AL DIA' : 'N/A'
    if (data[r][colIdx[COL.PRIORIDAD]] !== p) { data[r][colIdx[COL.PRIORIDAD]] = p; writes[COL.PRIORIDAD] = true }
  }

  var writeCols = Object.keys(writes).map(Number)
  for (var w = 0; w < writeCols.length; w++) {
    var wc = writeCols[w]
    var out = []
    for (var r = 0; r < rows; r++) out.push([data[r][colIdx[wc]]])
    sh.getRange(4, wc, rows, 1).setValues(out)
  }

  _colorearEdadPorEMPA(sh, lr, lc)
  try { colorearEMPAUsuario() } catch(e) {}
  try { colorearCapacitaciones() } catch(e) {}
  try { colorearCCV() } catch(e) {}
  try { colorearZARIT() } catch(e) {}
  try { colorearTS() } catch(e) {}
  try { colorearPS() } catch(e) {}
  try { colorearSector(sh, lr) } catch(e) {}
  try { colorearEstado(sh, lr) } catch(e) {}
  ss.toast('Estados recalculados en ' + rows + ' pacientes', '', 4)
}

function _colorearEdadPorEMPA(sh, lr, lc) {
  if (lr < 4) return
  var empaCCol = COL.EMPA_CUIDADOR
  var empaUCol = COL.EMPA_USUARIO
  if (empaCCol > lc && empaUCol > lc) return
  var edadCCol = COL.EDAD_CUIDADOR
  var edadUCol = COL.EDAD_USUARIO
  var r = lr - 3
  var c1 = Math.min(empaCCol, empaUCol)
  var cN = Math.max(empaCCol, empaUCol)
  if (cN > lc) cN = lc
  if (c1 > lc) return
  var data = sh.getRange(4, c1, r, cN - c1 + 1).getValues()
  var bgU = [], bgC = []
  for (var i = 0; i < r; i++) {
    var valC = empaCCol <= lc ? String(data[i][empaCCol - c1] || '').trim() : ''
    var valU = empaUCol <= lc ? String(data[i][empaUCol - c1] || '').trim() : ''
    bgC.push([valC === 'EMPA' ? '#e3f2fd' : valC === 'EMPAM' ? '#f3e5f5' : valC === 'N/A' ? '#f5f5f5' : '#ffffff'])
    bgU.push([valU === 'EMPA' ? '#e3f2fd' : valU === 'EMPAM' ? '#f3e5f5' : valU === 'N/A' ? '#f5f5f5' : '#ffffff'])
  }
  if (empaCCol <= lc && edadCCol <= lc) sh.getRange(4, edadCCol, r, 1).setBackgrounds(bgC)
  if (empaUCol <= lc && edadUCol <= lc) sh.getRange(4, edadUCol, r, 1).setBackgrounds(bgU)
}

function limpiarPacientesCompleto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Limpiando datos de pacientes…', '', 1)
  try { limpiarEspaciosPacientes() } catch(e) { ss.toast('Error en espacios: ' + e.message, '', 3) }
  SpreadsheetApp.flush()
  try { ponerMayusculasPacientes() } catch(e) { ss.toast('Error en mayúsculas: ' + e.message, '', 3) }
  SpreadsheetApp.flush()
  try { formatearRUTPacientes() } catch(e) { ss.toast('Error en RUN: ' + e.message, '', 3) }
  SpreadsheetApp.flush()
  try { formatearTelefonos() } catch(e) { ss.toast('Error en teléfonos: ' + e.message, '', 3) }
  ss.toast('Datos de pacientes limpiados y formateados', '', 3)
}

// ─── FORMATEAR TELÉFONOS ──────────────────────────────────────────────────

function formatearTelefonos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Formateando teléfonos…', '', 1)
  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Formatear Teléfonos',
    'Formatear número(s) en col 12 (TELÉFONO)\nEj: 9 1234 5678 / 2 2123 4567',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return

  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ui.alert('No se encontró la hoja Pacientes.'); return }

  var lr = sh.getLastRow()
  if (lr < 4) return
  var rows = lr - 3

  var data = sh.getRange(4, 12, rows, 1).getValues()
  var col12 = []
  var total = 0, dirty12 = false

  for (var r = 0; r < rows; r++) {
    var orig12 = String(data[r][0] || '').trim()
    var f12 = formatChilePhone(orig12)
    col12.push([f12])
    if (f12 !== orig12) { total++; dirty12 = true }
  }

  if (dirty12) sh.getRange(4, 12, rows, 1).setValues(col12)

  ss.toast(total + ' teléfonos formateados', '', 3)
}

// ─── FORMATEAR RUN ─────────────────────────────────────────────────────────

function formatearRUTPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Formateando RUN…', '', 1)
  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Formatear RUN',
    'Eliminar puntos y ajustar guion en col 8 (RUN) y col 18 (RUN CUIDADOR)\nEj: 12.345.678-9 → 12345678-9',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return

  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ui.alert('No se encontró la hoja Pacientes.'); return }

  var lr = sh.getLastRow()
  if (lr < 4) return
  var rows = lr - 3
  var total = 0
  var runCols = [COL.RUN, COL.RUN_CUIDADOR]
  var invalidos = 0
  for (var ci = 0; ci < runCols.length; ci++) {
    var col = runCols[ci]
    var rng = sh.getRange(4, col, rows, 1)
    rng.setNumberFormat('@')
    var data = rng.getValues()
    var dirty = false
    for (var i = 0; i < rows; i++) {
      var raw = String(data[i][0] || '').trim()
      if (raw) {
        var fmt = formatearRUT(raw)
        if (fmt !== raw) { data[i][0] = fmt; dirty = true; total++ }
      }
    }
    if (dirty) rng.setValues(data)
  }
  // Second pass: set notes (clear for valid, warn for invalid)
  for (var ci = 0; ci < runCols.length; ci++) {
    var col = runCols[ci]
    var data2 = sh.getRange(4, col, rows, 1).getValues()
    for (var i = 0; i < rows; i++) {
      var val = String(data2[i][0] || '').trim()
      if (val && val.indexOf('-') > 0 && val.length >= 4 && !_validarDigitoRUT(val)) {
        sh.getRange(i + 4, col).setNote('\u26a0\ufe0f RUN inv\u00e1lido: el d\u00edgito verificador no coincide')
        invalidos++
      } else {
        sh.getRange(i + 4, col).setNote(null)
      }
    }
  }
  var msg = total + ' RUN formateados'
  if (invalidos > 0) msg += ' (' + invalidos + ' con dígito verificador incorrecto — revisa las notas en las celdas)'
  ss.toast(msg, '', 4)
}

// ─── PONER MAYÚSCULAS ─────────────────────────────────────────────────────

function ponerMayusculasPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Convirtiendo a mayúsculas…', '', 1)
  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Poner Mayúsculas',
    'Convertir a MAYÚSCULAS columnas de texto:\nSECTOR, NOMBRE, APELLIDO, DIRECCIÓN, NOMBRE CUIDADOR',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return

  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ui.alert('No se encontró la hoja Pacientes.'); return }

  var lr = sh.getLastRow()
  if (lr < 4) return
  var rows = lr - 3
  var total = 0

  for (var ci = 0; ci < TEXT_UPPER.length; ci++) {
    var col = TEXT_UPPER[ci]
    var data = sh.getRange(4, col, rows, 1).getValues()
    var dirty = false
    for (var i = 0; i < rows; i++) {
      var val = String(data[i][0] || '').trim()
      if (val) {
        var upper = val.toUpperCase()
        if (upper !== val) { data[i][0] = upper; dirty = true; total++ }
      }
    }
    if (dirty) sh.getRange(4, col, rows, 1).setValues(data)
  }

  ss.toast(total + ' celdas convertidas a mayúsculas', '', 3)
}

// ─── ORDENAR PACIENTES ─────────────────────────────────────────────────────

function ordenarPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ss.toast('Hoja Pacientes no encontrada', '', 3); return }
  ss.toast('Ordenando pacientes…', '', 1)
  var lr = sh.getLastRow()
  if (lr < 4) { ss.toast('Sin datos para ordenar', '', 3); return }
  var lc = sh.getLastColumn()

  var ui = SpreadsheetApp.getUi()
  var resp = ui.alert('Ordenar Pacientes',
    'Ordenar todas las filas alfabéticamente por APELLIDO (col C)?\nEl índice (col A) se mueve con cada paciente.',
    ui.ButtonSet.YES_NO)
  if (resp !== ui.Button.YES) return
  var col = COL.APELLIDO

  var filter = sh.getFilter()
  if (filter) filter.remove()
  sh.showRows(4, lr - 3)

  var rng = sh.getRange(4, 1, lr - 3, lc)
  rng.sort({column: col, ascending: true})

  var seq = []
  for (var i = 0; i < lr - 3; i++) seq.push([i + 1])
  sh.getRange(4, 1, lr - 3, 1).setValues(seq)

  sh.getRange(3, 1, lr - 2, lc).createFilter()

  ss.toast('Pacientes ordenados A-Z y resecuenciados', '', 3)
}

// ─── FICHA PACIENTE ─────────────────────────────────────────────────────────

function verFichaPaciente() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ss.toast('Hoja Pacientes no encontrada', '', 3); return }

  var row = sh.getActiveRange() ? sh.getActiveRange().getRow() : -1
  if (row < 4) {
    var raw = SpreadsheetApp.getUi().prompt(
      'Buscar paciente',
      'Selecciona una fila primero, o ingresa RUN (con/sin guión) o N°:',
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL)
    if (raw.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return
    var q = raw.getResponseText().trim()
    if (!q) return
    row = _buscarFilaPaciente(sh, q)
    if (row < 4) { SpreadsheetApp.getUi().alert('Paciente no encontrado'); return }
  }

  var lc = Math.min(sh.getLastColumn(), PAC_ANCHOS.length)
  var vals = sh.getRange(3, 1, 1, lc).getValues()[0]
  var data = sh.getRange(row, 1, 1, lc).getValues()[0]
  var autor = ''
  try { autor = Session.getActiveUser().getEmail() } catch(e) { autor = 'Usuario' }

  var html = _buildFichaSidebarHtml(vals, data, lc, autor, row)
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(html).setTitle('Ficha Resumen Usuario PADDS').setWidth(520)
  )
}

function _buildFichaSidebarHtml(headers, data, lc, autor, row) {
  var n = String(data[2] || '') + ' ' + String(data[3] || '') + ' ' + String(data[4] || '')
  var ru = String(data[7] || ''); var es = String(data[5] || ''); var se = String(data[1] || '')
  var nh = String(data[0] || ''); var edad = String(data[9] || ''); var tel = String(data[11] || '')
  if (!n.trim()) n = 'Paciente ' + nh

  var tagCls = 'tag-d'; var tagTxt = es
  if (es === 'VIGENTE') { tagCls = 'tag-v'; tagTxt = '\u2714 ' + es }
  else if (es === 'FALLECIDO') { tagCls = 'tag-f'; tagTxt = '\u2718 ' + es }
  else if (es === 'EGRESO') { tagCls = 'tag-e'; tagTxt = '\u2192 ' + es }
  else if (es === 'SUSPENDIDO') { tagCls = 'tag-s'; tagTxt = '\u26A0 ' + es }

  var p = []
  p.push('<html><head><base target="_top"><meta charset="UTF-8"><style>')
  p.push('*{box-sizing:border-box}')
  p.push('body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:0;color:#1a1a1a;font-size:13px;background:#f8f9fa}')
  p.push('.hdr{background:linear-gradient(135deg,#1a237e,#283593);color:#fff;padding:14px 16px}')
  p.push('.hdr .nom{font-size:17px;font-weight:700;line-height:1.3}')
  p.push('.hdr .sub{font-size:11px;color:#9fa8da;margin-top:6px}')
  p.push('.hdr .sub sp{display:inline-block;margin-right:18px}')
  p.push('.badge{display:inline-block;padding:1px 10px;border-radius:10px;font-size:11px;font-weight:600;color:#fff}')
  p.push('.tag-v{background:#43a047}.tag-f{background:#e53935}.tag-e{background:#fb8c00}.tag-s{background:#8e24aa}.tag-d{background:#546e7a}')
  p.push('.bar-info{display:flex;flex-wrap:wrap;gap:6px;padding:10px 16px;background:#fff;border-bottom:1px solid #e0e0e0;font-size:12px}')
  p.push('.bar-info .b{background:#f0f2f5;border-radius:6px;padding:3px 10px}')
  p.push('.bar-info .l{color:#777;font-size:10px}')
  p.push('table{width:100%;border-collapse:collapse;margin:6px 0;background:#fff}')
  p.push('.sh{font-size:12px;font-weight:700;color:#fff;padding:5px 10px}')
  p.push('td{padding:3px 8px;font-size:12px;border:1px solid #e8e8e8;width:25%}')
  p.push('td.l{font-weight:600;color:#555;background:#fafafa}')
  p.push('td.v{color:#1a1a1a}')
  p.push('.vacio{color:#bbb;font-style:italic}')
  p.push('.s-n{color:#999} .s-v{color:#c62828;font-weight:600} .s-d{color:#2e7d32;font-weight:600}')
  p.push('.s-p{color:#e65100;font-weight:600} .s-x{color:#6a1b9a;font-weight:600}')
  p.push('.ftr{font-size:10px;color:#999;text-align:center;padding:12px 16px;border-top:1px solid #e0e0e0;margin-top:8px;background:#fff}')
  p.push('.bar-btn{position:sticky;bottom:0;padding:10px 16px;background:#fff;border-top:2px solid #1a237e;text-align:center}')
  p.push('.bar-btn button{background:#1a237e;color:#fff;border:none;padding:8px 24px;border-radius:6px;font-size:13px;cursor:pointer}')
  p.push('.bar-btn button:hover{background:#283593}')
  p.push('</style></head><body>')

  p.push('<div class="hdr"><div class="nom">' + _esc(n) + '</div>')
  p.push('<div class="sub"><sp>RUN ' + _esc(ru) + '</sp><sp>N° ' + _esc(nh) + '</sp></div></div>')

  p.push('<div class="bar-info">')
  p.push('<div class="b"><div class="l">Estado</div><span class="badge ' + tagCls + '">' + tagTxt + '</span></div>')
  p.push('<div class="b"><div class="l">Sector</div>' + _esc(se) + '</div>')
  p.push('<div class="b"><div class="l">Edad</div>' + _esc(edad) + ' años</div>')
  p.push('<div class="b"><div class="l">Teléfono</div>' + _esc(tel) + '</div>')
  p.push('</div>')

  for (var si = 0; si < PAC_SECCIONES.length; si++) {
    var sec = PAC_SECCIONES[si]
    if (sec.ini > lc) break
    var fin = Math.min(sec.fin, lc)
    var cBg = sec.bg || '#1a237e'
    var lBg = si % 2 === 0 ? '#ffffff' : '#f8f9fa'

    var pairs = []
    var c = sec.ini
    while (c <= fin) {
      var aH = String(headers[c - 1] || '')
      if (aH.indexOf('EDITOR') >= 0 || aH.indexOf('PROFESIONAL') >= 0) { c++; continue }
      var aV = data[c - 1]; c++
      var bH = null, bV = null
      if (c <= fin) { bH = String(headers[c - 1] || '')
        if (bH.indexOf('EDITOR') >= 0 || bH.indexOf('PROFESIONAL') >= 0) { c++; continue }
        bV = data[c - 1]; c++ }
      pairs.push({ lh: aH, lv: aV, rh: bH, rv: bV })
    }
    if (!pairs.length) continue

    p.push('<table>')
    p.push('<tr><td class="sh" colspan="4" style="background:' + cBg + '">' + _esc(sec.nombre) + '</td></tr>')
    for (var pi = 0; pi < pairs.length; pi++) {
      var pr = pairs[pi]
      var rb = pi % 2 === 0 ? '#ffffff' : lBg
      p.push('<tr><td class="l" style="background:' + rb + '">' + _esc(pr.lh) + '</td><td class="v" style="background:' + rb + '">' + _fmtValSidebar(pr.lv) + '</td>')
      if (pr.rh !== null) {
        p.push('<td class="l" style="background:' + rb + '">' + _esc(pr.rh) + '</td><td class="v" style="background:' + rb + '">' + _fmtValSidebar(pr.rv) + '</td>')
      } else {
        p.push('<td class="v" colspan="2" style="background:' + rb + '"></td>')
      }
      p.push('</tr>')
    }
    p.push('</table>')
  }

  p.push('<div class="ftr">' + _fmtFechaActual() + ' · Fila ' + row + '</div>')
  p.push('<div class="bar-btn">')
  p.push('<button onclick="descargarPDF()" id="btnPDF">Descargar PDF</button>')
  p.push('</div>')
  p.push('<script>')
  p.push('function descargarPDF(){')
  p.push('var btn=document.getElementById("btnPDF");btn.textContent="Generando...";btn.disabled=true;')
  p.push('google.script.run.withSuccessHandler(function(r){')
  p.push('btn.textContent="Descargar PDF";btn.disabled=false;')
  p.push('if(r.url){window.open(r.url,"_blank")}else{btn.textContent="Error"}')
  p.push('})._generarFichaPdfSidebar(' + row + ')')
  p.push('}')
  p.push('</script>')
  p.push('</body></html>')
  return p.join('')
}

function _fmtValHtml(v) {
  if (v === undefined || v === null) return '<span class="vacio">—</span>'
  if (typeof v === 'object' && v instanceof Date && !isNaN(v.getTime())) {
    return _pad2(v.getDate()) + '/' + _pad2(v.getMonth() + 1) + '/' + v.getFullYear()
  }
  var s = String(v).trim()
  if (s === '') return '<span class="vacio">—</span>'
  var cls = ''
  if (s === 'VENCIDO' || s === 'FALLECIDO') cls = 's-v'
  else if (s === 'AL DIA' || s === 'VIGENTE' || s === 'REVISADO') cls = 's-d'
  else if (s === 'POR VENCER' || s === 'PENDIENTE' || s === 'URGENTE') cls = 's-p'
  else if (s === 'SUSPENDIDO' || s === 'EGRESO') cls = 's-x'
  else if (s === 'N/A') cls = 's-n'
  return '<span class="' + cls + '">' + _esc(s) + '</span>'
}
function _fmtValSidebar(v) { return _fmtValHtml(v) }
function _fmtValPdf(v) { return _fmtValHtml(v) }

function _generarFichaPdfSidebar(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) return { error: 'Hoja no encontrada' }
  var lc = Math.min(sh.getLastColumn(), PAC_ANCHOS.length)
  var vals = sh.getRange(3, 1, 1, lc).getValues()[0]
  var data = sh.getRange(row, 1, 1, lc).getValues()[0]
  var autor = ''
  try { autor = Session.getActiveUser().getEmail() } catch(e) { autor = 'Usuario' }
  var html = _buildFichaHtml(vals, data, lc, autor)
  var nom = String(data[2] || '') + ' ' + String(data[3] || '') + ' ' + String(data[4] || '')
  if (!nom.trim()) nom = 'Paciente ' + String(data[0] || '')
  return _saveFichaPdf(html, nom.trim())
}


function _buscarFilaPaciente(sh, raw) {
  var esRun = raw.indexOf('-') >= 0 || raw.indexOf('k') >= 0 || raw.indexOf('K') >= 0 || /^\d{7,9}$/.test(raw)
  var lr = sh.getLastRow()
  if (lr < 4) return -1
  // Only read ID(1), SECTOR(2), NOMBRE(3), APELLIDO(4), APELLIDO2(5), RUN(8) = cols 1-8
  var data = sh.getRange(4, 1, lr - 3, 8).getValues()
  if (esRun) {
    var q = raw.replace(/[^0-9kK]/g, '').toLowerCase()
    for (var r = 0; r < data.length; r++) {
      var runCel = String(data[r][7] || '').replace(/[^0-9kK]/g, '').toLowerCase()
      if (runCel === q) return r + 4
    }
  } else {
    var num = parseInt(raw)
    if (num > 0) {
      if (num >= 4 && num <= lr) return num
      for (var r = 0; r < data.length; r++) {
        if (Number(data[r][0]) === num) return r + 4
      }
    }
    var q = raw.toLowerCase()
    for (var r = 0; r < data.length; r++) {
      for (var ci = 2; ci <= 4; ci++) {
        if (String(data[r][ci] || '').toLowerCase().indexOf(q) >= 0) return r + 4
      }
    }
  }
  return -1
}

function _buildFichaHtml(headers, data, lc, autor) {
  var secColors = PAC_SECCIONES.map(function(s) { return s.bg })
  var lightColors = PAC_SECCIONES.map(function(s) { return _lightenHex(s.bg) })

  var n = String(data[2] || '') + ' ' + String(data[3] || '') + ' ' + String(data[4] || '')
  var ru = String(data[7] || '')
  var es = String(data[5] || '')
  var se = String(data[1] || '')
  var nh = String(data[0] || '')

  var tagCls = 'tag-def'; var tagTxt = es
  if (es === 'VIGENTE') { tagCls = 'tag-v'; tagTxt = '\u2714 ' + es }
  else if (es === 'FALLECIDO') { tagCls = 'tag-f'; tagTxt = '\u2718 ' + es }
  else if (es === 'EGRESO') { tagCls = 'tag-e'; tagTxt = '\u2192 ' + es }
  else if (es === 'SUSPENDIDO') { tagCls = 'tag-s'; tagTxt = '\u26A0 ' + es }

  var h = '<html><head><meta charset="UTF-8"><style>' +
    '@page{size:A4;margin:10mm 8mm}' +
    'body{font-family:Arial,Helvetica,sans-serif;font-size:9.5pt;color:#222;line-height:1.4;margin:0;padding:0}' +
    '.bar{height:4px;background:linear-gradient(90deg,#1a237e,#4A86E8,#2E7D5B,#B5651D)}' +
    '.hdr{background:linear-gradient(135deg,#1a237e,#283593);color:#fff;padding:14px 18px 10px}' +
    '.hdr .nom{font-size:20pt;font-weight:bold;letter-spacing:0.3px}' +
    '.hdr .sub{font-size:8pt;color:#9fa8da;margin-top:6px}' +
    '.hdr .sub sp{display:inline-block;margin-right:28px}' +
    '.info{background:#f5f7fa;padding:8px 18px;border-bottom:3px solid #1a237e;margin-bottom:10px}' +
    '.info sp{display:inline-block;margin-right:30px;font-size:9pt;color:#444}' +
    '.info .lb{color:#888;font-size:7.5pt}' +
    '.badge{display:inline-block;padding:2px 12px;border-radius:12px;font-size:8pt;font-weight:bold;color:#fff}' +
    '.tag-v{background:linear-gradient(135deg,#43a047,#2e7d32)}.tag-f{background:linear-gradient(135deg,#e53935,#c62828)}' +
    '.tag-e{background:linear-gradient(135deg,#fb8c00,#e65100)}.tag-s{background:linear-gradient(135deg,#8e24aa,#6a1b9a)}.tag-def{background:#546e7a}' +
    'table{width:100%;border-collapse:collapse;margin-bottom:10px}' +
    '.sth{font-size:10pt;font-weight:bold;color:#fff;text-align:left;padding:5px 10px;letter-spacing:0.3px}' +
    'td{padding:3px 8px;font-size:8.5pt;border:0.5px solid #ddd}' +
    'td.l{font-weight:bold;color:#444;width:22%;background:#f8f8f8;border-left:3px solid #ddd}' +
    'td.v{color:#222;width:28%}' +
    '.vacio{color:#bbb;font-size:7pt;font-style:italic}' +
    '.s-v{color:#c62828;font-weight:bold} .s-d{color:#2e7d32;font-weight:bold}' +
    '.s-p{color:#e65100;font-weight:bold} .s-x{color:#8e24aa;font-weight:bold} .s-n{color:#999}' +
    '.ftr{font-size:7pt;color:#999;text-align:center;padding-top:10px;border-top:1px solid #e0e0e0;margin-top:6px}' +
    '.wm{position:fixed;bottom:2mm;left:0;right:0;text-align:center;font-size:6.5pt;color:#ccc}' +
    '</style></head><body>'

  h += '<div class="bar"></div>' +
    '<div class="hdr"><div class="nom">' + _esc(n) + '</div>' +
    '<div class="sub"><sp>RUN ' + _esc(ru) + '</sp><sp>N\u00b0 ' + _esc(nh) + '</sp></div></div>' +
    '<div class="info"><sp><span class="lb">Estado</span><br><span class="badge ' + tagCls + '">' + tagTxt + '</span></sp>' +
    '<sp><span class="lb">Sector</span><br>' + _esc(se) + '</sp>' +
    '<sp><span class="lb">Edad</span><br>' + _esc(String(data[9] || '')) + ' a\u00f1os</sp>' +
    '<sp><span class="lb">Tel\u00e9fono</span><br>' + _esc(String(data[11] || '')) + '</sp></div>'

  for (var si = 0; si < PAC_SECCIONES.length; si++) {
    var sec = PAC_SECCIONES[si]
    if (sec.ini > lc) break
    var fin = Math.min(sec.fin, lc)
    var cBg = secColors[si] || '#1a237e'
    var lBg = lightColors[si] || '#f5f5f5'

    var pairs = []
    var c = sec.ini
    while (c <= fin) {
      var aH = String(headers[c - 1] || '')
      if (aH.indexOf('EDITOR') >= 0 || aH.indexOf('PROFESIONAL') >= 0) { c++; continue }
      var aV = data[c - 1]
      c++
      var bH = null, bV = null
      if (c <= fin) { bH = String(headers[c - 1] || '');
        if (bH.indexOf('EDITOR') >= 0 || bH.indexOf('PROFESIONAL') >= 0) { c++; continue }
        bV = data[c - 1]; c++ }
      pairs.push({ lh: aH, lv: aV, rh: bH, rv: bV })
    }

    h += '<table>'
    h += '<tr><td class="sth" colspan="4" style="background:' + cBg + ';padding:5px 10px">' +
      _esc(sec.nombre) + '</td></tr>'

    for (var pi = 0; pi < pairs.length; pi++) {
      var p = pairs[pi]
      var rb = pi % 2 === 0 ? '#ffffff' : lBg
      var bClr = pi % 2 === 0 ? cBg : lBg
      h += '<tr>' +
        '<td class="l" style="background:' + rb + ';border-left:3px solid ' + cBg + '">' + _esc(p.lh) + '</td>' +
        '<td class="v" style="background:' + rb + '">' + _fmtValPdf(p.lv) + '</td>'
      if (p.rh !== null) {
        h += '<td class="l" style="background:' + rb + '">' + _esc(p.rh) + '</td>' +
          '<td class="v" style="background:' + rb + '">' + _fmtValPdf(p.rv) + '</td>'
      } else {
        h += '<td class="v" colspan="2" style="background:' + rb + '"></td>'
      }
      h += '</tr>'
    }
    h += '</table>'
  }

  h += '<div class="ftr">Ficha generada el ' + _fmtFechaActual() + ' — Sistema PADDS</div>'
  h += '<div class="wm">P.A.V.C. \u2014 Documento confidencial \u2014 ' + _fmtFechaActual() + '</div>'
  h += '</body></html>'
  return h
}



function _fmtFechaActual() {
  var d = new Date()
  return _pad2(d.getDate()) + '/' + _pad2(d.getMonth() + 1) + '/' + d.getFullYear()
}

function _saveFichaPdf(html, nombre) {
  var san = nombre.replace(/[\/\\:*?"<>|]/g, '_').substring(0, 80)
  var fname = 'Ficha_Resumen_' + san + '_' + _fmtFechaActual().replace(/\//g, '-') + '.pdf'

  var blob = Utilities.newBlob(html, 'text/html', 'temp.html').getAs('application/pdf')
  blob.setName(fname)

  var folderName = 'Fichas Pacientes PADDS'
  var folders = DriveApp.getFoldersByName(folderName)
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName)

  var file = folder.createFile(blob)
  return {
    url: file.getUrl(),
    download: 'https://docs.google.com/uc?export=download&id=' + file.getId(),
    name: fname
  }
}



// ─── PONER TOOLTIPS ────────────────────────────────────────────────────────

function ponerTooltipsPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Agregando tooltips a encabezados…', '', 1)
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) return
  var lc = Math.min(sh.getLastColumn(), _COLUMNAS.length)
  for (var c = 1; c <= lc; c++) {
    var col = _COLUMNAS[c]
    if (!col) continue
    var parts = [col.name]
    if (col.desc) parts.push(col.desc)
    if (col.vals) parts.push('Valores: ' + col.vals)
    if (col.auto) parts.push(col.auto)
    sh.getRange(3, c).setNote(parts.join('\n'))
  }
}

// ─── LIMPIAR ESPACIOS ──────────────────────────────────────────────────────

function limpiarEspaciosPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Limpiando espacios…', '', 1)
  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Limpiar Espacios',
    'Eliminar tabs, saltos de línea, espacios múltiples y caracteres no imprimibles en TODAS las celdas de texto de Pacientes?',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return

  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ui.alert('Hoja Pacientes no encontrada.'); return }
  var lr = sh.getLastRow()
  var lc = sh.getLastColumn()
  if (lr < 4) return

  var data = sh.getRange(4, 1, lr - 3, lc).getValues()
  var total = 0
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < data[r].length; c++) {
      if (typeof data[r][c] === 'string') {
        var cleaned = data[r][c].replace(/[\t\n\r]+/g, ' ').replace(/[ \xA0]+/g, ' ').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '').trim()
        if (cleaned !== data[r][c]) { data[r][c] = cleaned; total++ }
      }
    }
  }
  if (total > 0) sh.getRange(4, 1, lr - 3, lc).setValues(data)
  ss.toast(total + ' celdas limpiadas', '', 3)
}

// ─── REINDEXAR ─────────────────────────────────────────────────────────────

function reindexarPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Renumerando IDs…', '', 1)
  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Renumerar ID',
    '¿Renumerar la columna N° (ID) en orden 1, 2, 3...?\nÚtil después de ordenar o eliminar filas.',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return

  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ui.alert('Hoja Pacientes no encontrada.'); return }
  var lr = sh.getLastRow()
  if (lr < 4) return
  var seq = []
  for (var i = 0; i < lr - 3; i++) seq.push([i + 1])
  sh.getRange(4, 1, lr - 3, 1).setValues(seq)
  ss.toast('IDs renumerados: 1 a ' + (lr - 3), '', 3)
}

// ─── ELIMINAR PACIENTE ─────────────────────────────────────────────────────

function eliminarPaciente() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ss.toast('Hoja Pacientes no encontrada', '', 3); return }
  var a = sh.getActiveRange()
  if (!a) return
  var row = a.getRow()
  if (row < 4) { ss.toast('Selecciona una fila de paciente (fila 4 en adelante)', '', 3); return }
  var nom = String(sh.getRange(row, 3).getValue() || '').trim()
  var ape = String(sh.getRange(row, 4).getValue() || '').trim()
  var label = (nom || ape) ? (nom + ' ' + ape).trim() : 'Fila ' + row
  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Eliminar paciente',
    '¿Eliminar la fila ' + row + '?\n\n' + label + '\n\nEsta acción NO se puede deshacer.',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return
  sh.deleteRow(row)
  ss.toast('Paciente eliminado: ' + label, '', 4)
}

// ─── PROBAR DIAS_AVISO ─────────────────────────────────────────────────────

function probarDiasAviso() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ss.toast('Hoja Pacientes no encontrada', '', 3); return }

  ss.toast('Verificando ventana de aviso…', '', 1)
  var params = leerParametros()
  var dias = params['DIAS_AVISO'] || 30
  var lr = sh.getLastRow()
  if (lr < 4) { ss.toast('Sin datos en Pacientes', '', 2); return }
  var lc = sh.getLastColumn()

  var cols = [2, 3, 4, 5, 6]
  for (var c = 0; c < CONTROL_COLS.length; c++) {
    if (CONTROL_COLS[c][1] <= lc) cols.push(CONTROL_COLS[c][1])
    if (CONTROL_COLS[c][2] && CONTROL_COLS[c][2] <= lc) cols.push(CONTROL_COLS[c][2])
  }
  cols.sort(function(a, b) { return a - b })
  var col1 = cols[0], colN = cols[cols.length - 1]
  var idx = {}
  for (var i = 0; i < cols.length; i++) idx[cols[i]] = cols[i] - col1

  var data = sh.getRange(4, col1, lr - 3, colN - col1 + 1).getValues()
  var hoy = new Date()
  var tz = ss.getSpreadsheetTimeZone()

  var dentroVentana = [], fueraVentana = []
  for (var r = 0; r < data.length; r++) {
    var row = data[r]
    var vital = String(row[idx[6]] || '').trim()
    if (vital === 'FALLECIDO' || vital === 'EGRESO' || vital === 'SUSPENDIDO') continue
    for (var ci = 0; ci < CONTROL_COLS.length; ci++) {
      var map = CONTROL_COLS[ci]
      var fc = map[1], ec = map[2]
      if (!ec || idx[ec] === undefined) continue
      var status = String(row[idx[ec]] || '').trim()
      if (status !== 'AL DIA' && status !== 'V. RECETA' && status !== 'POR VENCER') continue
      var paramKey = map[0].toUpperCase()
      var mesesParam = params[paramKey]
      if (!mesesParam) continue
      var fechaRaw = row[idx[fc]]
      if (!(fechaRaw instanceof Date) || isNaN(fechaRaw.getTime())) continue
      var fechaVence = new Date(fechaRaw)
      fechaVence.setMonth(fechaVence.getMonth() + Number(mesesParam))
      var msRest = fechaVence.getTime() - hoy.getTime()
      var diasRest = Math.ceil(msRest / (1000 * 60 * 60 * 24))
      if (diasRest <= 0) continue
      var nom = (String(row[idx[3]] || '') + ' ' + String(row[idx[4]] || '')).trim()
      var item = {
        nom: nom || ('Fila ' + (r + 4)),
        control: map[0],
        fecha: Utilities.formatDate(fechaRaw, tz, 'dd/MM/yyyy'),
        vence: Utilities.formatDate(fechaVence, tz, 'dd/MM/yyyy'),
        restan: diasRest,
        aviso: dias,
        seria: _vigenteReceta(map[0], diasRest <= dias ? 'POR VENCER' : 'AL DIA'),
        actual: status,
      }
      if (diasRest <= dias && status === 'AL DIA') dentroVentana.push(item)
      else if (diasRest <= dias + 30) fueraVentana.push(item)
    }
  }

  dentroVentana.sort(function(a, b) { return a.restan - b.restan })
  fueraVentana.sort(function(a, b) { return a.restan - b.restan })

  var totalD = dentroVentana.length, totalF = fueraVentana.length

  // Build sidebar HTML
  var html = '<html><head><base target="_top"><style>' +
    'body{font-family:Arial,sans-serif;margin:0;padding:12px;color:#222;font-size:13px}' +
    'h2{margin:0 0 2px;font-size:16px;color:#1a237e}' +
    '.sub{font-size:11px;color:#888;margin-bottom:12px}' +
    '.card{border:1px solid #e0e0e0;border-radius:6px;padding:10px;margin-bottom:10px}' +
    '.num{font-size:28px;font-weight:300}' +
    '.num.r{color:#c62828} .num.o{color:#e65100} .num.g{color:#2e7d32}' +
    '.lbl{font-size:11px;color:#777}' +
    '.entry{padding:6px 0;border-bottom:1px solid #f0f0f0;font-size:12px}' +
    '.entry:last-child{border:none}' +
    '.entry .ctrl{font-weight:600;color:#333}' +
    '.entry .info{color:#888;font-size:11px}' +
    '.vacio{padding:20px;text-align:center;color:#999}' +
    '.btn{padding:10px 0;text-align:center}' +
    '.btn button{background:#1a237e;color:#fff;border:none;border-radius:4px;padding:8px 20px;font-size:13px;cursor:pointer}' +
    '.btn button:hover{background:#283593}' +
    '</style></head><body>' +
    '<h2>Prueba DIAS_AVISO</h2>' +
    '<div class="sub">Config: ' + dias + ' días de aviso · ' + (totalD + totalF) + ' resultados</div>' +
    '<div class="card">' +
    '<div class="num r">' + totalD + '</div><div class="lbl">Controles que deberían ser POR VENCER (AL DÍA dentro de ventana)</div>' +
    '</div>' +
    '<div class="card">' +
    '<div class="num g">' + totalF + '</div><div class="lbl">Próximos a entrar en ventana (próximos 30 días)</div>' +
    '</div>'

  if (totalD > 0) {
    html += '<h3 style="font-size:13px;color:#c62828;margin:8px 0 4px">⚠️ Por revisar</h3>'
    var maxD = Math.min(30, totalD)
    for (var di = 0; di < totalD; di++) {
      var it = dentroVentana[di]
      html += '<div class="entry"><div class="ctrl">' + _esc(it.nom) + ' — ' + it.control + '</div>' +
        '<div class="info">Últ: ' + it.fecha + ' · Vence: ' + it.vence + ' · Restan ' + it.restan + ' días · Actual: ' + it.actual + '</div></div>'
    }
    if (totalD > maxD) html += '<div class="entry" style="color:#999">... y ' + (totalD - maxD) + ' más</div>'
  }

  if (totalF > 0) {
    html += '<h3 style="font-size:13px;color:#e65100;margin:12px 0 4px">📅 Próximos (30 días)</h3>'
    var maxF = Math.min(20, totalF)
    for (var fi = 0; fi < maxF; fi++) {
      var it2 = fueraVentana[fi]
      html += '<div class="entry"><div class="ctrl">' + _esc(it2.nom) + ' — ' + it2.control + '</div>' +
        '<div class="info">Últ: ' + it2.fecha + ' · Vence: ' + it2.vence + ' · Restan ' + it2.restan + ' días</div></div>'
    }
    if (totalF > maxF) html += '<div class="entry" style="color:#999">... y ' + (totalF - maxF) + ' más</div>'
  }

  if (totalD === 0 && totalF === 0) {
    html += '<div class="vacio">✅ Todo correcto — no hay cambios pendientes</div>'
  }

  html += '<div class="btn"><button onclick="google.script.run.recalcularTodo()">Recalcular estados ahora</button></div>' +
    '<div class="sub" style="margin-top:8px;text-align:center">Usa el botón si quieres aplicar los cambios</div>' +
    '</body></html>'

  var output = HtmlService.createHtmlOutput(html).setTitle('Prueba DIAS_AVISO (' + dias + ' días)').setWidth(420)
  SpreadsheetApp.getUi().showSidebar(output)

  var shortMsg = totalD > 0
    ? totalD + ' control(es) deberían ser POR VENCER. ' + totalF + ' próximo(s).'
    : '✅ Todo correcto — ' + totalF + ' próximo(s) a entrar en ventana.'
  ss.toast(shortMsg, '', 5)
}

