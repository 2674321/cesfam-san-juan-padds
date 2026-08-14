// ════════════════════════════════════════════════════════
// 03_Pacientes.gs │ CRUD + Ficha PDF + VIGENCIA + PRIORIDAD + limpieza (hoja 'Pacientes')

// ─────────────────────────────────────────────────────────
// ─── PACIENTES: CRUD + OPERACIONES + VIGENCIA ─────────────────────────────

// ─── VIGENCIA POR FECHA (columna ↔ clave de Parámetros) ─────────────────────

var _FECHAS_COLOR = []
;(function() {
  for (var _fiC = 0; _fiC < _CONTROL_FECHAS.length; _fiC++) {
    _FECHAS_COLOR.push([_CONTROL_FECHAS[_fiC][1], _CONTROL_FECHAS[_fiC][2]])
  }

  _FECHAS_COLOR.push([80, 'PAÑALES'])
  _FECHAS_COLOR.push([81, 'INMUNIZACION'], [82, 'INMUNIZACION'])
  _FECHAS_COLOR.push([83, 'INMUNIZACION'], [84, 'INMUNIZACION'])
  for (var _fcCap = COL.CAP_INI; _fcCap <= COL.CAP_FIN; _fcCap++) {
    _FECHAS_COLOR.push([_fcCap, 'CAPACITACIONES'])
  }
})()
var _FECHAS_BY_COL = {}
for (var _fbm = 0; _fbm < _FECHAS_COLOR.length; _fbm++) _FECHAS_BY_COL[_FECHAS_COLOR[_fbm][0]] = _FECHAS_COLOR[_fbm][1]

function _mesesControl(params, key) {
  var k = String(key || '').toUpperCase()
  if (params && params['_DESACTIVADO_' + k]) return null
  if (params && params[k] !== undefined && params[k] !== null && !isNaN(Number(params[k]))) {
    return Number(params[k])
  }
  var DEF = {
    'PROXIMA CURACION': 0,
    'RECETAS CONTROLADAS': 3,
    'CURACIONES': 1,
    'SONDA FOLEY': 3,
    'CAPACITACIONES': 6,
    'INMUNIZACION': 12,
    'PAÑALES': 2,
  }
  return DEF[k] !== undefined ? DEF[k] : 12
}

function _precalcularMeses(params) {
  var c = {}, f = {}
  for (var i = 0; i < _CONTROL_FECHAS.length; i++) {
    var kc = _CONTROL_FECHAS[i][2]
    if (c[kc] === undefined) c[kc] = _mesesControl(params, kc)
  }
  for (var j = 0; j < _FECHAS_COLOR.length; j++) {
    var kf = _FECHAS_COLOR[j][1]
    if (f[kf] === undefined) f[kf] = _mesesControl(params, kf)
  }
  return { c: c, f: f }
}

// lista de un vistazo (rojo urgente · naranjo por revisar · verde al día ·

var _PRIORIDAD_TIRA_COLORS = {
  'URGENTE':     { bg: '#FEE2E2', fg: '#B91C1C' },
  'POR REVISAR': { bg: '#FFEDD5', fg: '#C2410C' },
  'AL DIA':      { bg: '#DCFCE7', fg: '#15803D' },
  'N/A':         { bg: '#EEF1F5', fg: '#64748B' },
}

function _pintarTiraPrioridad(sh, row, p) {
  var _tm = _PRIORIDAD_TIRA_COLORS[p] || _PRIORIDAD_TIRA_COLORS['N/A']
  sh.getRange(row, COL.ID)
    .setBackground(_tm.bg).setFontColor(_tm.fg)
    .setFontWeight(p === 'URGENTE' ? 'bold' : 'normal')
}

// ─── BUSCADOR ────────────────────────────────────────────────────────────────

function buscarEnPacientes() {
  var html =
'<html><head><base target="_top"><style>' +
'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:0;padding:0;color:#202124;font-size:14px;background:#fff}' +
'.bar{position:sticky;top:0;background:#fff;padding:12px 12px 8px;border-bottom:1px solid #E2E8F0;z-index:10}' +
'.bar input{width:100%;box-sizing:border-box;padding:10px 12px;font-size:14px;border:2px solid #CBD5E1;border-radius:8px;outline:none}' +
'.bar input:focus{border-color:#0F766E}' +
'.bar .info{font-size:11px;color:#64748B;margin-top:4px;text-align:right}' +
'#res{overflow-y:auto;padding:4px 0}' +
'.it{padding:8px 12px;border-bottom:1px solid #F1F5F9;cursor:pointer}' +
'.it:hover{background:#CCFBF1}' +
'.it .nom{font-size:14px;font-weight:500;color:#1a1a1a;line-height:1.3}' +
'.it .det{font-size:12px;color:#64748B;margin-top:2px}' +
'.it .det span{margin-right:12px}' +
'.it .rut{font-family:Consolas,monospace;font-size:12px;color:#0F766E}' +
'.it .tag{display:inline-block;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:500}' +
'.tag-v{background:#DCFCE7;color:#15803D}.tag-f{background:#FCE7F3;color:#B91C1C}.tag-e{background:#FFEDD5;color:#C2410C}.tag-s{background:#F3E8FF;color:#7E22CE}.tag-p{background:#FEF3C7;color:#B45309}' +
'.vacio{padding:40px 12px;text-align:center;color:#94A3B8;font-size:13px}' +
'</style></head><body>' +
'<div class="bar"><input type="text" id="q" placeholder="Nombre, RUN, apellido..." autofocus>' +
'<div class="info" id="cta"></div></div>' +
'<div id="res"><div class="vacio">Escribe para buscar en todas las columnas</div></div>' +
'<script>' +
'var inp=document.getElementById("q");var res=document.getElementById("res");var cta=document.getElementById("cta");' +
'var _t;' +
'function _e(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}' +
'function _tag(e){if(e==="FALLECIDO")return"tag-f";if(e==="EGRESO"||e==="EGRESO POR ALTA"||e==="SUSPENDIDO")return"tag-e";if(e==="PENDIENTE")return"tag-p";return"tag-v"}' +
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
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) return []
  var lr = sh.getLastRow()
  if (lr < 4) return []
  var term = _norm(q)
  if (!term) return []

  var data = sh.getRange(4, 1, lr - 3, 12).getValues()
  var sc = [0, 2, 3, 4, 7, 11]
  var results = []
  var maxResults = 50
  for (var r = 0; r < data.length && results.length < maxResults; r++) {
    for (var ci = 0; ci < sc.length; ci++) {
      var idx = sc[ci]
      if (idx >= data[r].length) continue
      var val = _norm(data[r][idx])
      var valDig = (idx === 7 || idx === 11) ? val.replace(/[^a-z0-9]/g, '') : val
      if (val.indexOf(term) !== -1 || (valDig && valDig.indexOf(term) !== -1)) {
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
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) return
  var rng = sh.getRange(fila, 1, 1, sh.getLastColumn())
  sh.setActiveRange(rng)
  ss.toast('Paciente en la fila ' + fila, 'PADDS', 2)
}

function aplicarFiltroBusqueda(optTerm) {
  var _cache = CacheService.getScriptCache()
  var _last = _cache.get('_fts')
  if (_last && (Date.now() - Number(_last)) < 250) return
  _cache.put('_fts', String(Date.now()), 10)

  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) return
  if (optTerm !== undefined) sh.getRange('B2').setValue(optTerm)
  _aplicarFiltrosPac(sh)
}

// Filtro combinado: busca el término de B2 y además oculta los pacientes cuyo
// estado (columna F) no coincida con el dropdown de G2. TODOS/vacío = sin filtro.
function _aplicarFiltrosPac(sh) {
  var term = _norm(String(sh.getRange('B2').getValue() || ''))
  var est = String(sh.getRange('G2').getValue() || '').trim().toUpperCase()
  var usando = term !== '' || (est !== '' && est !== 'TODOS')
  var lr = sh.getLastRow()
  if (lr < 4) return

  var f = sh.getFilter()
  if (f) f.remove()

  if (!usando) {
    sh.showRows(4, lr - 3)
    sh.getRange('D2').setValue('')
    var _lcR = sh.getLastColumn()
    try { sh.getRange(3, 1, lr - 2, _lcR).createFilter() } catch (eF) {}
    return
  }

  var data = sh.getRange(4, 1, lr - 3, 12).getValues()
  var ocultas = 0
  var ranges = []
  for (var r = 0; r < data.length; r++) {
    var estadoRow = _norm(data[r][5]).toUpperCase()
    var rowMatch = est === '' || est === 'TODOS' || estadoRow === est
    if (rowMatch && term !== '') {
      var nom = _norm(data[r][2]) + ' ' + _norm(data[r][3]) + ' ' + _norm(data[r][4])
      var run = _norm(data[r][7]).replace(/[^a-z0-9]/g, '')
      var tel = _norm(data[r][11]).replace(/[^a-z0-9]/g, '')
      rowMatch = nom.indexOf(term) !== -1 ||
                  run.indexOf(term) !== -1 ||
                  tel.indexOf(term) !== -1 ||
                  _norm(data[r][5]).indexOf(term) !== -1 ||
                  _norm(data[r][1]).indexOf(term) !== -1 ||
                  String(data[r][0]).indexOf(term) !== -1
    }
    if (!rowMatch) {
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
  sh.getRange('D2').setValue(visibles + ' de ' + data.length + ' pacientes')
}

// ─── FILTRO DE SECCIONES (dropdown F2, una celda) ────────────────────────────

// queda visible para que el dropdown nunca desaparezca.

function aplicarFiltroSecciones(sh) {
  sh = sh || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PAC)
  if (!sh) return
  var sel = String(sh.getRange('F2').getValue() || '').trim()
  var lc = sh.getLastColumn()
  var todas = sel === '' || sel.toUpperCase() === 'TODAS'
  for (var i = 0; i < PAC_SECCIONES.length; i++) {
    var sec = PAC_SECCIONES[i]

    if (sec.nombre === 'IDENTIFICACIÓN') continue
    if (sec.ini > lc) continue
    var fin = Math.min(sec.fin, lc)
    try {
      if (todas || sec.nombre === sel) {
        sh.showColumns(sec.ini, fin - sec.ini + 1)
      } else {
        sh.hideColumns(sec.ini, fin - sec.ini + 1)
      }
    } catch (eH) {}
  }
}

// ─── AGREGAR ─────────────────────────────────────────────────────────────────

function agregarPaciente() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) {   ss.toast('No se encontró la hoja ' + HOJA_PAC + '. Revisa que exista.', 'Pacientes', 4); return }

  var f = sh.getFilter()
  if (f) f.remove()
  var lc = sh.getLastColumn()

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
  sh.setRowHeight(nr, 26)

  var num = maxId + 1

  var blank = []
  for (var c = 0; c < lc; c++) {
    if (c === COL.ID - 1) { blank.push(num); continue }
    if (c === COL.VITAL - 1) { blank.push('VIGENTE'); continue }
    if (c === COL.SECTOR - 1) { blank.push('PENDIENTE'); continue }
    blank.push('')
  }
  sh.getRange(nr, 1, 1, lc).setValues([blank])

  try {
    var _ix = nr - 4
  var _fzBg = _UI.frozenBg[_ix % 2]
  var _dtBg = _UI.zebraBg[_ix % 2]
    var _nrBgs = []
    for (var _nc = 0; _nc < lc; _nc++) _nrBgs.push(_nc < 5 ? _fzBg : _dtBg)
    sh.getRange(nr, 1, 1, lc).setBackgrounds([_nrBgs])
      .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, _UI.border, SpreadsheetApp.BorderStyle.SOLID)
  } catch(eN) {}

  try {
    _pintarOpcionesFila(sh, nr, lc)
    _pintarRUT(sh, nr, COL.RUN, '')
    _pintarRUT(sh, nr, COL.RUN_CUIDADOR, '')
  } catch(ePc) {}

  _actualizarEstadosFila(nr)

  try {
    var _fNew = sh.getFilter()
    if (_fNew) _fNew.remove()
    sh.getRange(3, 1, nr - 2, lc).createFilter()
  } catch(e) {}

  sh.getRange(nr, 1).activate()
  ss.toast('Paciente #' + num + ' agregado — completa los datos en la fila', 'Pacientes', 4)
}

// ─── VIGENCIA POR FILA (color por fecha + prioridad general) ────────────────

function _colorearFechasFila(row, sh, lc, _params, _diasAviso, rowData, soloCols, mesesF) {
  if (!rowData) rowData = sh.getRange(row, 1, 1, lc).getValues()[0]
  var fallecido = String(rowData[COL.VITAL - 1] || '').trim() === 'FALLECIDO'

  var cols = [], sts = []
  for (var fi = 0; fi < _FECHAS_COLOR.length; fi++) {
    var def = _FECHAS_COLOR[fi]
    var fc = def[0]
    if (fc > lc) continue
    if (soloCols && soloCols.indexOf(fc) < 0) continue
    cols.push(fc)
    var meses = mesesF ? mesesF[def[1]] : _mesesControl(_params, def[1])
    sts.push(fallecido ? 'N/A' : _estadoFecha(rowData[fc - 1], meses, _diasAviso))
  }
  if (!cols.length) return

  var groups = _agruparContiguos(cols)
  var off = 0
  for (var g = 0; g < groups.length; g++) {
    var grp = groups[g]
    var n = grp[1]
    var bgs = [], fgs = [], wts = [], nfs = []
    for (var k = 0; k < n; k++) {
      var st = sts[off + k]
      var c = _ESTADO_FECHA_COLORS[st] || _ESTADO_FECHA_COLORS['N/A']
      bgs.push(c[0])
      fgs.push(c[1])
      wts.push(st === 'VENCIDO' ? 'bold' : 'normal')
      nfs.push(st === 'N/A' ? '@' : 'dd/mm/yyyy')
    }
    off += n
    var rng = sh.getRange(row, grp[0], 1, n)
    rng.setBackgrounds([bgs]).setFontColors([fgs]).setFontWeights([wts]).setNumberFormats([nfs])
  }
}

// ─── ESTILO DE FILA SEGÚN ESTADO (col 6) ─────────────────────────────────────

function _aplicarVitalFila(row, sh, lc, vitalUp) {
  var st = _VITAL_ROW_COLORS[vitalUp] || {}
  sh.getRange(row, 1, 1, lc)
    .setFontColor(st.fg || '#000000')
    .setStrikethrough(!!st.strike)
    .setFontStyle(st.italic ? 'italic' : 'normal')
  if (vitalUp === 'FALLECIDO') sh.getRange(row, 1).setFontWeight('bold')
}

function _recalcularPrioridad(row, sh, lc, _params, _diasAviso, rowData, mesesC) {
  if (!rowData) rowData = sh.getRange(row, 1, 1, lc).getValues()[0]
  var p = 'N/A'
  if (String(rowData[COL.VITAL - 1] || '').trim() !== 'FALLECIDO') {
    var hasV = false, hasPV = false, hasD = false
    for (var fi = 0; fi < _CONTROL_FECHAS.length; fi++) {
      var fc = _CONTROL_FECHAS[fi][1]
      if (fc > lc) continue
      var meses = mesesC ? mesesC[_CONTROL_FECHAS[fi][2]] : _mesesControl(_params, _CONTROL_FECHAS[fi][2])
      var st = _estadoFecha(rowData[fc - 1], meses, _diasAviso)
      if (st === 'VENCIDO') hasV = true
      else if (st === 'POR VENCER') hasPV = true
      else if (st === 'AL DIA') hasD = true
    }
    p = hasV ? 'URGENTE' : hasPV ? 'POR REVISAR' : hasD ? 'AL DIA' : 'N/A'
  }
  if (COL.PRIORIDAD <= lc) {
    var cur = String(rowData[COL.PRIORIDAD - 1] || '').trim()
    if (cur !== p) sh.getRange(row, COL.PRIORIDAD).setValue(p)
    var pm = _PRIORIDAD_PMAP[p]
    if (pm) sh.getRange(row, COL.PRIORIDAD).setBackground(pm[0]).setFontColor(pm[1])
  }
  _pintarTiraPrioridad(sh, row, p)
}

// ─── RECALCULAR ESTADOS POR FILA (agregarPaciente / uso manual) ─────────────

function _actualizarEstadosFila(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) return
  var _params = leerParametros()
  var _diasAviso = _params['DIAS_AVISO'] || 0
  var _M = _precalcularMeses(_params)
  var lc = sh.getLastColumn()
  var rowData = sh.getRange(row, 1, 1, lc).getValues()[0]
  var dirty = {}
  var vital = String(rowData[COL.VITAL - 1] || '').trim()

  if (vital !== 'FALLECIDO') {

    var _nacD = _parseDate(rowData[COL.F_NACIMIENTO - 1])
    if (_nacD) {
      var _edadAuto = _calcularEdad(_nacD)
      if (String(rowData[COL.EDAD_USUARIO - 1]) !== String(_edadAuto)) {
        rowData[COL.EDAD_USUARIO - 1] = _edadAuto
        dirty[COL.EDAD_USUARIO] = true
      }
    }

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
    var empacC = _asignarEMPA(parseInt(rowData[COL.EDAD_CUIDADOR - 1]) || 0)
    if (String(rowData[COL.EMPA_CUIDADOR - 1]) !== empacC) { rowData[COL.EMPA_CUIDADOR - 1] = empacC; dirty[COL.EMPA_CUIDADOR] = true }
    var empacU = _asignarEMPA(parseInt(rowData[COL.EDAD_USUARIO - 1]) || 0)
    if (String(rowData[COL.EMPA_USUARIO - 1]) !== empacU) { rowData[COL.EMPA_USUARIO - 1] = empacU; dirty[COL.EMPA_USUARIO] = true }
  }

  var dirtyCols = Object.keys(dirty).map(Number).sort(function(a, b) { return a - b })
  if (dirtyCols.length) {
    var groups = _agruparContiguos(dirtyCols)
    for (var gi = 0; gi < groups.length; gi++) {
      var g = groups[gi]
      var sub = []
      for (var si = 0; si < g[1]; si++) sub.push(rowData[g[0] + si - 1])
      sh.getRange(row, g[0], 1, g[1]).setValues([sub])
    }
  }

  try { _colorearFechasFila(row, sh, lc, _params, _diasAviso, rowData, null, _M.f) } catch(e) {}
  try { _pintarFechasInvalidas(sh, row, lc, row) } catch(ePI) {}

  var _estadoRaw = String(rowData[COL.ESTADO - 1] || '').trim()
  var _estadoNorm = _normalizarVitalEstado(_estadoRaw)
  if (_estadoNorm !== _estadoRaw) {
    sh.getRange(row, COL.ESTADO).setValue(_estadoNorm)
  }
  var _estadoUp = String(_estadoNorm || '').trim().toUpperCase()
  var _emE = _ESTADO_COLORS[_estadoUp]
  if (_emE) sh.getRange(row, COL.ESTADO).setBackground(_emE[0]).setFontColor(_emE[1])

  var _sectorUp = String(rowData[COL.SECTOR - 1] || '').trim().toUpperCase()
  var _smS = _SECTOR_COLORS[_sectorUp]
  if (_smS) sh.getRange(row, COL.SECTOR).setBackground(_smS[0]).setFontColor(_smS[1])

  _aplicarVitalFila(row, sh, lc, _estadoUp)

  try { _recalcularPrioridad(row, sh, lc, _params, _diasAviso, rowData, _M.c) } catch(e) {}

  for (var _cl = 0; _cl < _COLS_TEXTO_LIBRE.length; _cl++) {
    if (_COLS_TEXTO_LIBRE[_cl] <= lc) _limpiarFormatoCelda(sh, row, _COLS_TEXTO_LIBRE[_cl])
  }
}

// ─── RECALCULAR TODO (batch-read, batch-write) ─────────────────────────────

function recalcularTodo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) {   ss.toast('No se encontró la hoja ' + HOJA_PAC + '. Revisa que exista.', 'Pacientes', 4); return }
  var lr = sh.getLastRow()
  if (lr < 4) { ss.toast('Sin datos de pacientes para procesar', 'Pacientes', 2); return }
  ss.toast('Recalculando vigencias y prioridades…', 'PADDS', 1)
  var lc = sh.getLastColumn()

  var _params = leerParametros()
  var _diasAviso = _params['DIAS_AVISO'] || 0
  var _M = _precalcularMeses(_params)

  var needed = {}
  for (var fi = 0; fi < _FECHAS_COLOR.length; fi++) {
    var fc1 = _FECHAS_COLOR[fi][0]
    if (fc1 <= lc) needed[fc1] = true
  }
  needed[COL.VITAL] = true
  needed[COL.F_NACIMIENTO] = true
  needed[COL.EDAD_USUARIO] = true
  needed[COL.EDAD_CUIDADOR] = true
  needed[COL.SEXO] = true
  needed[COL.SEXO_CUIDADOR] = true
  needed[COL.EMPA_CUIDADOR] = true
  needed[COL.EMPA_USUARIO] = true
  needed[COL.PRIORIDAD] = true
  needed[COL.SECTOR] = true
  needed[COL.ESTADO] = true

  var cols = Object.keys(needed).map(Number).filter(function(x) { return x <= lc }).sort(function(a, b) { return a - b })
  var col1 = cols[0], colN = cols[cols.length - 1]
  var colIdx = {}
  for (var i = 0; i < cols.length; i++) colIdx[cols[i]] = cols[i] - col1

  var data = sh.getRange(4, col1, lr - 3, colN - col1 + 1).getValues()
  var rows = data.length
  var writes = {}

  for (var r = 0; r < rows; r++) {
    var r0 = data[r]
    var vital = String(r0[colIdx[COL.VITAL]] || '').trim()

    var vitalN = _normalizarVitalEstado(vital)
    if (vitalN !== vital) { r0[colIdx[COL.VITAL]] = vitalN; writes[COL.VITAL] = true; vital = vitalN }

    if (colIdx[COL.F_NACIMIENTO] !== undefined && colIdx[COL.EDAD_USUARIO] !== undefined) {
      var _nacD = _parseDate(r0[colIdx[COL.F_NACIMIENTO]])
      if (_nacD) {
        var _edadAuto = _calcularEdad(_nacD)
        if (String(r0[colIdx[COL.EDAD_USUARIO]]) !== String(_edadAuto)) {
          r0[colIdx[COL.EDAD_USUARIO]] = _edadAuto
          writes[COL.EDAD_USUARIO] = true
        }
      }
    }

    for (var _sxi3 = 0; _sxi3 < 2; _sxi3++) {
      var _sc3 = _sxi3 === 0 ? COL.SEXO : COL.SEXO_CUIDADOR
      if (colIdx[_sc3] === undefined) continue
      var _sv3 = r0[colIdx[_sc3]]
      if (_sv3 != null && String(_sv3).trim() !== '') {
        var _sn3 = _normalizarSexo(_sv3)
        if (_sn3 && String(_sv3) !== _sn3) {
          r0[colIdx[_sc3]] = _sn3
          writes[_sc3] = true
        }
      }
    }

    var empcC = _asignarEMPA(parseInt(r0[colIdx[COL.EDAD_CUIDADOR]]) || 0)
    if (String(r0[colIdx[COL.EMPA_CUIDADOR]]) !== empcC) { r0[colIdx[COL.EMPA_CUIDADOR]] = empcC; writes[COL.EMPA_CUIDADOR] = true }
    var empcU = _asignarEMPA(parseInt(r0[colIdx[COL.EDAD_USUARIO]]) || 0)
    if (String(r0[colIdx[COL.EMPA_USUARIO]]) !== empcU) { r0[colIdx[COL.EMPA_USUARIO]] = empcU; writes[COL.EMPA_USUARIO] = true }

    var p = 'N/A'
    if (vital !== 'FALLECIDO') {
      var hasV = false, hasPV = false, hasD = false
      for (var fi2 = 0; fi2 < _CONTROL_FECHAS.length; fi2++) {
        var fc2 = _CONTROL_FECHAS[fi2][1]
        if (colIdx[fc2] === undefined) continue
        var st = _estadoFecha(r0[colIdx[fc2]], _M.c[_CONTROL_FECHAS[fi2][2]], _diasAviso)
        if (st === 'VENCIDO') hasV = true
        else if (st === 'POR VENCER') hasPV = true
        else if (st === 'AL DIA') hasD = true
      }
      p = hasV ? 'URGENTE' : hasPV ? 'POR REVISAR' : hasD ? 'AL DIA' : 'N/A'
    }
    if (String(r0[colIdx[COL.PRIORIDAD]]) !== p) {
      r0[colIdx[COL.PRIORIDAD]] = p
      writes[COL.PRIORIDAD] = true
    }
  }

  var writeCols = Object.keys(writes).map(Number).sort(function(a, b) { return a - b })
  var wGroups = _agruparContiguos(writeCols)
  for (var w = 0; w < wGroups.length; w++) {
    var wc = wGroups[w][0]
    var wn = wGroups[w][1]
    var out = []
    for (var r2 = 0; r2 < rows; r2++) out.push(data[r2].slice(colIdx[wc], colIdx[wc] + wn))
    sh.getRange(4, wc, rows, wn).setValues(out)
  }

  try { colorearSector(sh, lr) } catch(e) {}
  try { colorearEstado(sh, lr) } catch(e) {}
  try { colorearFechas(sh, lr) } catch(e) {}
  try { _colorearPrioridad(sh, lr) } catch(e) {}
  try { _colorearEdadPorEMPA(sh, lr, lc) } catch(e) {}
  try {
    var _vFal = _VITAL_ROW_COLORS['FALLECIDO']
    var falRows = []
    var _vi = colIdx[COL.VITAL]
    for (var rf = 0; rf < rows; rf++) {
      if (String(data[rf][_vi] || '').trim().toUpperCase() === 'FALLECIDO') falRows.push(rf)
    }

    for (var _fr = 0; _fr < falRows.length; _fr++) {
      var start = falRows[_fr], len = 1
      while (_fr + 1 < falRows.length && falRows[_fr + 1] === falRows[_fr] + 1) { len++; _fr++ }
      sh.getRange(4 + start, 1, len, lc).setFontColor(_vFal.fg).setStrikethrough(true)
      sh.getRange(4 + start, 1, len, 1).setFontWeight('bold')
    }
  } catch(e) {}

  ss.toast('Vigencias recalculadas en ' + rows + ' pacientes', 'Pacientes', 4)
}

// ─── DIAGNÓSTICO: FILAS "EXTRA" (basura) EN PACIENTES ──────────────────────
// Inspecciona las últimas filas de Pacientes y detecta filas sin ID (columna A)
// debajo del último paciente: son basura (casillas desmarcadas / N/A sueltos)
// que ocupan espacio. El limpiador las elimina por el criterio del ID.

function diagnosticarFilasExtra() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) { ss.toast('No se encontró Pacientes', 'Pacientes', 4); return }
  var lr = sh.getLastRow()
  var lc = sh.getLastColumn()
  var desde = Math.max(4, lr - 12)
  var filas = lr - desde + 1
  var data = sh.getRange(desde, 1, filas, lc).getValues()

  var ultimaId = 3
  for (var i = 0; i < filas; i++) {
    var vId = Number(data[i][0])
    if (!isNaN(vId) && vId > 0) ultimaId = desde + i
  }

  var lineas = []
  var sinId = 0
  for (var i2 = 0; i2 < filas; i2++) {
    var nF = 0, nNA = 0, nFalse = 0, otros = []
    for (var c = 0; c < lc; c++) {
      var v = data[i2][c]
      if (v == null || String(v).trim() === '') continue
      var s = String(v).trim()
      if (s === 'N/A') { nNA++; continue }
      if (s === 'FALSE' || s === 'false' || v === false) { nFalse++; continue }
      nF++
      if (otros.length < 2) otros.push(s)
    }
    var filaN = desde + i2
    var esBasura = filaN > ultimaId
    if (esBasura) sinId++
    lineas.push('Fila ' + filaN + ': datos=' + nF + ' N/A=' + nNA + ' casillas=' + nFalse +
      (esBasura ? ' → SIN ID (basura)' : (otros.length ? ' · id=[' + otros.join(', ') + ']' : '')))
  }

  var resumen = 'Última fila con ID: ' + ultimaId + ' (de ' + lr + ' filas)\n' +
    lineas.join('\n') + '\n\nFilas sin ID al final: ' + sinId +
    '\n\n🧹 Córrelo: Pacientes → Datos → Limpiar filas vacías'
  ss.toast(resumen, 'Diagnóstico', 10)
}

// ─── LIMPIAR FILAS VACÍAS AL FINAL (Pacientes) ─────────────────────────────
// Elimina las filas completamente vacías de la hoja Pacientes (desde la fila 4)
// para que no queden filas huérfanas acumuladas al final. Nunca borra filas
// con valores ni fórmulas. Se repite el barrido por si quedaron huecos.

function limpiarFilasVaciasPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) { ss.toast('No se encontró la hoja Pacientes', 'Pacientes', 4); return }
  var total = _compactarPacientes(sh) + _limpiarFilasVaciasLoop(sh, 4)
  ss.toast(total > 0
    ? 'Se eliminaron ' + total + ' fila(s) sin ID / basura del final de Pacientes'
    : 'Pacientes está limpio: no había filas extra al final', 'Pacientes', 4)
}

function _colorearPrioridad(sh, lr) {
  if (!sh) sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PAC)
  if (!sh) return
  if (!lr) lr = sh.getLastRow()
  if (lr < 4 || sh.getLastColumn() < COL.PRIORIDAD) return
  var rows = lr - 3
  var data = sh.getRange(4, COL.PRIORIDAD, rows, 1).getValues()
  var bgs = [], fgs = [], tBgs = [], tFgs = [], tWts = []
  var pmap = _PRIORIDAD_PMAP
  for (var r = 0; r < rows; r++) {
    var k = String(data[r][0] || '').trim().toUpperCase()
    var m = pmap[k] || null
    bgs.push([m ? m[0] : null])
    fgs.push([m ? m[1] : '#000000'])
    var t = _PRIORIDAD_TIRA_COLORS[k] || _PRIORIDAD_TIRA_COLORS['N/A']
    tBgs.push([t.bg])
    tFgs.push([t.fg])
    tWts.push([k === 'URGENTE' ? 'bold' : 'normal'])
  }
  sh.getRange(4, COL.PRIORIDAD, rows, 1).setBackgrounds(bgs).setFontColors(fgs)
  sh.getRange(4, COL.ID, rows, 1).setBackgrounds(tBgs).setFontColors(tFgs).setFontWeights(tWts)
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
          bgC.push([valC === 'EMPA' ? '#E0F2FE' : valC === 'EMPAM' ? '#F3E8FF' : valC === 'N/A' ? '#F1F5F9' : '#ffffff'])
          bgU.push([valU === 'EMPA' ? '#E0F2FE' : valU === 'EMPAM' ? '#F3E8FF' : valU === 'N/A' ? '#F1F5F9' : '#ffffff'])
  }
  if (empaCCol <= lc && edadCCol <= lc) sh.getRange(4, edadCCol, r, 1).setBackgrounds(bgC)
  if (empaUCol <= lc && edadUCol <= lc) sh.getRange(4, edadUCol, r, 1).setBackgrounds(bgU)
}

// ─── LIMPIAR Y FORMATEAR DATOS ──────────────────────────────────────────────

function limpiarPacientesCompleto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Corregir datos de pacientes',
    'Esto revisará TODAS las filas y corregirá:\n' +
    '· espacios, saltos de línea y caracteres invisibles\n' +
    '· mayúsculas en nombres y direcciones\n' +
    '· RUN (dígito verificador y formato)\n' +
    '· teléfonos (formato chileno)\n\n¿Continuar?',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return

  ss.toast('Corrigiendo datos de todos los pacientes…', 'PADDS', 1)
  try { limpiarEspaciosPacientes(true) } catch(e) { ss.toast('Error en espacios: ' + e.message, 'PADDS', 3) }
  try { ponerMayusculasPacientes(true) } catch(e) { ss.toast('Error en mayúsculas: ' + e.message, 'PADDS', 3) }
  try { formatearRUTPacientes(true) } catch(e) { ss.toast('Error en RUN: ' + e.message, 'PADDS', 3) }
  try { formatearTelefonos(true) } catch(e) { ss.toast('Error en teléfonos: ' + e.message, 'PADDS', 3) }
  ss.toast('Datos de pacientes corregidos y formateados', 'Pacientes', 4)
}

// ─── FORMATEAR TELÉFONOS ──────────────────────────────────────────────────

function formatearTelefonos(confirmado) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  if (!confirmado) {
    ss.toast('Formateando teléfonos…', 'PADDS', 1)
    var ui = SpreadsheetApp.getUi()
    var r = ui.alert('Formatear Teléfonos',
      'Formatear número(s) en col 12 (TELÉFONO)\nEj: 9 1234 5678 / 2 2123 4567',
      ui.ButtonSet.YES_NO)
    if (r !== ui.Button.YES) return
  }

  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) { ui.alert('No se encontró la hoja ' + HOJA_PAC + '.'); return }

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

  if (!confirmado) ss.toast(total + ' teléfonos formateados', 'Pacientes', 3)
}

// ─── FORMATEAR RUN ─────────────────────────────────────────────────────────

function formatearRUTPacientes(confirmado) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  if (!confirmado) {
    ss.toast('Formateando RUN…', 'PADDS', 1)
    var ui = SpreadsheetApp.getUi()
    var r = ui.alert('Formatear RUN',
      'Eliminar puntos y ajustar guion en col 8 (RUN) y col 18 (RUN CUIDADOR)\nEj: 12.345.678-9 → 12345678-9',
      ui.ButtonSet.YES_NO)
    if (r !== ui.Button.YES) return
  }

  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) { ui.alert('No se encontró la hoja ' + HOJA_PAC + '.'); return }

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

  for (var ci = 0; ci < runCols.length; ci++) {
    var col = runCols[ci]
    var data2 = sh.getRange(4, col, rows, 1).getValues()
    var notes = sh.getRange(4, col, rows, 1).getNotes()
    var bgs = [], fgs = []
    for (var i = 0; i < rows; i++) {
      var val = String(data2[i][0] || '').trim()
      var nota = (val && val.indexOf('-') > 0 && val.length >= 4) ? (_notaRUN(val) || '') : ''
      if (nota) invalidos++
          bgs.push([nota ? '#FEE2E2' : ((i + 4) % 2 === 0 ? _UI.zebraBg[0] : _UI.zebraBg[1])])
          fgs.push([nota ? '#B91C1C' : '#000000'])
      if (String(notes[i][0] || '') !== nota) sh.getRange(i + 4, col).setNote(nota || null)
    }
    sh.getRange(4, col, rows, 1).setBackgrounds(bgs).setFontColors(fgs)
  }
  var msg = total + ' RUN formateados'
  if (invalidos > 0) msg += ' (' + invalidos + ' con dígito verificador incorrecto — resaltados en rojo)'
  if (!confirmado) ss.toast(msg, 'Verificación de RUN', 4)
}

// ─── VERIFICAR INTEGRIDAD DE RUN ──────────────────────────────────────────

function verificarIntegridadRUN() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()
  var chk = _hojaPacientesValida(ss)
  if (!chk.ok) { ui.alert('Verificar RUN', chk.msg, ui.ButtonSet.OK); return }
  var sh = chk.sh
  var lr = sh.getLastRow()
  if (lr < 4) { ui.alert('Verificar RUN', 'No hay datos en ' + HOJA_PAC + '.', ui.ButtonSet.OK); return }
  var lc = Math.min(sh.getLastColumn(), 112)
  var data = sh.getRange(4, 1, lr - 3, lc).getValues()

  var invalidos = [], sinRun = [], dups = {}
  for (var i = 0; i < data.length; i++) {
    var rowN = i + 4
    var nombre = [data[i][2], data[i][3], data[i][4]].join(' ').replace(/\s+/g, ' ').trim()
    for (var ci = 0; ci < 2; ci++) {
      var col = ci === 0 ? COL.RUN : COL.RUN_CUIDADOR
      var raw = String(data[i][col - 1] || '').trim()
      if (!raw) {
        if (ci === 0) sinRun.push([rowN, nombre])
        continue
      }
      var fmt = formatearRUT(raw)
      var etiqueta = (ci === 0 ? 'Paciente' : 'Cuidador') + ' · fila ' + rowN + ' · ' + nombre
      if (!(fmt.indexOf('-') > 0 && fmt.length >= 4) || !_validarDigitoRUT(fmt)) {
        invalidos.push([etiqueta, fmt])
      }
      if (ci === 0) {
        if (!dups[fmt]) dups[fmt] = []
        dups[fmt].push('fila ' + rowN + ' · ' + nombre)
      }
    }
  }
  var dupList = []
  for (var k in dups) if (dups[k].length > 1) dupList.push([k, dups[k]])
  dupList.sort(function(a, b) { return a[1].length < b[1].length ? 1 : -1 })

  var _filas = function(lista, max) {
    var s = ''
    for (var x = 0; x < lista.length && x < max; x++) s += '<div class="row">' + _esc(lista[x][1]) + '</div>'
    if (lista.length > max) s += '<div class="row dim">… y ' + (lista.length - max) + ' más</div>'
    return s
  }
  var html = '<style>' + _uiCss() + '</style><body>' +
    '<div class="h">🔎 Verificación de RUN</div>' +
    '<div class="sub">Estado de los RUN registrados en Pacientes</div>' +
    '<div class="card"><div class="h err">⚠️ Dígito verificador inválido: ' + invalidos.length + '</div>' +
    '<div class="dim">' + _filas(invalidos, 25) + '</div></div>' +
    '<div class="card"><div class="h warn">🔁 RUN duplicados: ' + dupList.length + '</div>'
  for (var d2 = 0; d2 < dupList.length && d2 < 10; d2++) {
    html += '<div class="li">' + _esc(dupList[d2][0]) + '<br><span class="dim">' +
      _esc(dupList[d2][1].join(' · ')) + '</span></div>'
  }
  if (dupList.length > 10) html += '<div class="li dim">… y ' + (dupList.length - 10) + ' más</div>'
  html += '</div>' +
    '<div class="card"><div class="h warn">❓ Pacientes sin RUN: ' + sinRun.length + '</div>' +
    '<div class="dim">' + _filas(sinRun, 20) + '</div></div>' +
    '<div class="sub">Las celdas con dígito incorrecto ya llevan una nota ⚠️ en la celda. ' +
    'Para reformatear o limpiar notas usa 🛠️ Datos → "✨ Corregir datos (…)".</div></body>'
  ui.showSidebar(HtmlService.createHtmlOutput(html).setTitle('Verificación de RUN').setWidth(380))
}

// ─── PONER MAYÚSCULAS ─────────────────────────────────────────────────────

function ponerMayusculasPacientes(confirmado) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  if (!confirmado) {
    ss.toast('Convirtiendo a mayúsculas…', 'PADDS', 1)
    var ui = SpreadsheetApp.getUi()
    var r = ui.alert('Poner Mayúsculas',
      'Convertir a MAYÚSCULAS columnas de texto:\nSECTOR, NOMBRE, APELLIDO, DIRECCIÓN, NOMBRE CUIDADOR',
      ui.ButtonSet.YES_NO)
    if (r !== ui.Button.YES) return
  }

  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) { ui.alert('No se encontró la hoja ' + HOJA_PAC + '.'); return }

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

  if (!confirmado) ss.toast(total + ' celdas convertidas a mayúsculas', 'Pacientes', 3)
}

// ─── ORDENAR PACIENTES ─────────────────────────────────────────────────────

function ordenarPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) {   ss.toast('No se encontró la hoja ' + HOJA_PAC + '. Revisa que exista.', 'Pacientes', 4); return }
  ss.toast('Ordenando pacientes…', 'PADDS', 1)
  var lr = sh.getLastRow()
  if (lr < 4) { ss.toast('Sin datos para ordenar', 'Pacientes', 3); return }
  var lc = sh.getLastColumn()

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

  ss.toast('Pacientes ordenados A-Z y resecuenciados', 'Pacientes', 3)
}

// ─── FICHA PACIENTE ─────────────────────────────────────────────────────────

function verFichaPaciente() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) {   ss.toast('No se encontró la hoja ' + HOJA_PAC + '. Revisa que exista.', 'Pacientes', 4); return }

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

function _fichaMeta(es, data, lc, tagDef, prDef) {
  var paramsV = leerParametros() || {}
  var diasAvV = Number(paramsV['DIAS_AVISO']) || 15
  var fallecidoV = es === 'FALLECIDO'
  var vig = {}
  for (var _vgi = 0; _vgi < _CONTROL_FECHAS.length; _vgi++) {
    var _vd2 = _CONTROL_FECHAS[_vgi]
    if (_vd2[1] > lc) continue
    vig[_vd2[1]] = fallecidoV ? 'N/A' : _estadoFecha(data[_vd2[1] - 1], _mesesControl(paramsV, _vd2[2]), diasAvV)
  }
  var prV = String(data[COL.PRIORIDAD - 1] || '').trim()
  var prCls = prDef, prTxt = prV || '—'
  if (prV === 'URGENTE') { prCls = 'tag-f'; prTxt = prV }
  else if (prV === 'POR REVISAR') { prCls = 'tag-e'; prTxt = prV }
  else if (prV === 'AL DIA') { prCls = 'tag-v'; prTxt = prV }
  var tagCls = tagDef, tagTxt = es
  if (es === 'VIGENTE') { tagCls = 'tag-v'; tagTxt = '\u2714 ' + es }
  else if (es === 'FALLECIDO') { tagCls = 'tag-f'; tagTxt = '\u2718 ' + es }
  else if (es === 'EGRESO') { tagCls = 'tag-e'; tagTxt = '\u2192 ' + es }
  else if (es === 'EGRESO POR ALTA') { tagCls = 'tag-e'; tagTxt = '\u2192 ' + es }
  else if (es === 'SUSPENDIDO') { tagCls = 'tag-s'; tagTxt = '\u26A0 ' + es }
  return { tagCls: tagCls, tagTxt: tagTxt, prCls: prCls, prTxt: prTxt, vig: vig }
}

function _csvConteo(vig, lc, cols) {
  var venc = 0, pend = 0
  for (var i = 0; i < cols.length; i++) {
    if (cols[i] > lc) continue
    var st = vig[cols[i]]
    if (st === 'VENCIDO') venc++
    else if (st === 'PENDIENTE' || st === undefined) pend++
  }
  return { venc: venc, pend: pend }
}

function _buildSeccionesHtml(headers, data, lc, vig, o) {
  var out = []
  for (var si = 0; si < PAC_SECCIONES.length; si++) {
    var sec = PAC_SECCIONES[si]
    if (sec.ini > lc) break
    var fin = Math.min(sec.fin, lc)
    var cBg = sec.bg || '#1E293B'
    var hBg = o.hdrBg ? (typeof o.hdrBg === 'function' ? o.hdrBg(si, cBg) : o.hdrBg) : cBg
    var lBg = o.lBgFor(si)
    var pairs = []
    var c = sec.ini
    while (c <= fin) {
      if (o.skip && o.skip(c)) { c++; continue }
      var aH = String(headers[c - 1] || '')
      if (aH.indexOf('EDITOR') >= 0 || aH.indexOf('PROFESIONAL') >= 0) { c++; continue }
      var aV = data[c - 1]; var aC = c; c++
      var bH = null, bV = null, bC = null
      if (c <= fin) {
        if (o.skip && o.skip(c)) { c++; continue }
        bH = String(headers[c - 1] || '')
        if (bH.indexOf('EDITOR') >= 0 || bH.indexOf('PROFESIONAL') >= 0) { c++; continue }
        bV = data[c - 1]; bC = c; c++ }
      pairs.push({ lh: aH, lv: aV, lc: aC, rh: bH, rv: bV, rc: bC })
    }
    if (o.skipEmpty && !pairs.length) continue
    var lCol = o.lBdCol ? (typeof o.lBdCol === 'function' ? o.lBdCol(si, cBg) : o.lBdCol) : cBg
    var lSty = o.lBd ? ';border-left:2px solid ' + lCol : ''
    out.push('<table>')
    out.push('<tr><td class="' + o.hdrCls + '" colspan="4" style="background:' + hBg + (o.hdrPad || '') + '">' + _esc(sec.nombre) + '</td></tr>')
    for (var pi = 0; pi < pairs.length; pi++) {
      var pr = pairs[pi]
      var rb = pi % 2 === 0 ? '#ffffff' : lBg
      out.push('<tr><td class="l" style="background:' + rb + lSty + '">' + _esc(pr.lh) + '</td><td class="v" style="background:' + rb + '">' + _fmtValFicha(pr.lv, vig[pr.lc]) + '</td>')
      if (pr.rh !== null) {
        out.push('<td class="l" style="background:' + rb + '">' + _esc(pr.rh) + '</td><td class="v" style="background:' + rb + '">' + _fmtValFicha(pr.rv, vig[pr.rc]) + '</td>')
      } else {
        out.push('<td class="v" colspan="2" style="background:' + rb + '"></td>')
      }
      out.push('</tr>')
    }
    out.push('</table>')
  }
  return out.join('')
}

function _buildFichaSidebarHtml(headers, data, lc, autor, row) {
  var n = String(data[2] || '') + ' ' + String(data[3] || '') + ' ' + String(data[4] || '')
  var ru = String(data[7] || ''); var es = String(data[5] || ''); var se = String(data[1] || '')
  var nh = String(data[0] || ''); var edad = String(data[9] || ''); var tel = String(data[11] || '')
  if (!n.trim()) n = 'Paciente ' + nh

  var _m = _fichaMeta(es, data, lc, 'tag-d', 'tag-d')
  var tagCls = _m.tagCls, tagTxt = _m.tagTxt, prCls = _m.prCls, prTxt = _m.prTxt, vig = _m.vig

  var p = []
  p.push('<html><head><base target="_top"><meta charset="UTF-8"><style>')
  p.push('*{box-sizing:border-box}')
  p.push('body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:0;color:#1E293B;font-size:13px;background:#F1F5F9}')
  p.push('.hdr{background:#1E293B;color:#fff;padding:16px 18px 14px;border-bottom:3px solid #0F766E}')
  p.push('.hdr .nom{font-size:17px;font-weight:700;line-height:1.3}')
  p.push('.hdr .sub{font-size:11px;color:#CBD5E1;margin-top:6px;text-transform:uppercase;letter-spacing:0.8px}')
  p.push('.hdr .sub sp{display:inline-block;margin-right:18px}')
  p.push('.badge{display:inline-block;padding:2px 10px;border-radius:3px;font-size:11px;font-weight:600;color:#fff}')
  p.push('.tag-v{background:#15803D}.tag-f{background:#B91C1C}.tag-e{background:#C2410C}.tag-s{background:#7E22CE}.tag-d{background:#475569}')
  p.push('.bar-info{display:flex;flex-wrap:wrap;gap:6px;padding:10px 16px;background:#fff;border-bottom:1px solid #e3e8ef;font-size:12px}')
  p.push('.bar-info .b{background:#f4f6f8;border:1px solid #e3e8ef;border-radius:4px;padding:4px 10px}')
  p.push('.bar-info .l{color:#8a94a6;font-size:9.5px;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:2px}')
  p.push('table{width:100%;border-collapse:collapse;margin:8px 0;background:#fff}')
  p.push('.sh{font-size:11px;font-weight:700;color:#fff;padding:7px 12px;text-transform:uppercase;letter-spacing:0.6px;border-bottom:2px solid #0F766E}')
  p.push('td{padding:4px 10px;font-size:12px;border:1px solid #e6eaf0;width:25%}')
  p.push('td.l{font-weight:600;color:#5a6474;background:#f4f6f8;border-left:2px solid #dfe5ec}')
  p.push('td.v{color:#212529}')
  p.push('.vacio{color:#b0b8c4;font-style:italic}')
  p.push('.s-n{color:#8a94a6} .s-v{color:#B91C1C;font-weight:600} .s-d{color:#15803D;font-weight:600}')
  p.push('.s-p{color:#C2410C;font-weight:600} .s-x{color:#7E22CE;font-weight:600}')
  p.push('.v-d,.v-pv,.v-vd,.v-pd,.v-na{display:inline-block;padding:1px 5px;border-radius:4px;font-weight:600;font-size:11px}')
  p.push('.v-d{background:#DCFCE7;color:#15803D}.v-pv{background:#FFEDD5;color:#C2410C}.v-vd{background:#FEE2E2;color:#B91C1C}.v-pd{background:#FEF3C7;color:#B45309}.v-na{background:#F1F5F9;color:#64748B}')
  p.push('.leyV{font-size:9.5px;color:#6b7484;line-height:1.8;padding:0 2px}')
  p.push('.leyV i{display:inline-block;width:8px;height:8px;border-radius:2px;margin:0 3px 0 8px;vertical-align:middle}')
  p.push('.leyV i:first-child{margin-left:0}')
  p.push('.csvbox{margin:8px 16px 0;background:#fff;border:1px solid #dce3ec;border-radius:4px;overflow:hidden}')
  p.push('.csvhdr{background:#1E293B;color:#fff;font-size:11px;font-weight:700;padding:6px 12px;letter-spacing:0.4px}')
  p.push('.csvhdr.venc{background:#B91C1C}.csvhdr.pend{background:#C2410C}.csvcol{padding:8px 12px 6px}')
  p.push('.csvsec{font-size:9px;color:#1E293B;font-weight:700;letter-spacing:0.8px;text-transform:uppercase;margin:6px 0 4px}')
  p.push('.csvit{display:flex;justify-content:space-between;align-items:center;font-size:11px;padding:3px 0;border-bottom:1px solid #f0f2f5}')
  p.push('.csvit:last-child{border-bottom:none}.csvl{color:#64748B;font-size:10px}')
  p.push('.csvbox .v-d,.csvbox .v-pv,.csvbox .v-vd,.csvbox .v-pd,.csvbox .v-na{padding:0 6px}')
  p.push('.ftr{font-size:10px;color:#8a94a6;text-align:center;padding:12px 16px;border-top:1px solid #e3e8ef;margin-top:8px;background:#fff}')
  p.push('.bar-btn{position:sticky;bottom:0;padding:10px 16px;background:#fff;border-top:2px solid #1E293B;text-align:center}')
  p.push('.bar-btn button{background:#1E293B;color:#fff;border:none;padding:9px 26px;border-radius:4px;font-size:13px;font-weight:600;cursor:pointer}')
  p.push('.bar-btn button:hover{background:#0F172A}')
  p.push('</style></head><body>')

  p.push('<div class="hdr"><div class="nom">' + _esc(n) + '</div>')
  p.push('<div class="sub"><sp>RUN ' + _esc(ru) + '</sp><sp>N° ' + _esc(nh) + '</sp></div></div>')

  p.push('<div class="bar-info">')
  p.push('<div class="b"><div class="l">Estado</div><span class="badge ' + tagCls + '">' + tagTxt + '</span></div>')
  p.push('<div class="b"><div class="l">Sector</div>' + _esc(se) + '</div>')
  p.push('<div class="b"><div class="l">Edad</div>' + _esc(edad) + ' años</div>')
  p.push('<div class="b"><div class="l">Teléfono</div>' + _esc(tel) + '</div>')
  p.push('<div class="b"><div class="l">Prioridad</div><span class="badge ' + prCls + '">' + _esc(prTxt) + '</span></div>')
  p.push('</div>')

  var _csvUsrDef = [[51, 'EXÁMENES'], [52, 'C. MÉDICO'], [53, 'CCV MÉDICO'], [54, 'CSCV ENF.']]
  var _csvCudDef = [[23, 'EMPA/EMPAM'], [24, 'EXÁMENES'], [25, 'CCV VIGENTE']]
  var csvUsrHtml = ''
  for (var _cu = 0; _cu < _csvUsrDef.length; _cu++) {
    var _ccU = _csvUsrDef[_cu][0]
    if (_ccU > lc) continue
    var _stU = vig[_ccU]
    csvUsrHtml += '<div class="csvit"><span class="csvl">' + _csvUsrDef[_cu][1] + '</span><span>' +
      _fmtValFicha(data[_ccU - 1], _stU) + '</span></div>'
  }
  var csvCudHtml = ''
  for (var _cuu = 0; _cuu < _csvCudDef.length; _cuu++) {
    var _ccC = _csvCudDef[_cuu][0]
    if (_ccC > lc) continue
    var _stC = vig[_ccC]
    csvCudHtml += '<div class="csvit"><span class="csvl">' + _csvCudDef[_cuu][1] + '</span><span>' +
      _fmtValFicha(data[_ccC - 1], _stC) + '</span></div>'
  }
  var csvSt = _csvConteo(vig, lc, [51, 52, 53, 54, 23, 24, 25])
  var csvMensaje = csvSt.venc > 0 ? '\u26A0 ' + csvSt.venc + ' CONTROL(ES) VENCIDO(S)'
    : csvSt.pend > 0 ? '\u26A0 ' + csvSt.pend + ' PENDIENTE(S) SIN REGISTRO' : '\u2713 TODO AL D\u00cdA'
  var csvCls = csvSt.venc > 0 ? ' venc' : csvSt.pend > 0 ? ' pend' : ''
  if (csvUsrHtml || csvCudHtml) {
    p.push('<div class="csvbox">')
    p.push('<div class="csvhdr' + csvCls + '">SERVICIO CSV — ' + csvMensaje + '</div>')
    if (csvUsrHtml) p.push('<div class="csvcol"><div class="csvsec">USUARIO</div>' + csvUsrHtml + '</div>')
    if (csvCudHtml) p.push('<div class="csvcol"><div class="csvsec">CUIDADOR</div>' + csvCudHtml + '</div>')
    p.push('</div>')
  }

  p.push(_buildSeccionesHtml(headers, data, lc, vig, {
    hdrCls: 'sh', hdrBg: function() { return '#1E293B' },
    lBgFor: function(si) { return si % 2 === 0 ? '#ffffff' : '#f8fafc' }, skipEmpty: true }))

  p.push('<div class="ftr">')
  p.push('<div class="leyV"><i style="background:#DCFCE7"></i>AL DIA<i style="background:#FFEDD5"></i>POR VENCER<i style="background:#FEE2E2"></i>VENCIDO<i style="background:#FEF3C7"></i>PENDIENTE<i style="background:#F1F5F9"></i>N/A</div>')
  p.push('<div style="margin-top:8px">' + _fmtFechaActual() + ' · Fila ' + row + '</div>')
  p.push('<div style="font-size:9px;color:#bbb;margin-top:2px">Sistema PADDS v3.1 · Creado por Patricio Varela C. · Interno TENS · Contacto: patriciovarelacontreras@gmail.com</div>')
  p.push('</div>')
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
  if (typeof v === 'boolean') return v ? '<span class="s-d">✓ Sí</span>' : '<span class="vacio">—</span>'
  if (typeof v === 'object' && v instanceof Date && !isNaN(v.getTime())) {
    return _pad2(v.getDate()) + '/' + _pad2(v.getMonth() + 1) + '/' + v.getFullYear()
  }
  var s = String(v).trim()
  if (s === '') return '<span class="vacio">—</span>'
  var cls = ''
  if (s === 'VENCIDO' || s === 'FALLECIDO' || s === 'URGENTE') cls = 's-v'
  else if (s === 'AL DIA' || s === 'VIGENTE' || s === 'REVISADO') cls = 's-d'
  else if (s === 'POR VENCER' || s === 'PENDIENTE') cls = 's-p'
  else if (s === 'SUSPENDIDO' || s === 'EGRESO' || s === 'EGRESO POR ALTA') cls = 's-x'
  else if (s === 'N/A') cls = 's-n'
  else if (s === 'SI' || s === 'R' || s === 'ENTREGADO') cls = 's-d'
  else if (s === 'NO') cls = 's-v'
  else if (s === 'P') cls = 's-p'
  return '<span class="' + cls + '">' + _esc(s) + '</span>'
}
function _fmtValFicha(v, vig) {
  var cls = _vigCls(vig)
  if (!cls) return _fmtValHtml(v)
  var t
  if (v === undefined || v === null || String(v).trim() === '') {
    t = vig === 'PENDIENTE' ? 'PENDIENTE' : '—'
  } else if (typeof v === 'object' && v instanceof Date && !isNaN(v.getTime())) {
    t = _pad2(v.getDate()) + '/' + _pad2(v.getMonth() + 1) + '/' + v.getFullYear()
  } else {
    t = _esc(String(v).trim())
  }
  return '<span class="' + cls + '">' + t + '</span>'
}

function _vigCls(st) {
  if (st === 'AL DIA') return 'v-d'
  if (st === 'POR VENCER') return 'v-pv'
  if (st === 'VENCIDO') return 'v-vd'
  if (st === 'PENDIENTE') return 'v-pd'
  if (st === 'N/A') return 'v-na'
  return ''
}

function _generarFichaPdfSidebar(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
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
  var n = String(data[2] || '') + ' ' + String(data[3] || '') + ' ' + String(data[4] || '')
  var ru = String(data[7] || '')
  var es = String(data[5] || '')
  var se = String(data[1] || '')
  var nh = String(data[0] || '')

  var _m = _fichaMeta(es, data, lc, 'tag-def', 'tag-s')
  var tagCls = _m.tagCls, tagTxt = _m.tagTxt, prCls = _m.prCls, prTxt = _m.prTxt, vig = _m.vig

  var csvP = _csvConteo(vig, lc, [51, 52, 53, 54, 23, 24, 25])

  var h = '<html><head><meta charset="UTF-8"><style>' +
    '@page{size:A4;margin:9mm 7mm}' +
    'body{font-family:Arial,Helvetica,sans-serif;font-size:9.5pt;color:#212529;line-height:1.45;margin:0;padding:0}' +
    '.bar{height:3px;background:#1E293B;border-bottom:1.5px solid #0F766E}' +
    '.hdr{background:#1E293B;color:#fff;padding:16px 20px 12px;border-bottom:3px solid #0F766E}' +
    '.hdr .nom{font-size:19pt;font-weight:bold;letter-spacing:0.4px}' +
    '.hdr .sub{font-size:8pt;color:#CBD5E1;margin-top:6px;text-transform:uppercase;letter-spacing:1px}' +
    '.hdr .sub sp{display:inline-block;margin-right:28px}' +
    '.info{background:#fbfcfe;padding:9px 20px;border-top:1px solid #e3e8ef;border-bottom:1px solid #e3e8ef;margin-bottom:12px}' +
    '.info sp{display:inline-block;margin-right:26px;font-size:9pt;color:#3a4452}' +
    '.info .lb{font-size:6.8pt;color:#8a94a6;text-transform:uppercase;letter-spacing:0.6px}' +
    '.badge{display:inline-block;padding:2px 10px;border-radius:3px;font-size:8pt;font-weight:bold;color:#fff}' +
    '.tag-v{background:#15803D}.tag-f{background:#B91C1C}' +
    '.tag-e{background:#C2410C}.tag-s{background:#7E22CE}.tag-def{background:#475569}' +
    'table{width:100%;border-collapse:collapse;margin-bottom:12px}' +
    '.sth{font-size:9.5pt;font-weight:bold;color:#fff;text-align:left;padding:6px 10px;letter-spacing:0.6px;text-transform:uppercase;border-bottom:2px solid #0F766E}' +
    'td{padding:3.5px 8px;font-size:8.5pt;border:0.5px solid #E2E8F0}' +
    'td.l{font-weight:bold;color:#4a5464;width:22%;background:#f4f6f8;border-left:2px solid #dfe5ec}' +
    'td.v{color:#212529;width:28%}' +
    '.vacio{color:#b0b8c4;font-size:7.5pt;font-style:italic}' +
    '.s-v{color:#B91C1C;font-weight:bold} .s-d{color:#15803D;font-weight:bold}' +
    '.s-p{color:#C2410C;font-weight:bold} .s-x{color:#7E22CE;font-weight:bold} .s-n{color:#8a94a6}' +
    '.v-d,.v-pv,.v-vd,.v-pd,.v-na{display:inline-block;padding:0 5px;border-radius:2px;font-weight:700}' +
    '.v-d{background:#e3f1e8;color:#15803D}.v-pv{background:#fbeed9;color:#C2410C}.v-vd{background:#f9e3e1;color:#B91C1C}.v-pd{background:#f7f2d9;color:#B45309}.v-na{background:#eef1f5;color:#7a8494}' +
    '.ftr{font-size:7pt;color:#8a94a6;text-align:center;padding-top:10px;border-top:1px solid #E2E8F0;margin-top:8px}' +
    '.wm{position:fixed;bottom:2mm;left:0;right:0;text-align:center;font-size:6.5pt;color:#d5dae2}' +
    '</style></head><body>'

  h += '<div class="bar"></div>' +
    '<div class="hdr"><div class="nom">' + _esc(n) + '</div>' +
    '<div class="sub"><sp>RUN ' + _esc(ru) + '</sp><sp>N\u00b0 ' + _esc(nh) + '</sp></div></div>' +
    '<div class="info"><sp><span class="lb">Estado</span><br><span class="badge ' + tagCls + '">' + tagTxt + '</span></sp>' +
    '<sp><span class="lb">Sector</span><br>' + _esc(se) + '</sp>' +
    '<sp><span class="lb">Edad</span><br>' + _esc(String(data[9] || '')) + ' a\u00f1os</sp>' +
    '<sp><span class="lb">Tel\u00e9fono</span><br>' + _esc(String(data[11] || '')) + '</sp>' +
    '<sp><span class="lb">Prioridad</span><br><span class="badge ' + prCls + '">' + _esc(prTxt) + '</span></sp>' +
    '<sp><span class="lb">Servicio CSV</span><br><span class="badge ' + (csvP.venc > 0 ? 'tag-f' : csvP.pend > 0 ? 'tag-e' : 'tag-v') + '">' +
    (csvP.venc > 0 ? '\u26A0 ' + csvP.venc + ' vencido(s)' : csvP.pend > 0 ? '\u26A0 ' + csvP.pend + ' pendiente(s)' : '\u2713 todo al d\u00eda') + '</span></sp></div>'

  h += _buildSeccionesHtml(headers, data, lc, vig, {
    hdrCls: 'sth', hdrPad: ';padding:6px 10px', hdrBg: function() { return '#1E293B' },
    lBdCol: function(si, cBg) { return cBg }, lBd: true,
    lBgFor: function(si) { return si % 2 === 0 ? '#ffffff' : '#f6f8fa' },
    skip: function(c) { return c === COL.PRIORIDAD } })

  h += '<div class="ftr">Ficha generada el ' + _fmtFechaActual() + ' — Sistema PADDS v3.1 · Creado por Patricio Varela C. · Interno TENS · Contacto: patriciovarelacontreras@gmail.com</div>'
  h += '<div class="wm">P.A.V.C. \u2014 Documento confidencial \u2014 ' + _fmtFechaActual() + '</div>'
  h += '</body></html>'
  return h
}

function _fmtFechaActual() {
  return fmtFecha(new Date())
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
  ss.toast('Agregando tooltips a encabezados…', 'PADDS', 1)
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) return
  var lc = Math.min(sh.getLastColumn(), _COLUMNAS.length)
  for (var c = 1; c <= lc; c++) {
    var col = _COLUMNAS[c]
    if (!col) continue
    var sec = ''
    for (var s = 0; s < PAC_SECCIONES.length; s++) {
      if (c >= PAC_SECCIONES[s].ini && c <= PAC_SECCIONES[s].fin) { sec = PAC_SECCIONES[s].nombre; break }
    }
    var parts = [col.name]
    if (sec) parts.push('Sección: ' + sec)
    if (col.desc) parts.push(col.desc)
    if (col.vals) parts.push('Valores: ' + col.vals)
    if (col.auto) parts.push(col.auto)
    sh.getRange(3, c).setNote(parts.join('\n'))
  }
  ss.toast('Tooltips configurados en ' + lc + ' columnas', 'Pacientes', 3)
}

// ─── LIMPIAR ESPACIOS ──────────────────────────────────────────────────────

function limpiarEspaciosPacientes(confirmado) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  if (!confirmado) {
    ss.toast('Limpiando espacios…', 'PADDS', 1)
    var ui = SpreadsheetApp.getUi()
    var r = ui.alert('Limpiar Espacios',
      'Eliminar tabs, saltos de línea, espacios múltiples y caracteres no imprimibles en TODAS las celdas de texto de Pacientes?',
      ui.ButtonSet.YES_NO)
    if (r !== ui.Button.YES) return
  }

  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) { ui.alert('Hoja ' + HOJA_PAC + ' no encontrada.'); return }
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
  if (!confirmado) ss.toast(total + ' celdas limpiadas', 'Pacientes', 3)
}

// ─── REINDEXAR ─────────────────────────────────────────────────────────────

function reindexarPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Renumerando IDs…', 'PADDS', 1)

  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) { ss.toast('No se encontró la hoja ' + HOJA_PAC + '.', 'Pacientes', 4); return }
  var lr = sh.getLastRow()
  if (lr < 4) return
  var seq = []
  for (var i = 0; i < lr - 3; i++) seq.push([i + 1])
  sh.getRange(4, 1, lr - 3, 1).setValues(seq)
  ss.toast('IDs renumerados: 1 a ' + (lr - 3), 'Pacientes', 3)
}

// ─── ELIMINAR PACIENTE ─────────────────────────────────────────────────────

function eliminarPaciente() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) {   ss.toast('No se encontró la hoja ' + HOJA_PAC + '. Revisa que exista.', 'Pacientes', 4); return }
  var a = sh.getActiveRange()
  if (!a) return
  var row = a.getRow()
  if (row < 4) { ss.toast('Selecciona una fila de paciente (fila 4 en adelante)', 'Pacientes', 3); return }
  var nom = String(sh.getRange(row, 3).getValue() || '').trim()
  var ape = String(sh.getRange(row, 4).getValue() || '').trim()
  var label = (nom || ape) ? (nom + ' ' + ape).trim() : 'Fila ' + row
  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Eliminar paciente',
    '¿Eliminar la fila ' + row + '?\n\n' + label + '\n\nEsta acción NO se puede deshacer.',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return
  sh.deleteRow(row)
  ss.toast('Paciente eliminado: ' + label, 'Pacientes', 4)
}

function depurarDuplicados() {
  var ui = SpreadsheetApp.getUi()
  var r = ui.alert('Depurar duplicados y fechas falsas',
    'Revisará y corregirá automáticamente:\n' +
    '1) Líneas repetidas en CONTROLES MISCELÁNEOS, OTRAS PATOLOGÍAS y OBSERVACIONES (queda 1 copia).\n' +
    '2) Fechas vacías/falsas (31/12/1969, 31/12/1899) y fechas absurdas (año <1901 o >2100) en cualquier columna.\n' +
    '3) Entradas repetidas con la MISMA fecha y la MISMA prestación (p.ej. dos ZARIT del mismo día) → se deja solo la primera.\n\n' +
    'Ningún dato distinto se modifica. ¿Continuar?',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Depurando duplicados y fechas falsas…', 'PADDS', 1)
  try { depurarRegistrosDuplicados(true) } catch (e) { ss.toast('Registros: ' + e.message, 'PADDS', 4) }
  try { depurarEntradasDuplicadas(true) } catch (e) { ss.toast('Entradas: ' + e.message, 'PADDS', 4) }
}

function depurarRegistrosDuplicados(skipConfirm) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()
  var chk = _hojaPacientesValida(ss)
  if (!chk.ok) { ui.alert('Depurar registros', chk.msg, ui.ButtonSet.OK); return }
  var sh = chk.sh
  var lr = sh.getLastRow()
  if (lr < 4) return

  var resp = null
  if (!skipConfirm) {
    resp = ui.alert('Depurar registros',
      'Se corregirán automáticamente:\n' +
      '1) Líneas repetidas en CONTROLES MISCELÁNEOS, OTRAS PATOLOGÍAS y OBSERVACIONES (queda solo 1 copia).\n' +
      '2) Fechas vacías/falsas (31/12/1969, 31/12/1899) y fechas absurdas (año <1901 o >2100) en CUALQUIER columna, incluidas TELEFONO, OBSERVACIONES y EDITOR.\n\n' +
      'Solo se tocan copias idénticas o fechas vacías/absurdas: ningún dato distinto se modifica.\n¿Continuar?',
      ui.ButtonSet.YES_NO)
    if (resp !== ui.Button.YES) return
  }

  var lc = Math.min(sh.getLastColumn(), 112)
  var data = sh.getRange(4, 1, lr - 3, lc).getValues()
  var qDupMC = 0, qDupObs = 0, qEpoch = 0, qFuturo = 0, qMal = 0
  var set = []

  var colsFijar = [COL.CONTROLES_MISCELANEOS, COL.OTRAS_PATOLOGIAS, COL.OBSERVACIONES]
  for (var fc3 = 0; fc3 < colsFijar.length; fc3++) {
    if (colsFijar[fc3] <= lc) {
      try { sh.getRange(4, colsFijar[fc3], lr - 3, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP) } catch (eW2) {}
    }
  }
  if (lr >= 4) try { sh.setRowHeights(4, lr - 3, 26) } catch (eRH2) {}

  for (var r = 0; r < data.length; r++) {
    var rowN = r + 4

    var colsDup = []
    if (COL.CONTROLES_MISCELANEOS <= lc) colsDup.push(COL.CONTROLES_MISCELANEOS)
    if (COL.OTRAS_PATOLOGIAS <= lc) colsDup.push(COL.OTRAS_PATOLOGIAS)
    if (COL.OBSERVACIONES <= lc) colsDup.push(COL.OBSERVACIONES)
    for (var d = 0; d < colsDup.length; d++) {
      var col = colsDup[d]
      var raw0 = String(data[r][col - 1] || '')
      if (!raw0.trim()) continue
      var raw = _separarEntradas(raw0)
      var lines = raw.split('\n')
      var seen = {}, kept = [], dupe = 0
      for (var li = 0; li < lines.length; li++) {
        var ln = lines[li]
        if (!ln.trim()) continue
        if (seen[ln]) dupe++
        else { seen[ln] = true; kept.push(ln) }
      }
      if (dupe) {
        set.push([rowN, col, kept.join('\n')])
        if (col === COL.CONTROLES_MISCELANEOS) qDupMC += dupe
        else qDupObs += dupe
      }
    }

    for (var c2 = 0; c2 < lc; c2++) {
      var v2 = data[r][c2]
      if (v2 === '' || v2 === null || v2 === undefined) continue
      var mal = _fechaCorrupta(v2, (c2 + 1) === COL.F_NACIMIENTO)
      if (!mal) continue
      set.push([rowN, c2 + 1, ''])
      if (mal === 'epoch') qEpoch++
      else if (mal === 'mal') qMal++
      else qFuturo++
    }
  }

  if (!set.length) {
    ss.toast('✅ Sin hallazgos: no hay líneas duplicadas ni fechas vacías/absurdas', 'Pacientes', 4)
    return
  }

  var porCol = {}
  for (var s = 0; s < set.length; s++) {
    var k = set[s][1]
    if (!porCol[k]) porCol[k] = []
    porCol[k].push([set[s][0], set[s][2]])
  }
  for (var colKey in porCol) {
    var updates = porCol[colKey]
    var colData = sh.getRange(4, parseInt(colKey), lr - 3, 1).getValues()
    for (var u = 0; u < updates.length; u++) colData[updates[u][0] - 4][0] = updates[u][1]
    sh.getRange(4, parseInt(colKey), lr - 3, 1).setValues(colData)
    SpreadsheetApp.flush()
  }

  ui.alert('Depurado completado',
    'Líneas duplicadas eliminadas:\n' +
    '  · CONTROLES MISCELÁNEOS: ' + qDupMC + '\n' +
    '  · OTRAS PATOLOGÍAS: ' + qDupObs + '\n' +
    'Fechas vacías (31/12/1969, 31/12/1899) limpiadas: ' + qEpoch + '\n' +
    'Fechas absurdas (año <1901 o >2100) limpiadas: ' + qFuturo + '\n' +
    'Fechas inexistentes (ej. 40/13/2026) limpiadas: ' + qMal + '\n\n' +
    'Ninguna línea distinta fue modificada.',
    ui.ButtonSet.OK)
}

// ¿El valor es una fecha vacía/falsa o con año absurdo? (nunca es legítima)

function _fechaCorrupta(v, protegerNac) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    if (v.getTime() === 0) return 'epoch'
    if (!protegerNac) {

      var yA = v.getFullYear(), mA = v.getMonth(), dA = v.getDate()
      if ((yA === 1899 || yA === 1969) && mA === 11 && (dA === 30 || dA === 31)) return 'epoch'
    }
    var yDate = v.getFullYear()
    if (yDate < 1901 || yDate > 2100) return 'futuro'
    return false
  }
  if (typeof v !== 'string') return false
  var s = v.trim()
  if (!/^\d{1,2}\/\d{1,2}\/\d{2,}$/.test(s)) return false
  // Fecha con forma dd/mm/aaaa pero que no existe (ej. 40/13/2026, 30/02/2025).
  if (!_parseDate(s)) return 'mal'
  if (!protegerNac &&
      (s === '31/12/1969' || s === '30/12/1969' || s === '31/12/1899' || s === '30/12/1899')) return 'epoch'
  var mY = s.match(/(\d{2,})\s*$/)
  if (!mY) return false
  var yr = parseInt(mY[1], 10)
  if (yr < 1901 || yr > 2100) return 'futuro'
  return false
}

// ─── DEPURACIÓN POR ENTRADA (misma fecha y misma prestación) ──────────────

function _claveEntrada(ln) {
  var m = ln.match(/\[(\d{1,2}\/\d{1,2}\/\d{2,4})\]/)
  var cuerpo = m ? ln.slice(m.index + m[0].length) : ln
  var tipo = String(cuerpo).toLowerCase()
    .replace(/[^a-z0-9áéíóúñü]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 40)
  return (m ? m[1] : 'SIN_FECHA') + '|' + tipo
}

function _separarEntradas(raw) {
  return String(raw || '').split('\n').map(function(ln) {
    var t = ln.trim()
    return t.replace(/\[(\d{1,2}\/\d{1,2}\/\d{2,4})\]/g, function(m0, m1, off) {
      return off === 0 ? m0 : '\n' + m0
    })
  }).join('\n')
}

function depurarEntradasDuplicadas(skipConfirm) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  var ui = SpreadsheetApp.getUi()
  if (!sh) {
    ui.alert('Depurar entradas duplicadas', 'No se encontró la hoja ' + HOJA_PAC + '.', ui.ButtonSet.OK)
    return
  }
  var lr = sh.getLastRow()
  if (lr < 4) return
  var resp = null
  if (!skipConfirm) {
    resp = ui.alert('Depurar entradas duplicadas',
      'Revisa CONTROLES MISCELÁNEOS, OTRAS PATOLOGÍAS y OBSERVACIONES:\n' +
      '· Cada visita/prestación con fecha queda en su PROPIA LÍNEA.\n' +
      '· Si existe más de una entrada con la MISMA fecha y la MISMA prestación\n' +
      '  (p.ej. dos ZARIT del 07/08/2026), se deja solo la primera copia.\n\n' +
      'Las entradas con fecha o prestación distintas NO se tocan.\n¿Continuar?',
      ui.ButtonSet.YES_NO)
    if (resp !== ui.Button.YES) return
  }

  var lc = sh.getLastColumn()
  var colsD = [COL.CONTROLES_MISCELANEOS, COL.OTRAS_PATOLOGIAS, COL.OBSERVACIONES]
  var totalCel = 0, totalLin = 0, totalSep = 0
  for (var ci = 0; ci < colsD.length; ci++) {
    var col = colsD[ci]
    if (col > lc) continue
    var rng = sh.getRange(4, col, lr - 3, 1)
    var vals = rng.getValues()
    var celdas = 0, lineas = 0, sep = 0
    for (var r = 0; r < vals.length; r++) {
      var raw = String(vals[r][0] || '')
      if (!raw.trim()) continue
      var separado = _separarEntradas(raw)
      if (separado !== raw) sep++
      var lines = separado.split('\n')
      var seen = {}, kept = [], hubo = false
      for (var li = 0; li < lines.length; li++) {
        var ln = lines[li].trim()
        if (!ln) continue
        var key = _claveEntrada(ln)
        if (seen[key]) { lineas++; hubo = true; continue }
        seen[key] = true
        kept.push(ln)
      }
      if (hubo || sep > 0) {
        vals[r][0] = kept.join('\n')
        celdas++
      }
    }
    if (celdas) { rng.setValues(vals); SpreadsheetApp.flush() }
    try { rng.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP) } catch (eW) {}
    totalCel += celdas
    totalLin += lineas
    totalSep += sep
  }
  if (lr >= 4) try { sh.setRowHeights(4, lr - 3, 26) } catch (eRH) {}

  if (!totalLin && !totalSep) {
    ss.toast('✅ Sin hallazgos: no hay entradas repetidas ni texto continuo por separar', 'Pacientes', 4)
    return
  }
  ui.alert('Depurado completado',
    'Entradas duplicadas eliminadas: ' + totalLin + '\n' +
    'Celdas con texto separado en líneas: ' + totalSep + '\n' +
    'Celdas corregidas: ' + totalCel + '\n\n' +
    'Cada prestación quedó en su propia línea.',
    ui.ButtonSet.OK)
}
