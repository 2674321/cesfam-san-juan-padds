// ════════════════════════════════════════════════════════
// ÍNDICE ▏13_V4.gs │ PADDS V4 — capa de producto
//
// Interfaces de la versión 4: Centro de control, Buscador global,
// Ficha del paciente V4, Alertas/Tareas, Auditoría de datos y
// Guía de columnas interactiva. Consume las funciones y datos
// existentes (fuente de verdad: hoja Pacientes); no duplica lógica.
// ════════════════════════════════════════════════════════

var _V4_ANCHO = 440
var _V4_FICHA_ANCHO = 460

function _uiV4Shell(titulo, bodyHtml, ancho) {
  var html = '<!DOCTYPE html><html><head><base target="_top"><style>' + _uiCss() + _uiExtCss() + '</style></head><body>' +
    bodyHtml + '</body></html>'
  return HtmlService.createHtmlOutput(html).setTitle(titulo).setWidth(ancho || _V4_ANCHO)
}

// ─────────────────────────────────────────────────────────
// ─── INICIO / NAVEGACIÓN ─────────────────────────────────

function abrirDashboardV4() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Dashboard')
  if (!sh) { crearDashboard(); sh = ss.getSheetByName('Dashboard') }
  if (!sh) { ss.toast('No se pudo crear el Dashboard', 'PADDS', 4); return }
  sh.activate()
  ss.toast('Dashboard abierto', 'PADDS', 1)
}

function _irA(nombre) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(nombre)
  if (!sh) { ss.toast('No existe la hoja ' + nombre, 'PADDS', 3); return }
  sh.activate()
}

// ─────────────────────────────────────────────────────────
// ─── CENTRO DE CONTROL ───────────────────────────────────

function _datosCentroControl() {
  var out = { pacientes: 0, vigentes: 0, egresados: 0, alertasC: 0, alertasI: 0, formPend: null, hasPac: false }
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (sh && sh.getLastRow() >= 4) {
    out.hasPac = true
    var lc = Math.min(sh.getLastColumn(), 112)
    var d = sh.getRange(4, 1, sh.getLastRow() - 3, lc).getValues()
    for (var r = 0; r < d.length; r++) {
      var es = String(d[r][5] || '').trim()
      out.pacientes++
      if (es === 'VIGENTE') out.vigentes++
      else if (es === 'EGRESO' || es === 'EGRESO POR ALTA' || es === 'ALTA') out.egresados++
    }
  }
  try { var al = _calcularAlertas(); if (al) { out.alertasC = al.vencidos.length; out.alertasI = al.porVencer.length } } catch (eA) {}
  try { var est = _obtenerEstadisticas(); if (est) out.formPend = est.pend } catch (eE) {}
  return out
}

function abrirCentroControl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var d = _datosCentroControl()
  var b = []
  b.push('<div class="h">🧭 Centro de control</div>')
  b.push('<div class="sub">' + new Date().toLocaleDateString('es-CL', { weekday: 'long', day: 'numeric', month: 'long' }) + ' · PADDS V4</div>')
  if (!d.hasPac) { b.push(_uiEmpty('📭', 'Sin datos', 'La hoja ' + HOJA_PAC + ' no tiene pacientes registrados.')); return _mostrar(b) }

  b.push('<div class="grid2">' +
    _uiKpi(d.pacientes, 'Pacientes') +
    _uiKpi(d.vigentes, 'Vigentes', '') +
    _uiKpi(d.alertasC, 'Controles vencidos', '') +
    _uiKpi(d.alertasI, 'Por vencer', '') + '</div>')

  b.push('<div class="card"><div class="card-t">Acciones rápidas</div>' +
    '<button class="btn btn-teal btn-block" onclick="google.script.run.buscarPacienteGlobal()">🔎 Buscar paciente</button>' +
    '<button class="btn btn-block" onclick="google.script.run.verFichaPacienteV4()">📄 Ficha de paciente</button>' +
    '<button class="btn btn-sec btn-block" onclick="google.script.run.mostrarTareas()">✅ Tareas</button>' +
    '<button class="btn btn-sec btn-block" onclick="google.script.run.mostrarAuditoriaDatos()">🛡️ Auditoría de datos</button>' +
    '<button class="btn btn-sec btn-block" onclick="google.script.run.mostrarGuiaColumnas()">📚 Guía de columnas</button>' +
    '</div>')

  b.push('<div class="card"><div class="card-t">Hojas</div>' +
    '<button class="btn btn-sec btn-block" onclick="google.script.run._irA(\'Pacientes\')">👥 Pacientes</button>' +
    '<button class="btn btn-sec btn-block" onclick="google.script.run._irA(\'Recepción Formulario Profesional\')">📥 Recepción</button>' +
    '<button class="btn btn-sec btn-block" onclick="google.script.run._irA(\'Ingresos\')">📋 Ingresos</button>' +
    '<button class="btn btn-sec btn-block" onclick="google.script.run.irAHoy()">📅 Agenda · hoy</button>' +
    '<button class="btn btn-sec btn-block" onclick="google.script.run.abrirDashboardV4()">📊 Dashboard</button>' +
    '</div>')

  if (d.formPend != null) b.push('<div class="sub">📥 Recepción: <b>' + d.formPend + '</b> formulario(s) pendiente(s)</div>')
  b.push('<div class="sub">Consejo: usa "📄 Ficha de paciente" para ver un caso completo sin recorrer las 112 columnas.</div>')
  _mostrar(b)
}

function _mostrar(partes) {
  SpreadsheetApp.getUi().showSidebar(_uiV4Shell('Centro de control · PADDS V4', partes.join('')))
}

// ─────────────────────────────────────────────────────────
// ─── BUSCADOR GLOBAL V4 ──────────────────────────────────

function buscarPacienteGlobal() {
  var html =
'<!DOCTYPE html><html><head><base target="_top"><style>' + _uiCss() + _uiExtCss() +
'body{padding:0}#q{width:100%;box-sizing:border-box;padding:10px 12px;font-size:14px;border:2px solid #CBD5E1;border-radius:8px;outline:none;margin:12px 12px 6px}' +
'#q:focus{border-color:#0F766E}' +
'.bar-info{font-size:11px;color:#64748B;margin:0 12px 8px;text-align:right}' +
'#res{overflow-y:auto;padding:0 6px 12px}' +
'.it{background:#fff;border:1px solid #E2E8F0;border-radius:8px;padding:8px 10px;margin-bottom:6px}' +
'.it .nom{font-size:14px;font-weight:600;color:#1E293B}' +
'.it .det{font-size:12px;color:#64748B;margin:2px 0}' +
'.it .rut{font-family:Consolas,monospace;color:#0F766E}' +
'.it .ops{margin-top:6px;display:flex;gap:6px}' +
'</style></head><body>' +
'<input type="text" id="q" placeholder="Nombre, RUN, apellido o teléfono..." autofocus>' +
'<div class="bar-info" id="cta"></div>' +
'<div id="res"><div class="empty">Escribe para buscar en Pacientes</div></div>' +
'<script>' +
'var inp=document.getElementById("q"),res=document.getElementById("res"),cta=document.getElementById("cta"),_t;' +
'function _e(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}' +
'function _tag(e){if(e==="FALLECIDO")return"chip-neu";if(e==="EGRESO"||e==="EGRESO POR ALTA")return"chip-info";if(e==="SUSPENDIDO")return"chip-neu";if(e==="PENDIENTE")return"chip-warn";if(e==="ALTA"||e==="TRASLADO")return"chip-ok";return"chip-ok"}' +
'function _add(o){var el=document.createElement("div");el.className="it";' +
'var n=document.createElement("div");n.className="nom";n.textContent=_e(o.n);el.appendChild(n);' +
'var dt=document.createElement("div");dt.className="det";' +
'dt.innerHTML="<span class=\\"rut\\">"+_e(o.ru)+"</span> &nbsp; <span>"+_e(o.s)+"</span>";el.appendChild(dt);' +
'var tg=document.createElement("span");tg.className="chip "+_tag(o.e);tg.textContent=_e(o.e);dt.appendChild(tg);' +
'var ops=document.createElement("div");ops.className="ops";' +
'var b1=document.createElement("button");b1.className="btn btn-sm btn-teal";b1.textContent="📄 Ficha";b1.onclick=function(){google.script.run._abrirFichaV4(o.f)};ops.appendChild(b1);' +
'var b2=document.createElement("button");b2.className="btn btn-sm btn-sec";b2.textContent="📍 Ir a la fila";b2.onclick=function(){google.script.run._navegarAFila(o.f)};ops.appendChild(b2);' +
'el.appendChild(ops);res.appendChild(el)}' +
'inp.addEventListener("input",function(){clearTimeout(_t);_t=setTimeout(function(){' +
'var v=inp.value.trim();if(!v){res.innerHTML="<div class=\\"empty\\">Escribe para buscar</div>";cta.textContent="";return}' +
'cta.textContent="Buscando...";' +
'google.script.run.withSuccessHandler(function(r){cta.textContent=r.length+" resultado"+(r.length===1?"":"s");res.innerHTML="";' +
'if(!r.length){res.innerHTML=_uiEmptyV();return}for(var i=0;i<r.length;i++)_add(r[i])})._doBuscarPacientes(v)},250)});' +
'function _uiEmptyV(){return \'<div class="empty"><div class="ic">🔎</div><div style="font-weight:600;color:#475569">No se encontraron pacientes</div><div style="font-size:12px;margin-top:4px">Prueba por RUN, nombre o apellido</div></div>\'}' +
'</script></body></html>'
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(html).setTitle('Buscar paciente · PADDS V4').setWidth(_V4_ANCHO)
  )
}

// ─────────────────────────────────────────────────────────
// ─── FICHA DEL PACIENTE V4 ───────────────────────────────

function verFichaPacienteV4() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PAC)
  if (!sh) { SpreadsheetApp.getActiveSpreadsheet().toast('No se encontró la hoja ' + HOJA_PAC, 'Pacientes', 4); return }
  var row = sh.getActiveRange() ? sh.getActiveRange().getRow() : -1
  if (row < 4) {
    var raw = SpreadsheetApp.getUi().prompt(
      'Ficha del paciente',
      'Selecciona una fila primero, o ingresa RUN (con/sin guión) o N°:',
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL)
    if (raw.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return
    var q = raw.getResponseText().trim()
    if (!q) return
    row = _buscarFilaPaciente(sh, q)
    if (row < 4) { SpreadsheetApp.getUi().alert('Paciente no encontrado'); return }
  }
  _abrirFichaV4(row)
}

function _abrirFichaV4(fila) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh || fila < 4) return
  var lc = Math.min(sh.getLastColumn(), 112)
  var vals = sh.getRange(fila, 1, 1, lc).getValues()[0]
  var html = _filaV4Html(vals, fila, lc)
  ss.getUi().showSidebar(
    HtmlService.createHtmlOutput(html).setTitle('Ficha del paciente · PADDS').setWidth(_V4_FICHA_ANCHO)
  )
}

function _fechaV4(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return fmtFecha(v)
  var s = String(v == null ? '' : v).trim()
  if (!s || s === 'N/A') return ''
  return s
}

function _valorV4(v, col) {
  var s = String(v == null ? '' : v).trim()
  if (!s || s === 'N/A') return '<span class="st-neu">—</span>'
  if (v === true) return '<span class="st-ok">✓</span>'
  if (v === false) return '—'
  if (col && _CHECKBOX_COLS.indexOf(col) >= 0) return '<span class="st-ok">✓</span>'
  if (v instanceof Date && !isNaN(v.getTime())) return fmtFecha(v)
  if (s === 'SI') return '<span class="chip chip-ok">SI</span>'
  if (s === 'NO') return '<span class="chip chip-neu">NO</span>'
  return _esc(s)
}

function _seccionV4Html(vals, sec, lc) {
  var rows = []
  for (var c = sec.ini; c <= Math.min(sec.fin, lc); c++) {
    var col = _COLUMNAS[c]
    if (!col) continue
    var val = _valorV4(vals[c - 1], c)
    if (val === '<span class="st-neu">—</span>') continue
    rows.push('<div class="kv"><span class="k">' + _esc(col.name) + '</span><span class="v">' + val + '</span></div>')
  }
  if (!rows.length) return '<div class="sub" style="margin:6px 12px 10px">Sin información registrada</div>'
  return rows.join('')
}

function _estadoChipV4(es) {
  var map = {
    'VIGENTE': ['chip-ok', _UI_STATE.SUCCESS], 'PENDIENTE': ['chip-warn', _UI_STATE.WARNING],
    'EGRESO': ['chip-info', _UI_STATE.INFO], 'EGRESO POR ALTA': ['chip-info', _UI_STATE.INFO],
    'ALTA': ['chip-ok', _UI_STATE.SUCCESS], 'TRASLADO': ['chip-info', _UI_STATE.INFO],
    'FALLECIDO': ['chip-neu', _UI_STATE.NEUTRAL], 'SUSPENDIDO': ['chip-neu', _UI_STATE.NEUTRAL],
  }
  var m = map[es] || ['chip-neu', _UI_STATE.NEUTRAL]
  return '<span class="chip ' + m[0] + '">' + _esc(es || 'SIN ESTADO') + '</span>'
}

function _filaV4Html(vals, fila, lc) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var params = leerParametros() || {}
  var diasAviso = Number(params['DIAS_AVISO']) || 15
  var tz = ss.getSpreadsheetTimeZone()

  var nom = [vals[2], vals[3], vals[4]].map(function(x) { return String(x || '').trim() }).filter(Boolean).join(' ') || ('Fila ' + fila)
  var run = String(vals[7] || '').trim()
  var edad = String(vals[9] || '').trim()
  var sexo = String(vals[6] || '').trim()
  var sector = String(vals[1] || '').trim()
  var es = String(vals[5] || '').trim()
  var prioridad = String(vals[109] || '').trim()

  var proximo = null, cc = { v: 0, pv: 0, p: 0 }
  for (var i = 0; i < _CONTROL_FECHAS.length; i++) {
    var def = _CONTROL_FECHAS[i]
    if (def[1] > lc) continue
    var raw = vals[def[1] - 1]
    var st = _estadoFecha(raw, _mesesControl(params, def[2]), diasAviso)
    if (st === 'VENCIDO') cc.v++
    else if (st === 'POR VENCER') cc.pv++
    else if (st === 'PENDIENTE') cc.p++
    if (st === 'VENCIDO' || st === 'POR VENCER') {
      var fP = _parseDate(raw)
      var keyP = st === 'VENCIDO' ? 0 : 1
      if (!proximo || keyP < proximo.key || (keyP === proximo.key && fP && proximo.f && fP < proximo.f)) {
        proximo = { key: keyP, name: def[0], fecha: fP, status: st }
      }
    }
  }

  var b = []
  b.push('<div class="h" style="font-size:17px;line-height:1.3">' + _esc(nom) + '</div>')
  b.push('<div class="sub">' + _esc(run || 'Sin RUN') + ' · ' + _esc(edad || '?') + ' años · ' + _esc(sexo || '—') + ' · Fila ' + fila + '</div>')
  b.push('<div style="margin-bottom:12px">' + _estadoChipV4(es) +
    (sector ? ' <span class="chip chip-neu">📍 ' + _esc(sector) + '</span>' : '') +
    (prioridad === 'URGENTE' ? ' <span class="chip chip-err">🚨 URGENTE</span>' : (prioridad === 'POR REVISAR' ? ' <span class="chip chip-warn">⚠️ POR REVISAR</span>' : '')) +
    '</div>')

  b.push('<div class="grid2">' +
    '<div class="kpi"><div class="n" style="font-size:15px">' + _esc(String(vals[25] || '—')) + '</div><div class="l">Dependencia (Barthel)</div></div>' +
    '<div class="kpi"><div class="n" style="font-size:13px">' + _esc(String(vals[15] || '—')) + '</div><div class="l">Cuidador</div></div>' +
    '<div class="kpi"><div class="n" style="font-size:13px">' + (proximo ? _esc(proximo.name) : '—') + '</div><div class="l">' + (proximo ? (proximo.status === 'VENCIDO' ? '<span class="st-err">VENCIDO</span>' : '<span class="st-warn">por vencer</span>') : 'controles al día') + '</div></div>' +
    '<div class="kpi"><div class="n" style="font-size:15px">' + cc.v + ' · ' + cc.pv + ' · ' + cc.p + '</div><div class="l">Vencidos · Próx. · Sin fecha</div></div>' +
    '</div>')

  b.push('<div style="display:flex;gap:6px;margin:10px 0 12px">' +
    '<button class="btn btn-teal btn-sm" onclick="google.script.run._navegarAFila(' + fila + ')">📍 Ir a la fila</button>' +
    '<button class="btn btn-sec btn-sm" onclick="google.script.run._generarFichaPdfSidebar(' + fila + ')">📄 PDF</button>' +
    '<button class="btn btn-sec btn-sm" onclick="google.script.run.buscarPacienteGlobal()">🔎 Buscar</button>' +
    '</div>')

  for (var s = 0; s < PAC_SECCIONES.length; s++) {
    var sec = PAC_SECCIONES[s]
    if (sec.ini > lc) continue
    b.push(_uiAcc(sec.nombre, _seccionV4Html(vals, sec, lc), s === 4))
  }
  return '<style>' + _uiCss() + _uiExtCss() + '</style><body>' + b.join('') + '</body>'
}

// ─────────────────────────────────────────────────────────
// ─── ALERTAS + TAREAS V4 ─────────────────────────────────

function _tareasDesdeAlertas() {
  var r = _calcularAlertas()
  if (!r) return []
  var t = []
  var push = function(sev, control, accion, detalle, fila, label) {
    t.push({ sev: sev, control: control, accion: accion, detalle: detalle, fila: fila, label: label })
  }
  for (var i = 0; i < r.vencidos.length; i++) {
    var v = r.vencidos[i]
    push('CRITICA', v.control, 'Programar / actualizar ' + v.control, v.razon, v.fila, v.label)
  }
  for (var i2 = 0; i2 < r.porVencer.length; i2++) {
    var pv = r.porVencer[i2]
    push('IMPORTANTE', pv.control, 'Agendar ' + pv.control + ' antes del vencimiento', pv.razon, pv.fila, pv.label)
  }
  for (var i3 = 0; i3 < r.pendientes.length; i3++) {
    var p = r.pendientes[i3]
    push('PENDIENTE', p.control, 'Registrar fecha de ' + p.control, 'Sin fecha registrada', p.fila, p.label)
  }
  for (var i4 = 0; i4 < r.urgentes.length; i4++) {
    var u = r.urgentes[i4]
    push('CRITICA', 'Prioridad', 'Revisar caso prioritario', u.detalle, u.fila, u.label)
  }
  var orden = { CRITICA: 0, IMPORTANTE: 1, PENDIENTE: 2 }
  t.sort(function(a, b) { return (orden[a.sev] - orden[b.sev]) || a.label.localeCompare(b.label) })
  return t
}

function _tareasV4Html() {
  var t = _tareasDesdeAlertas()
  if (!t.length) return _uiEmpty('✅', 'Todo al día', 'No hay tareas pendientes derivadas de los datos.')
  var nC = 0, nI = 0, nP = 0
  for (var i = 0; i < t.length; i++) { if (t[i].sev === 'CRITICA') nC++; else if (t[i].sev === 'IMPORTANTE') nI++; else nP++ }
  var b = ['<div class="grid3">' +
    _uiKpi(nC, 'Críticas', '') + _uiKpi(nI, 'Importantes', '') + _uiKpi(nP, 'Pendientes', '') + '</div>']
  for (var i2 = 0; i2 < t.length; i2++) {
    var ta = t[i2]
    var st = _sevEstado(ta.sev)
    b.push('<div class="card" style="margin-bottom:8px">' +
      '<div class="row-l"><span style="font-weight:600">' + _esc(ta.label) + '</span>' +
      '<span class="badge ' + st.badge + '">' + ta.sev + '</span></div>' +
      '<div class="sub" style="margin:4px 0">' + _esc(ta.control) + (ta.detalle ? ' · ' + _esc(ta.detalle) : '') + '</div>' +
      '<div class="row-l"><span class="dim">' + _esc(ta.accion) + '</span>' +
      '<button class="btn btn-sec btn-sm" onclick="google.script.run._navegarAFila(' + ta.fila + ')">Ver</button></div>' +
      '</div>')
  }
  return b.join('')
}

function _alertasV4Html() {
  var r = _calcularAlertas()
  if (!r) return ''
  var total = r.urgentes.length + r.vencidos.length + r.porVencer.length + r.pendientes.length
  if (!total) return _uiEmpty('✅', 'Todo está al día', 'No existen alertas pendientes.')
  var b = ['<div class="grid3">' +
    _uiKpi(r.vencidos.length, 'Vencidos', '') +
    _uiKpi(r.porVencer.length, 'Por vencer', '') +
    _uiKpi(r.pendientes.length, 'Sin fecha', '') + '</div>']
  var secciones = [
    ['🚨 Prioritarios', r.urgentes, function(x) { return x.detalle }],
    ['🔴 Vencidos', r.vencidos, function(x) { return x.control + ' · ' + x.razon }],
    ['🟠 Por vencer', r.porVencer, function(x) { return x.control + ' · ' + x.razon }],
    ['🟡 Sin fecha', r.pendientes, function(x) { return x.control }],
  ]
  for (var s = 0; s < secciones.length; s++) {
    var sec = secciones[s]
    if (!sec[1].length) continue
    b.push('<div class="acc collapsed"><div class="acc-h">' + sec[0] + ' (' + sec[1].length + ')</div><div class="acc-c">')
    for (var i = 0; i < sec[1].length; i++) {
      var it = sec[1][i]
      b.push('<div class="kv"><span class="k">' + _esc(it.label) + '<div class="dim">' + _esc(sec[2](it)) + '</div></span>' +
        '<span><button class="btn btn-sec btn-sm" onclick="google.script.run._navegarAFila(' + it.fila + ')">Ver</button></span></div>')
    }
    b.push('</div></div>')
  }
  return b.join('')
}

function _mostrarAlertasTareas(abrirTareas) {
  var body = '<div class="tabs">' +
    '<button class="tab' + (abrirTareas ? '' : ' active') + '" onclick="switchT(0)">🚨 Alertas</button>' +
    '<button class="tab' + (abrirTareas ? ' active' : '') + '" onclick="switchT(1)">✅ Tareas</button></div>' +
    '<div id="tabA">' + _alertasV4Html() + '</div>' +
    '<div id="tabT" style="display:' + (abrirTareas ? '' : 'none') + '">' + _tareasV4Html() + '</div>' +
    '<script>function switchT(i){document.getElementById("tabA").style.display=i===0?"":"none";' +
    'document.getElementById("tabT").style.display=i===1?"":"none";' +
    'var ts=document.getElementsByClassName("tab");for(var x=0;x<ts.length;x++)ts[x].className="tab"+(x===i?" active":"");}</script>'
  SpreadsheetApp.getUi().showSidebar(_uiV4Shell('Alertas y tareas · PADDS V4', body))
}

function mostrarAlertasV4() {
  _mostrarAlertasTareas(false)
}

function mostrarTareas() {
  _mostrarAlertasTareas(true)
}

// ─────────────────────────────────────────────────────────
// ─── AUDITORÍA DE DATOS ──────────────────────────────────

function _auditarDatos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) { ss.toast('No se encontró la hoja ' + HOJA_PAC, 'Pacientes', 4); return null }
  var lr = sh.getLastRow()
  if (lr < 4) { ss.toast('Sin datos para auditar', 'Pacientes', 3); return null }
  var lc = Math.min(sh.getLastColumn(), 112)
  var data = sh.getRange(4, 1, lr - 3, lc).getValues()

  var out = {
    sinNombre: [], sinRun: [], runInvalidos: [], sinSector: [], sinEstado: [],
    sinCuidador: [], fechasCorruptas: [], runDuplicados: [], correctos: 0, total: 0,
  }
  var dups = {}
  for (var r = 0; r < data.length; r++) {
    var row = data[r]
    var fila = r + 4
    var nombre = [row[2], row[3], row[4]].map(function(x) { return String(x || '').trim() }).filter(Boolean).join(' ')
    var label = nombre || ('Fila ' + fila)
    out.total++
    var problemas = 0

    if (!nombre.trim()) { out.sinNombre.push({ fila: fila, label: label, detalle: 'Sin nombre ni apellidos' }); problemas++ }
    var runRaw = String(row[7] || '').trim()
    if (!runRaw) { out.sinRun.push({ fila: fila, label: label, detalle: 'RUN sin registrar' }); problemas++ }
    else {
      var fmt = formatearRUT(runRaw)
      if (!(fmt.indexOf('-') > 0 && fmt.length >= 4) || !_validarDigitoRUT(fmt)) {
        out.runInvalidos.push({ fila: fila, label: label, detalle: 'RUN: ' + fmt }); problemas++
      } else {
        if (!dups[fmt]) dups[fmt] = []
        dups[fmt].push({ fila: fila, label: label })
      }
    }
    var sector = String(row[1] || '').trim()
    if (!sector || sector === 'PENDIENTE') { out.sinSector.push({ fila: fila, label: label, detalle: 'Sector: ' + (sector || 'sin asignar') }); problemas++ }
    var estado = String(row[5] || '').trim()
    if (!estado || PAC_VALIDACIONES[6].indexOf(estado) < 0) { out.sinEstado.push({ fila: fila, label: label, detalle: 'Estado: ' + (estado || 'vacío') }); problemas++ }
    var cuidador = String(row[15] || '').trim()
    if (!cuidador) { out.sinCuidador.push({ fila: fila, label: label, detalle: 'Sin cuidador registrado' }); problemas++ }
    for (var c = 0; c < _FECHAS_VA.length; c++) {
      var fc = _FECHAS_VA[c]
      if (fc > lc) continue
      var v = row[fc - 1]
      if (v != null && String(v).trim() && String(v).trim() !== 'N/A' && !(v instanceof Date)) {
        var fcRes = _fechaCorrupta(String(v), fc === 9)
        if (fcRes) {
          out.fechasCorruptas.push({ fila: fila, label: label, detalle: 'Col ' + fc + ': "' + String(v) + '" (' + fcRes + ')' })
          problemas++
          break
        }
      }
    }
    if (!problemas) out.correctos++
  }
  for (var k in dups) if (dups[k].length > 1) out.runDuplicados.push({ fila: dups[k][0].fila, label: dups[k][0].label, detalle: 'RUN ' + k + ' en ' + dups[k].length + ' filas' })
  return out
}

function mostrarAuditoriaDatos() {
  var a = _auditarDatos()
  if (!a) return
  var b = []
  b.push('<div class="h">🛡️ Auditoría de datos</div>')
  b.push('<div class="sub">' + a.total + ' pacientes revisados · ' + a.correctos + ' sin problemas</div>')
  var grupos = [
    ['🔴 RUN inválidos', a.runInvalidos],
    ['🔴 RUN duplicados', a.runDuplicados],
    ['🟠 Fechas corruptas', a.fechasCorruptas],
    ['🟠 Sin sector', a.sinSector],
    ['🟠 Sin estado válido', a.sinEstado],
    ['🟡 Sin RUN', a.sinRun],
    ['🟡 Sin cuidador', a.sinCuidador],
    ['🟡 Sin nombre', a.sinNombre],
  ]
  var tieneAlgo = false
  for (var g = 0; g < grupos.length; g++) {
    var gr = grupos[g]
    if (!gr[1].length) continue
    tieneAlgo = true
    b.push('<div class="acc"><div class="acc-h">' + gr[0] + ' (' + gr[1].length + ')</div><div class="acc-c">')
    for (var i = 0; i < Math.min(gr[1].length, 20); i++) {
      var it = gr[1][i]
      b.push('<div class="kv"><span class="k">' + _esc(it.label) + '<div class="dim">' + _esc(it.detalle) + '</div></span>' +
        '<span><button class="btn btn-sec btn-sm" onclick="google.script.run._navegarAFila(' + it.fila + ')">Ver</button></span></div>')
    }
    if (gr[1].length > 20) b.push('<div class="dim" style="padding:4px 0">… y ' + (gr[1].length - 20) + ' más</div>')
    b.push('</div></div>')
  }
  if (!tieneAlgo) b.push(_uiEmpty('✅', 'Registros correctos', 'No se detectaron problemas de calidad de datos.'))
  SpreadsheetApp.getUi().showSidebar(_uiV4Shell('Auditoría de datos · PADDS V4', b.join('')))
}

// ─────────────────────────────────────────────────────────
// ─── GUÍA DE COLUMNAS INTERACTIVA ────────────────────────

function mostrarGuiaColumnas() {
  var b = []
  b.push('<div class="h">📚 Guía de columnas</div>')
  b.push('<div class="sub">Busca por nombre, concepto, sección o tipo. Clic en una sección para filtrar.</div>')
  b.push('<input type="text" id="gq" class="input" placeholder="Buscar columna o concepto..." style="margin-bottom:10px">')
  b.push('<div id="gfiltros" style="margin-bottom:10px"></div>')
  b.push('<div id="glist">')
  for (var c = 1; c < _COLUMNAS.length; c++) {
    var col = _COLUMNAS[c]
    if (!col || !col.name) continue
    var sec = ''
    for (var s = 0; s < PAC_SECCIONES.length; s++) {
      if (c >= PAC_SECCIONES[s].ini && c <= PAC_SECCIONES[s].fin) { sec = PAC_SECCIONES[s].nombre; break }
    }
    var txt = (col.name + ' ' + (col.desc || '') + ' ' + (col.vals || '') + ' ' + sec).toLowerCase()
    b.push('<div class="acc" data-s="' + _esc(sec) + '" data-t="' + txt + '">' +
      '<div class="acc-h">' + _esc(String(c)) + ' · ' + _esc(col.name) + '</div>' +
      '<div class="acc-c">' +
      (col.desc ? '<div style="font-size:12px;margin-bottom:4px">' + _esc(col.desc) + '</div>' : '') +
      (col.vals ? '<div class="dim">Valores: ' + _esc(col.vals) + '</div>' : '') +
      '<div style="margin-top:6px"><span class="chip chip-info">' + _esc(col.auto || 'Manual') + '</span>' +
      (sec ? '<span class="chip chip-neu">' + _esc(sec) + '</span>' : '') +
      '</div></div></div>')
  }
  b.push('</div>')
  b.push('<script>' +
    'var filtros={};var secs={};' +
    'var accs=document.querySelectorAll("#glist .acc");' +
    'accs.forEach(function(a){var s=a.getAttribute("data-s");if(s&&!secs[s])secs[s]=1});' +
    'var fz=document.getElementById("gfiltros");' +
    'Object.keys(secs).sort().forEach(function(s){var c=document.createElement("span");c.className="chip";c.textContent=s;' +
    'c.style.cursor="pointer";c.onclick=function(){var k="sec";if(filtros[k]===s){delete filtros[k];applyF();return}filtros[k]=s;applyF()};fz.appendChild(c)});' +
    'var q=document.getElementById("gq");' +
    'function applyF(){var t=q.value.toLowerCase().trim();' +
    'accs.forEach(function(a){var okT=!t||a.getAttribute("data-t").indexOf(t)>=0;' +
    'var okS=!filtros.sec||a.getAttribute("data-s")===filtros.sec;' +
    'a.style.display=(okT&&okS)?"":"none"})}' +
    'q.addEventListener("input",applyF);' +
    '</script>')
  SpreadsheetApp.getUi().showSidebar(_uiV4Shell('Guía de columnas · PADDS V4', b.join('')))
}
