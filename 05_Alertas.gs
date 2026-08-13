// ════════════════════════════════════════════════════════
// ÍNDICE ▏05_Alertas.gs │ alertas de vigencia y sidebar de resumen

// ─────────────────────────────────────────────────────────
// ─── ALERTAS: RESUMEN DE VENCIDOS / PENDIENTES / URGENTES ──────────────────

function mostrarAlertas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Generando alertas…', 'PADDS', 1)
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) { ss.toast('No se encontró la hoja ' + HOJA_PAC + '. Revisa que exista.', 'Error', 4); return }

  var lr = sh.getLastRow()
  if (lr < 4) { ss.toast('Sin datos de pacientes para generar alertas', 'Alertas', 2); return }

  var cols = [2, 3, 4, 5, 6, 8, 109]
  var lc = sh.getLastColumn()
  for (var c = 0; c < _CONTROL_FECHAS.length; c++) {
    if (_CONTROL_FECHAS[c][1] <= lc) cols.push(_CONTROL_FECHAS[c][1])
  }
  cols.sort(function(a, b) { return a - b })
  var col1 = cols[0], colN = cols[cols.length - 1]
  var idx = {}
  for (var i = 0; i < cols.length; i++) idx[cols[i]] = cols[i] - col1

  var data = sh.getRange(4, col1, lr - 3, colN - col1 + 1).getValues()
  var params = leerParametros()
  var tz = ss.getSpreadsheetTimeZone()
  var hoy = new Date()
  var diasAviso = params['DIAS_AVISO'] || 30

  var urgentes = [], vencidos = [], porVencer = [], pendientes = []

  for (var r = 0; r < data.length; r++) {
    var row = data[r]
    var fila = r + 4

    var nom = String(row[idx[3]] || '').trim()
    var ape = String(row[idx[4]] || '').trim()
    var run = String(row[idx[8]] || '').trim()
    var label = (nom || ape) ? (nom + ' ' + ape).trim() : ('[Fila ' + fila + ']')
    if (run) label += ' [' + run + ']'

    var vital = String(row[idx[6]] || '').trim()
    if (vital === 'FALLECIDO' || vital === 'EGRESO' || vital === 'EGRESO POR ALTA' || vital === 'SUSPENDIDO') continue

    var sector = String(row[idx[2]] || '').trim()
    var urgentFlags = []
    if (sector === 'NARANJO') urgentFlags.push(sector)
    if (idx[109] !== undefined) {
      var pg = String(row[idx[109]] || '').trim()
      if (pg === 'URGENTE' || pg === 'POR REVISAR') urgentFlags.push(pg)
    }

    var cc = { v:0, pv:0, p:0 }

    for (var ci = 0; ci < _CONTROL_FECHAS.length; ci++) {
      var def = _CONTROL_FECHAS[ci]
      var fc = def[1]
      if (idx[fc] === undefined) continue
      var fechaRaw = row[idx[fc]]
      var pName = def[2].toUpperCase()
      var mesesParam = params[pName] || 12
      var status = _estadoFecha(fechaRaw, mesesParam, diasAviso)
      var fechaOk = fechaRaw instanceof Date && !isNaN(fechaRaw.getTime())
      var fechaStr = fechaOk ? Utilities.formatDate(fechaRaw, tz, 'dd/MM/yyyy') : ''

      if (status === 'PENDIENTE') {
        cc.p++
        pendientes.push({ label: label, fila: fila, control: def[0], razon: 'Sin fecha', tooltip: def[0] + ': sin fecha registrada · Máx ' + mesesParam + ' meses' })
        continue
      }
      if (status === 'N/A') continue

      var diffMs = fechaOk ? (hoy.getTime() - fechaRaw.getTime()) : 0
      var diffDays = fechaOk ? Math.round(diffMs / (1000 * 60 * 60 * 24)) : 0
      var diffMonths = mesesParam > 0 ? Math.round(diffDays / 30.44) : 0

      if (status === 'VENCIDO') {
        cc.v++
        var ago = ''
        if (diffDays >= 365) { var years = Math.floor(diffDays / 365); ago = 'hace ' + years + (years > 1 ? ' años' : ' año') }
        else if (diffDays >= 30) { ago = 'hace ' + diffMonths + (diffMonths > 1 ? ' meses' : ' mes') }
        else { ago = 'hace ' + diffDays + (diffDays > 1 ? ' días' : ' día') }
        vencidos.push({ label: label, fila: fila, control: def[0], razon: ago, tooltip: 'Último: ' + fechaStr + ' · Máx: ' + mesesParam + ' meses · Excedió por ' + diffMonths + ' meses' })
      } else if (status === 'POR VENCER') {
        cc.pv++
        var restan = Math.max(0, -diffDays)
        var restanTexto = ''
        if (restan >= 30) { var rm = Math.round(restan / 30.44); restanTexto = 'vence en ' + rm + (rm > 1 ? ' meses' : ' mes') }
        else { restanTexto = 'vence en ' + restan + (restan > 1 ? ' días' : ' día') }
        porVencer.push({ label: label, fila: fila, control: def[0], razon: restanTexto, tooltip: 'Último: ' + fechaStr + ' · Restan ' + restan + ' días · Aviso: ' + diasAviso + ' días · Máx: ' + mesesParam + ' meses' })
      }
    }

    if (urgentFlags.length) {
      var detalle = 'Prioridad: ' + urgentFlags.join(' + ')
      var partes = []
      if (cc.v) partes.push(cc.v + ' vencido' + (cc.v > 1 ? 's' : ''))
      if (cc.pv) partes.push(cc.pv + ' próximo' + (cc.pv > 1 ? 's' : '') + ' a vencer')
      if (cc.p) partes.push(cc.p + ' sin realizar')
      if (partes.length) detalle += ' · ' + partes.join(', ')
      urgentes.push({ label: label, fila: fila, detalle: detalle, tooltip: cc.v + ' vencidos, ' + cc.pv + ' próximos, ' + cc.p + ' sin realizar' })
    }
  }

  var _sort = function(a, b) { return a.label.localeCompare(b.label) }
  urgentes.sort(_sort); vencidos.sort(_sort); porVencer.sort(_sort); pendientes.sort(_sort)

  _showAlertasSidebar(urgentes, vencidos, porVencer, pendientes, params)
}

function _showAlertasSidebar(urgentes, vencidos, porVencer, pendientes, params) {
  var total = urgentes.length + vencidos.length + porVencer.length + pendientes.length
  var dias = params ? (params['DIAS_AVISO'] != null ? Number(params['DIAS_AVISO']) : 30) : 30
  var part = []

  part.push('<style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;margin:0;color:#222;overflow-x:hidden}' +
    '.hdr{padding:16px 16px 6px}' +
    '.hdr h2{margin:0;font-size:18px}' +
    '.hdr .sub{font-size:12px;color:#888;margin-top:2px}' +
    '.info{font-size:11px;color:#777;margin:6px 16px 10px;padding:6px 10px;background:#f8f9fa;border-radius:4px;line-height:1.6}' +
    '.info i{font-style:normal;padding:1px 5px;border-radius:3px;font-weight:600}' +
    '.i-r{color:#c62828}.i-o{color:#e65100}.i-y{color:#f57f17}' +
    'h3{font-size:13px;margin:10px 16px 4px;padding:5px 8px;border-left:3px solid #ccc;font-weight:600;cursor:pointer;user-select:none}' +
    'h3:after{content:" \\25BC";font-size:10px;float:right;opacity:.6}' +
    'h3.collapsed:after{content:" \\25B6"}' +
    '.h-r{border-color:#c62828;color:#b71c1c}' +
    '.h-o{border-color:#e65100;color:#d84315}' +
    '.h-y{border-color:#f9a825;color:#f57f17}' +
    'table{width:100%;border-collapse:collapse}' +
    'td{padding:5px 16px;border-bottom:1px solid #f0f0f0;font-size:12px;line-height:1.4}' +
    'td.l{font-weight:600;width:44%}' +
    'td.l a{color:#222;text-decoration:none;cursor:pointer}' +
    'td.l a:hover{color:#1565c0;text-decoration:underline}' +
    'td.c{color:#555;width:28%}' +
    'td.r{color:#888;width:28%;text-align:right}' +
    '.vacio{padding:30px 16px;text-align:center;color:#999;font-size:14px}' +
    '.ftr{font-size:11px;color:#bbb;text-align:center;padding:10px 16px;border-top:1px solid #eee;margin-top:8px}' +
    '</style>' +
    '<script>' +
    'function irAPaciente(fila){google.script.run._irAPaciente(fila)}' +
    'function toggleSec(el){var n=el.nextElementSibling;if(n){n.style.display=n.style.display==="none"?"":"none";el.classList.toggle("collapsed")}}' +
    '</script>')

  part.push('<div class="hdr"><h2>Alertas</h2><div class="sub">' + total + ' pendiente' + (total !== 1 ? 's' : '') + ' · ' + dias + ' días de aviso</div></div>')

  part.push('<div class="info">' +
    '<i class="i-r">🔴 Vencido</i> superó meses · ' +
    '<i class="i-o">🟠 Próximo</i> ≤ ' + dias + ' días · ' +
    '<i class="i-y">🟡 Sin fecha</i> nunca registrado · ' +
    '<i class="i-r">🔴 Prioritario</i> col B (SECTOR) o col 109 (PRIORIDAD)' +
    '</div>')

  function seccion(hCls, icono, titulo, items, cols, tooltip) {
    if (!items.length) return
    var collapsed = items.length > 10 ? ' collapsed' : ''
    part.push('<h3 class="' + hCls + collapsed + '" title="' + _esc(tooltip || '') + '" onclick="toggleSec(this)">' + icono + ' ' + titulo + '</h3>')
    part.push('<div class="tblwrap" style="display:' + (collapsed ? 'none' : '') + '"><table>')
    for (var i = 0; i < items.length; i++) {
      var it = items[i]
      part.push('<tr title="' + _esc(it.tooltip || '') + '">')
      part.push('<td class="l wrap"><a onclick="irAPaciente(' + _esc(it.fila) + ')" title="Ir a paciente">' + _esc(it.label) + '</a></td>')
      if (cols === 2) {
        part.push('<td class="c wrap" colspan="2">' + _esc(it.detalle || it.control || '') + '</td>')
      } else if (cols === 3) {
        part.push('<td class="c wrap">' + _esc(it.control) + '</td>')
        part.push('<td class="r wrap">' + _esc(it.razon) + '</td>')
      }
      part.push('</tr>')
    }
    part.push('</table></div>')
  }

  seccion('h-r', '🔴', urgentes.length + ' paciente' + (urgentes.length !== 1 ? 's' : '') + ' prioritario' + (urgentes.length !== 1 ? 's' : ''), urgentes, 2,
    'Paciente marcado Prioridad Naranjo/Urgente (B) o Prioridad General (109) = Urgente/Por Revisar')
  seccion('h-r', '🔴', vencidos.length + ' vencido' + (vencidos.length !== 1 ? 's' : ''), vencidos, 3,
    'Fecha del último control superó el máximo de meses (ver Parámetros)')
  seccion('h-o', '🟠', porVencer.length + ' próximo' + (porVencer.length !== 1 ? 's' : '') + ' a vencer', porVencer, 3,
    'Fecha dentro de los ' + dias + ' días de aviso configurados en Parámetros')
  seccion('h-y', '🟡', pendientes.length + ' sin realizar', pendientes, 2,
    'Nunca se ha registrado fecha para este control')

  if (total === 0) {
    part.push('<div class="vacio">Todo al día · sin alertas pendientes</div>')
  }

  part.push('<div class="ftr">' + new Date().toLocaleString() + '</div>')

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(part.join('')).setTitle('Alertas - PADDS 2026').setWidth(500).setHeight(540)
  )
}

function _irAPaciente(fila) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (sh) {
    sh.getRange(fila, 1).activate()
    ss.toast('Paciente en la fila ' + fila, 'PADDS', 2)
  }
}

