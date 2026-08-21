// ════════════════════════════════════════════════════════
// ÍNDICE ▏05_Alertas.gs │ alertas de vigencia y sidebar de resumen

// ─────────────────────────────────────────────────────────
// ─── MOTOR COMPARTIDO V4 ──────────────────────────────────────────────────
// Calcula vencidos / próximos / pendientes / urgentes de Pacientes. Lo usan
// Alertas, Tareas y el Centro de Control (no se duplica la lógica).
// Devuelve null si la hoja no existe; en el cuerpo usa ss para la zona horaria.
function _calcularAlertas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) { ss.toast('No se encontró la hoja ' + HOJA_PAC + '. Revisa que exista.', 'Pacientes', 4); return null }

  var lr = sh.getLastRow()
  if (lr < 4) { ss.toast('Sin datos de pacientes para generar alertas', 'Alertas', 2); return null }

  var cols = [2, 3, 4, 5, 6, 8, 110]
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
    if (idx[110] !== undefined) {
      var pg = String(row[idx[110]] || '').trim()
      if (pg === 'URGENTE' || pg === 'POR REVISAR') urgentFlags.push(pg)
    }

    var cc = { v:0, pv:0, p:0 }

    for (var ci = 0; ci < _CONTROL_FECHAS.length; ci++) {
      var def = _CONTROL_FECHAS[ci]
      var fc = def[1]
      if (idx[fc] === undefined) continue
      var fechaRaw = row[idx[fc]]
      var pName = def[2].toUpperCase()
      if (params['_DESACTIVADO_' + pName]) continue
      var mesesParam = params[pName] || 12
      var status = _estadoFecha(fechaRaw, mesesParam, diasAviso)
      var fechaOk = fechaRaw instanceof Date && !isNaN(fechaRaw.getTime())
      var fechaStr = fechaOk ? Utilities.formatDate(fechaRaw, tz, 'dd/MM/yyyy') : ''

      if (status === 'PENDIENTE') {
        cc.p++
        pendientes.push({ label: label, fila: fila, control: def[0], razon: 'Sin fecha', tooltip: def[0] + ': sin fecha registrada · Máx ' + _fmtPlazo(params, pName) })
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
        vencidos.push({ label: label, fila: fila, control: def[0], razon: ago, tooltip: 'Último: ' + fechaStr + ' · Máx: ' + _fmtPlazo(params, pName) + ' · Excedió por ' + diffMonths + ' meses' })
      } else if (status === 'POR VENCER') {
        cc.pv++
        var restan = Math.max(0, -diffDays)
        var restanTexto = ''
        if (restan >= 30) { var rm = Math.round(restan / 30.44); restanTexto = 'vence en ' + rm + (rm > 1 ? ' meses' : ' mes') }
        else { restanTexto = 'vence en ' + restan + (restan > 1 ? ' días' : ' día') }
        porVencer.push({ label: label, fila: fila, control: def[0], razon: restanTexto, tooltip: 'Último: ' + fechaStr + ' · Restan ' + restan + ' días · Aviso: ' + diasAviso + ' días · Máx: ' + _fmtPlazo(params, pName) })
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

  return { urgentes: urgentes, vencidos: vencidos, porVencer: porVencer, pendientes: pendientes, params: params }
}


