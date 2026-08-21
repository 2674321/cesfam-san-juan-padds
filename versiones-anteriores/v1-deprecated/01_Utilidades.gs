// ─── UTILIDADES GENÉRICAS ────────────────────────────────────────────────────

function _asignarEMPA(edad) { return edad < 20 ? 'N/A' : edad <= 64 ? 'EMPA' : 'EMPAM' }

function _parseDate(val) {
  if (typeof val === 'object' && val instanceof Date && !isNaN(val.getTime())) return val
  if (typeof val === 'number' && val > 40000) return new Date((val - 25569) * 86400000)
  if (typeof val === 'string' && val.trim()) {
    var parts = val.trim().split('/')
    if (parts.length === 3) return new Date(+parts[2], +parts[1] - 1, +parts[0])
  }
  return null
}

function calcStatus(fecha, mesesMax, actual, diasAviso, fechaRaw) {
  if (actual === 'N/A' || actual === 'NA') return 'N/A'
  if (fechaRaw !== undefined && (String(fechaRaw).trim() === 'N/A' || String(fechaRaw).trim() === 'NA')) return 'N/A'
  if (!fecha || typeof fecha !== 'object' || isNaN(fecha.getTime())) return 'PENDIENTE'
  if (diasAviso === undefined || diasAviso === null) diasAviso = 0
  var h = new Date()
  var fechaVence = new Date(fecha)
  fechaVence.setMonth(fechaVence.getMonth() + mesesMax)
  if (h > fechaVence) return 'VENCIDO'
  if (diasAviso > 0) {
    var msRest = fechaVence.getTime() - h.getTime()
    var diasRest = Math.ceil(msRest / (1000 * 60 * 60 * 24))
    if (diasRest <= diasAviso) return 'POR VENCER'
  }
  return 'AL DIA'
}

function formatearRUT(rut) {
  if (rut == null || rut === '') return rut
  var s = String(rut).replace(/\./g, '').trim()
  if (!s) return rut
  // If already has dash, normalize DV to uppercase
  var dash = s.indexOf('-')
  if (dash >= 0) {
    var num = s.slice(0, dash).replace(/[^0-9]/g, '')
    var dv = s.slice(dash + 1).replace(/[^0-9kK]/g, '').toUpperCase()
    if (num && dv) return num + '-' + dv
    return s
  }
  // No dash: need at least 2 chars to infer DV
  if (s.length < 2) return rut
  var num = s.slice(0, -1).replace(/[^0-9]/g, '')
  var dv = s.slice(-1).toUpperCase().replace(/[^0-9K]/, '')
  if (!num || !dv) return rut
  return num + '-' + dv
}

function _validarDigitoRUT(rut) {
  if (!rut || typeof rut !== 'string') return false
  // Must have exactly one dash with digits before and 1 char after
  var dash = rut.indexOf('-')
  if (dash < 1 || dash !== rut.lastIndexOf('-')) return false
  var numeros = rut.slice(0, dash)
  var dv = rut.slice(dash + 1)
  if (!numeros || !dv || dv.length !== 1) return false
  if (!/^\d+$/.test(numeros)) return false
  if (!/^[0-9K]$/.test(dv)) return false

  var suma = 0, factor = 2
  for (var i = numeros.length - 1; i >= 0; i--) {
    suma += parseInt(numeros[i], 10) * factor
    factor = factor === 7 ? 2 : factor + 1
  }
  var resto = suma % 11
  var dvCalc = 11 - resto
  if (dvCalc === 11) return dv === '0'
  if (dvCalc === 10) return dv === 'K'
  return dv === String(dvCalc)
}

function _esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;')
}

function _pad2(n) { var s = String(n); return s.length < 2 ? '0' + s : s }

function fmtFecha(d) {
  return _pad2(d.getDate()) + '/' + _pad2(d.getMonth()+1) + '/' + d.getFullYear()
}

function colToLetter(n) {
  var s = ''
  while (n > 0) { n--; s = String.fromCharCode(65 + n % 26) + s; n = Math.floor(n / 26) }
  return s
}

function getFormulaSep() {
  var locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale() || 'es_CL'
  return locale.indexOf('en_') === 0 ? ',' : ';'
}

function limpiarBloque(sh, row, col) {
  var maxR = sh.getMaxRows()
  var numR = Math.min(_calcularRS(), maxR - row + 1)
  if (numR < 1) return
  var rng = sh.getRange(row, col, numR, PC + 1)
  var merges = rng.getMergedRanges()
  for (var i = 0; i < merges.length; i++) merges[i].breakApart()
  rng.clear()
  sh.setRowHeights(row, numR, 21)
}

function _findBloqueInicio(sh, row, col) {
  if (row === 1) return 1
  var vals = sh.getRange(1, col, row - 1, 1).getValues()
  for (var r = vals.length - 1; r >= 0; r--) {
    var v = String(vals[r][0] || '').trim()
    if (v.indexOf('Semana ') === 0 || v.indexOf('👤') === 0 || v.indexOf('Profesional') === 0) return r + 1
  }
  // Fallback: look for day headers (date pattern) to find nearest block start
  for (var r = vals.length - 1; r >= 0; r--) {
    var v = String(vals[r][0] || '').trim()
    if (/^\d{2}\/\d{2}\/\d{4}/.test(v)) {
      var blockStart = r + 1
      // Walk back to find the header (row 1 or after another day)
      while (blockStart > 1) {
        var prev = String(vals[blockStart - 2] || '').trim()
        if (/^\d{2}\/\d{2}\/\d{4}/.test(prev)) break
        blockStart--
      }
      return blockStart
    }
  }
  return -1
}

// ─── PARÁMETROS (con caché) ──────────────────────────────────────────────────

var _paramCache = null
var _paramCacheTime = 0
var _PARAM_TTL = 30000
var _PARAM_CACHE_KEY = '_pac_params'

function leerParametros() {
  var now = Date.now()
  if (_paramCache && (now - _paramCacheTime) < _PARAM_TTL) return _paramCache

  // Try ScriptCache across executions
  try {
    var sc = CacheService.getScriptCache()
    var cached = sc.get(_PARAM_CACHE_KEY)
    if (cached) {
      _paramCache = JSON.parse(cached)
      _paramCacheTime = now
      return _paramCache
    }
  } catch(e) {}

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Parametros')
  if (!sh) return {}
  var lr = sh.getLastRow()
  var data = lr > 0 ? sh.getRange(1, 1, lr, 2).getValues() : []
  var p = {}
  for (var r = 0; r < data.length; r++) {
    var name = String(data[r][0] || '').trim()
    if (!name) continue
    if (name.toUpperCase().indexOf('PARÁMETRO') === 0 || name.toUpperCase().indexOf('PARAMETRO') === 0) continue
    var key = name.toUpperCase().replace(/^VIGENCIA\s+/, '').replace(/\s+/g, ' ')
    if (key.indexOf('DIAS') === 0 || key.indexOf('DÍAS') === 0) {
      p['DIAS_AVISO'] = parseInt(data[r][1])
      continue
    }
    var val = Number(data[r][1])
    if (!isNaN(val) && val > 0) p[key] = val
  }
  _paramCache = p
  _paramCacheTime = now
  try { CacheService.getScriptCache().put(_PARAM_CACHE_KEY, JSON.stringify(p), 30) } catch(e) {}
  return p
}

// ─── FORMATEAR TELÉFONOS ─────────────────────────────────────────────────────

function formatChilePhone(str) {
  if (!str) return str
  var t = String(str).trim()
  if (t === '') return t
  var parts = t.split(/[\/,;|]+/).filter(function(p) { return p.trim() !== '' })
  var out = []
  for (var i = 0; i < parts.length; i++) {
    var digits = parts[i].replace(/\D/g, '')
    if (digits.length === 0) continue
    var f = fmtNum(digits)
    if (f) out.push(f)
  }
  return out.length > 0 ? out.join(' / ') : t
}

function fmtNum(d) {
  if (d.length === 0) return null
  if (d.length > 9 && d.indexOf('56') === 0) d = d.slice(2)
  if (d.length === 9 && d[0] === '9') return d[0] + ' ' + d.slice(1,5) + ' ' + d.slice(5)
  if (d.length === 9 && d[0] === '2') return d[0] + ' ' + d.slice(1,5) + ' ' + d.slice(5)
  if (d.length === 8) return d.slice(0,4) + ' ' + d.slice(4)
  if (d.length === 7) return d.slice(0,3) + ' ' + d.slice(3)
  if (d.length > 12) {
    var nums = splitNums(d)
    if (nums.length > 1) {
      var tmp = []
      for (var j = 0; j < nums.length; j++) {
        var sf = fmtNum(nums[j])
        if (sf) tmp.push(sf)
      }
      return tmp.length > 0 ? tmp.join(' / ') : d
    }
  }
  return d
}

function splitNums(d) {
  var r = [], i = 0
  while (i < d.length) {
    if (i + 9 <= d.length && d[i] === '9') { r.push(d.slice(i, i+9)); i += 9; continue }
    if (i + 9 <= d.length && d[i] === '2') { r.push(d.slice(i, i+9)); i += 9; continue }
    if (i + 8 <= d.length) { r.push(d.slice(i, i+8)); i += 8; continue }
    if (i + 7 <= d.length) { r.push(d.slice(i, i+7)); i += 7; continue }
    if (r.length > 0) r[r.length - 1] += d.slice(i)
    break
  }
  return r
}

function _lightenHex(hex, rAdd, gAdd, bAdd) {
  rAdd = rAdd || 140; gAdd = gAdd || 100; bAdd = bAdd || 110
  var r = parseInt(hex.slice(1,3), 16), g = parseInt(hex.slice(3,5), 16), b = parseInt(hex.slice(5,7), 16)
  r = Math.min(255, r + rAdd); g = Math.min(255, g + gAdd); b = Math.min(255, b + bAdd)
  return '#' + [r,g,b].map(function(x) { var h = x.toString(16); return h.length < 2 ? '0' + h : h }).join('')
}

function _fmtCeldaParaComparar(val) {
  if (typeof val === 'object' && val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'dd/MM/yyyy')
  }
  return String(val || '').trim()
}

