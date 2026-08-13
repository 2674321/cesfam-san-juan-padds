// ════════════════════════════════════════════════════════
// ÍNDICE ▏01_Utilidades.gs │ helpers genéricos reusables

// ─────────────────────────────────────────────────────────
// ─── UTILIDADES GENÉRICAS ────────────────────────────────────────────────────

function _asignarEMPA(edad) { return edad < 20 ? 'N/A' : edad <= 64 ? 'EMPA' : 'EMPAM' }

function _parseDate(val) {
  if (typeof val === 'object' && val instanceof Date && !isNaN(val.getTime())) return val
  if (typeof val === 'number' && val > 40000) return new Date((val - 25569) * 86400000)
  if (typeof val === 'string' && val.trim()) {
    var parts = val.trim().split('/')
    if (parts.length === 3) {
      var d = +parts[0], m = +parts[1], y = +parts[2]

      if (!(d >= 1 && d <= 31 && m >= 1 && m <= 12)) return null
      if (y < 100) y += 2000
      var dd = new Date(y, m - 1, d)

      if (dd.getFullYear() === y && dd.getMonth() === m - 1 && dd.getDate() === d) return dd
    }
  }
  return null
}

function _calcularEdad(fechaNac) {
  var d = _parseDate(fechaNac)
  if (!d || isNaN(d.getTime())) return ''
  var hoy = new Date()
  var edad = hoy.getFullYear() - d.getFullYear()
  var m = hoy.getMonth() - d.getMonth()
  if (m < 0 || (m === 0 && hoy.getDate() < d.getDate())) edad--
  return edad < 0 ? '' : edad
}

// hacia abajo, de abajo hacia arriba. Nunca borra hojas ni filas con

function _borrarFilasVacias(sh, desdeFila) {
  if (!sh || desdeFila < 1) return 0
  var max = sh.getMaxRows()
  if (desdeFila > max) return 0
  var lc = sh.getLastColumn()
  if (lc < 1) return 0
  var total = max - desdeFila + 1
  var vals = sh.getRange(desdeFila, 1, total, lc).getValues()
  var formulas = sh.getRange(desdeFila, 1, total, lc).getFormulas()
  var vacias = []
  for (var r = 0; r < total; r++) {
    var esVacia = true
    for (var c = 0; c < lc; c++) {
      if (formulas[r][c] !== '') { esVacia = false; break }
      var v = vals[r][c]
      if (v != null && String(v).trim() !== '') { esVacia = false; break }
    }
    if (esVacia) vacias.push(desdeFila + r)
  }
  if (!vacias.length) return 0
  for (var i = vacias.length - 1; i >= 0; i--) sh.deleteRow(vacias[i])
  return vacias.length
}

function _normalizarSexo(v) {
  var s = String(v || '').trim().toUpperCase()
    .replace(/[ÁÀÄÂ]/g, 'A').replace(/[ÉÈËÊ]/g, 'E').replace(/[ÍÌÏÎ]/g, 'I')
    .replace(/[ÓÒÖÔ]/g, 'O').replace(/[ÚÙÜÛ]/g, 'U').replace(/Ñ/g, 'N')
  if (!s) return null
  s = s.replace(/\s+/g, ' ')
  if (s === 'F' || s === 'M') return s
  if (s === 'P' || s === 'PD' || s === 'PENDIENTE' || s.indexOf('PEND') === 0) return 'PENDIENTE'
  if (s === 'FEMENINO' || s === 'FEMININO' || s === 'FEM' || s === 'FEMALE' ||
      s === 'MUJER' || s === 'MUJERES' ||
      s.indexOf('FEMENIN') === 0 || s.indexOf('FEMININ') === 0 || s.indexOf('MUJER') === 0) return 'F'
  if (s === 'MASCULINO' || s === 'MASC' || s === 'MASC.' || s === 'MALE' ||
      s === 'HOMBRE' || s === 'HOMBRES' || s === 'VARON' || s === 'VARONES' ||
      s.indexOf('MASCULIN') === 0 || s.indexOf('HOMBRE') === 0 || s.indexOf('VARON') === 0) return 'M'
  return null
}

function _quitarAcentos(s) {
  return String(s || '')
    .replace(/[ÁÀÄÂÃ]/g, 'A').replace(/[ÉÈËÊ]/g, 'E').replace(/[ÍÌÏÎ]/g, 'I')
    .replace(/[ÓÒÖÔÕ]/g, 'O').replace(/[ÚÙÜÛ]/g, 'U').replace(/Ñ/g, 'N')
    .replace(/[áàäâã]/g, 'a').replace(/[éèëê]/g, 'e').replace(/[íìïî]/g, 'i')
    .replace(/[óòöôõ]/g, 'o').replace(/[úùüû]/g, 'u').replace(/ñ/g, 'n')
}

// "EGRESO", etc. Devuelve el valor original si no se reconoce (nunca altera

var _ESTADOS_CANONICOS = ['VIGENTE', 'FALLECIDO', 'EGRESO', 'EGRESO POR ALTA',
  'SUSPENDIDO', 'ALTA', 'TRASLADO', 'PENDIENTE']
function _normalizarVitalEstado(v) {
  var s = String(v == null ? '' : v).trim().replace(/\s+/g, ' ')
  if (!s) return s
  var up = _quitarAcentos(s).toUpperCase()
  if (_ESTADOS_CANONICOS.indexOf(up) >= 0) return up
  if (up === 'PEND.' || up === 'PEND' || up.indexOf('PEND') === 0) return 'PENDIENTE'
  if (up === 'EGRESADO' || up === 'EGRESO DEL PROGRAMA') return 'EGRESO'
  if (up.indexOf('EGRESO') === 0 && up.indexOf('ALTA') > 0) return 'EGRESO POR ALTA'
  if (up.indexOf('EGRESO') === 0) return 'EGRESO'
  if (up === 'ALTA DEL PROGRAMA' || up.indexOf('ALTA MEDICA') === 0) return 'ALTA'
  if (up === 'SUSPENDIDO TEMPORALMENTE' || up.indexOf('SUSPEND') === 0) return 'SUSPENDIDO'
  if (up.indexOf('FALLEC') === 0) return 'FALLECIDO'
  if (up.indexOf('TRASLAD') === 0) return 'TRASLADO'
  if (up.indexOf('VIGENT') === 0) return 'VIGENTE'
  return s
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

// ─── ESTADO DESDE FECHA ÚNICA ──────────────────────────────────────────────

function _estadoFecha(v, mesesMax, diasAviso) {
  if (v != null && (String(v).trim() === 'N/A' || String(v).trim() === 'NA')) return 'N/A'
  var f = _parseDate(v)
  if (!f) return 'PENDIENTE'
  return calcStatus(f, mesesMax, '', diasAviso, v)
}

function formatearRUT(rut) {
  if (rut == null || rut === '') return rut
  var s = String(rut).replace(/\./g, '').trim()
  if (!s) return rut

  var dash = s.indexOf('-')
  if (dash >= 0) {
    var num = s.slice(0, dash).replace(/[^0-9]/g, '')
    var dv = s.slice(dash + 1).replace(/[^0-9kK]/g, '').toUpperCase()
    if (num && dv) return num + '-' + dv
    return s
  }

  if (s.length < 2) return rut
  var num = s.slice(0, -1).replace(/[^0-9]/g, '')
  var dv = s.slice(-1).toUpperCase().replace(/[^0-9K]/, '')
  if (!num || !dv) return rut
  return num + '-' + dv
}

function _validarDigitoRUT(rut) {
  if (!rut || typeof rut !== 'string') return false

  var s = rut.replace(/\./g, '').replace(/\s/g, '').toUpperCase()
  if (!/^\d{1,8}-[0-9K]$/.test(s)) return false
  var parts = s.split('-')
  var numeros = parts[0]
  var dv = parts[1]

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

function _norm(s) {
  return String(s == null ? '' : s).toLowerCase()
    .replace(/[áàâä]/g, 'a').replace(/[éèêë]/g, 'e')
    .replace(/[íìîï]/g, 'i').replace(/[óòôö]/g, 'o')
    .replace(/[úùûü]/g, 'u').replace(/ñ/g, 'n').replace(/ç/g, 'c')
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

// Descombina (uno por uno, rango exacto) cualquier intervalo que CRUCE la fila n
// (getRow() <= n < getLastRow()). Congelar filas que cortan una combinación
// lanza "Debes seleccionar todas las celdas de un intervalo combinado…".
function _unmergeQueCruzaFila(sh, n) {
  try {
    var mrs = sh.getMergedRanges()
    for (var i = 0; i < mrs.length; i++) {
      var m = mrs[i]
      if (m.getRow() <= n && m.getLastRow() > n) {
        try { m.unmerge() } catch (eU) {}
      }
    }
  } catch (e0) {}
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

  for (var r = vals.length - 1; r >= 0; r--) {
    var v = String(vals[r][0] || '').trim()
    if (/^\d{2}\/\d{2}\/\d{4}/.test(v)) {
      var blockStart = r + 1

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

  try {
    var sc = CacheService.getScriptCache()
    var cached = sc.get(_PARAM_CACHE_KEY)
    if (cached) {
      _paramCache = JSON.parse(cached)
      _paramCacheTime = now
      return _paramCache
    }
  } catch(e) {}

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Parámetros')
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

function _ajustarHex(hex, dr, dg, db) {
  if (!hex || hex.charAt(0) !== '#') return hex
  var r = Math.max(0, Math.min(255, parseInt(hex.slice(1,3), 16) + dr))
  var g = Math.max(0, Math.min(255, parseInt(hex.slice(3,5), 16) + dg))
  var b = Math.max(0, Math.min(255, parseInt(hex.slice(5,7), 16) + db))
  return '#' + [r,g,b].map(function(x) { var h = x.toString(16); return h.length < 2 ? '0' + h : h }).join('')
}

function _lightenHex(hex, rAdd, gAdd, bAdd) {
  return _ajustarHex(hex, rAdd || 140, gAdd || 100, bAdd || 110)
}

function _fmtCeldaParaComparar(val) {
  if (typeof val === 'object' && val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'dd/MM/yyyy')
  }
  return String(val || '').trim()
}

function _hojaPacientesValida(ss) {
  var sh = ss.getSheetByName(HOJA_PAC)
  if (!sh) return { ok: false, msg: 'No existe la hoja "' + HOJA_PAC + '".' }
  if (sh.getMaxColumns() < (typeof _COLUMNAS !== 'undefined' && _COLUMNAS._count ? _COLUMNAS._count : 111)) {
    return { ok: false, msg: 'La hoja "' + HOJA_PAC + '" no tiene la estructura esperada (' + sh.getMaxColumns() + ' columnas; se esperan 111). Ejecuta 🩺 Pacientes → 🛠️ Mantenimiento de datos → "🎨 Formatear hoja".' }
  }
  return { ok: true, sh: sh }
}

// ─── COMPARTIDOS ENTRE ARCHIVOS ─────────────────────────────────────────────
var _NOTA_RUN_INV = '\u26a0\ufe0f RUN inv\u00e1lido: el d\u00edgito verificador no coincide'
var _COLS_TEXTO_LIBRE = [3, 4, 5, 11, 12, 16, 20, 48, 50, 62, 63, 75, 97, 99, 110]
var _PRIORIDAD_PMAP = {
  'URGENTE': _ESTADO_FECHA_COLORS['VENCIDO'],
  'POR REVISAR': _ESTADO_FECHA_COLORS['POR VENCER'],
  'AL DIA': _ESTADO_FECHA_COLORS['AL DIA'],
  'N/A': _ESTADO_FECHA_COLORS['N/A'],
}

function _agruparContiguos(cols) {
  if (!cols || !cols.length) return []
  var groups = [[cols[0], 1]]
  for (var gi = 1; gi < cols.length; gi++) {
    var prev = groups[groups.length - 1]
    if (cols[gi] === prev[0] + prev[1]) prev[1]++
    else groups.push([cols[gi], 1])
  }
  return groups
}

