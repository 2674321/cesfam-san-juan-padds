// ════════════════════════════════════════════════════════
// ÍNDICE ▏01_Utilidades.gs │ helpers genéricos reusables

// ─────────────────────────────────────────────────────────
// ─── UTILIDADES GENÉRICAS ────────────────────────────────────────────────────

function _asignarEMPA(edad) { return edad < 20 ? 'N/A' : edad <= 64 ? 'EMPA' : 'EMPAM' }

// Mayúsculas para nombres/apellidos: colapsa espacios y normaliza
// (convención de la hoja: nombres en MAYÚSCULAS).
function _mayusNombre(v) {
  if (v == null) return ''
  return String(v).trim().replace(/\s+/g, ' ').toUpperCase()
}

// Texto amigable del plazo configurado en Parámetros ("30 días", "6 meses"
// o "N/A" si la vigencia está desactivada).
function _fmtPlazo(params, key) {
  var k = String(key || '').toUpperCase()
  if (params && params['_DESACTIVADO_' + k]) return 'N/A'
  var dias = params && params['_DIAS_' + k]
  if (dias) return Math.round(Number(dias)) + ' días'
  var m = params && params[k]
  if (m === undefined || m === null || isNaN(m)) m = _mesesControl(params, k)
  if (m === null || m === undefined) return 'N/A'
  var n = Number(m)
  return (n === Math.round(n) ? n : Math.round(n * 10) / 10) + ' meses'
}

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

// ─── DIVISIÓN AUTOMÁTICA NOMBRE / APELLIDO / APELLIDO 2 ─────────────────────
// Separa en NOMBRE + APELLIDO (paterno) + APELLIDO 2 (materno).
// - Si hay apellido(s) explícito(s): el apellido con 2+ palabras se parte en
//   paterno + materno, y el exceso de palabras del nombre (3+) baja a los
//   apellidos que falten.
// - Si todo viene en el campo nombre (modo "un solo campo"): regla simple por
//   cantidad de palabras: 2 → 1 nombre + 1 apellido; 3 → 1 nombre + 2 apellidos;
//   4+ → 2 nombres + 2 apellidos (lo que sobre va a APELLIDO 2).
function _dividirNombreApellidos(nombre, apPaterno, apMaterno) {
  var n = String(nombre || '').toUpperCase().replace(/\s+/g, ' ').trim()
  var a1 = String(apPaterno || '').toUpperCase().replace(/\s+/g, ' ').trim()
  var a2 = String(apMaterno || '').toUpperCase().replace(/\s+/g, ' ').trim()

  // Caso normal: hay apellido(s) explícito(s).
  if (a1 || a2) {
    // Apellido con 2+ palabras y sin materno → la 1ª es paterno, el resto materno.
    if (!a2 && a1.indexOf(' ') > 0) {
      var ta = a1.split(' ')
      a1 = ta.shift()
      a2 = ta.join(' ')
    }
    // Nombre con 3+ palabras → las 2 primeras son nombre; el resto baja a los
    // apellidos: paterno si falta, si no materno.
    var tn = n.split(' ')
    if (tn.length > 2 && tn[0]) {
      n = tn.slice(0, 2).join(' ')
      var sobra = tn.slice(2).join(' ')
      if (!a1) a1 = sobra
      else if (!a2) a2 = sobra
      else a2 = (a2 + ' ' + sobra).trim()
    }
    return { nombre: n, apellido: a1, apellido2: a2 }
  }

  // Modo "un solo campo": regla simple por cantidad de palabras.
  var tp = n.split(' ')
  if (tp.length < 2 || !tp[0]) return { nombre: n, apellido: '', apellido2: '' }
  var numNombres = tp.length >= 4 ? 2 : 1
  return {
    nombre: tp.slice(0, numNombres).join(' '),
    apellido: tp[numNombres] || '',
    apellido2: tp.slice(numNombres + 1).join(' ') || '',
  }
}

function _borrarFilasVacias(sh, desdeFila) {
  if (!sh || desdeFila < 1) return 0
  var max = sh.getMaxRows()
  if (desdeFila > max) return 0
  var lc = sh.getLastColumn()
  if (lc < 1) return 0
  var total = max - desdeFila + 1
  var vals = sh.getRange(desdeFila, 1, total, lc).getValues()
  var formulas = sh.getRange(desdeFila, 1, total, lc).getFormulas()
  var runs = []
  var ini = -1
  for (var r = 0; r <= total; r++) {
    var esVacia = true
    if (r < total) {
      for (var c = 0; c < lc; c++) {
        if (formulas[r][c] !== '') { esVacia = false; break }
        var v = vals[r][c]
        if (v == null) continue
        var s = String(v).trim()
        // Celdas con validación de casilla (checkbox) desmarcadas devuelven
        // FALSE: en la hoja se ven desmarcadas (sin dato), así que NO cuentan
        // como contenido. El resto (incluido "N/A", que es un dato real) sí.
        if (s === '' || s === 'FALSE' || s === 'false' || v === false) continue
        esVacia = false
        break
      }
    }
    if (esVacia) {
      if (ini < 0) ini = r
    } else {
      if (ini >= 0) {
        runs.push([desdeFila + ini, r - ini])
        ini = -1
      }
    }
  }
  if (!runs.length) return 0
  var borradas = 0
  for (var i = runs.length - 1; i >= 0; i--) {
    try {
      sh.deleteRows(runs[i][0], runs[i][1])
      borradas += runs[i][1]
    } catch (eD) {
      // Si el borrado falla (p. ej. filas dentro de un rango combinado),
      // se quita la combinación y se reintenta; si aun así no se puede,
      // se ignora la fila (nunca se tocan datos).
      try {
        sh.getRange(runs[i][0], 1, runs[i][1], 1).breakApart()
        sh.deleteRows(runs[i][0], runs[i][1])
        borradas += runs[i][1]
      } catch (eD2) {}
    }
  }
  return borradas
}

// Barrido con reintentos: algunas filas "vacías" resisten el borrado la
// primera vez (rangos combinados, etc.). Repite hasta que no quede ninguna.
function _limpiarFilasVaciasLoop(sh, desdeFila) {
  var total = 0
  for (var ronda = 0; ronda < 3; ronda++) {
    var n = _borrarFilasVacias(sh, desdeFila)
    total += n
    if (n === 0) break
  }
  return total
}

// Compacta la hoja Pacientes: borra TODA fila que esté debajo de la última
// fila con ID (columna A), porque en Pacientes cada fila es un paciente y
// cualquier fila sin ID al final es basura (casillas desmarcadas, N/A sueltos,
// filas huérfanas de versiones antiguas). Nunca toca un paciente.
function _compactarPacientes(sh) {
  if (!sh) return 0
  var lr = sh.getLastRow()
  if (lr < 4) return 0
  var ids = sh.getRange(4, 1, lr - 3, 1).getValues()
  var ultimoIdx = -1
  for (var i = ids.length - 1; i >= 0; i--) {
    var v = Number(ids[i][0])
    if (!isNaN(v) && v > 0) { ultimoIdx = i; break }
  }
  if (ultimoIdx < 0) return 0
  var ultimaFila = 4 + ultimoIdx
  if (ultimaFila >= lr) return 0
  var n = lr - ultimaFila
  try { sh.deleteRows(ultimaFila + 1, n); return n } catch (eC) {
    try {
      sh.getRange(ultimaFila + 1, 1, Math.min(n, 1), 1).breakApart()
      sh.deleteRows(ultimaFila + 1, n)
      return n
    } catch (eC2) { return 0 }
  }
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
  if (mesesMax === null || mesesMax === undefined) return 'N/A'
  if (actual === 'N/A' || actual === 'NA') return 'N/A'
  if (fechaRaw !== undefined && (String(fechaRaw).trim() === 'N/A' || String(fechaRaw).trim() === 'NA')) return 'N/A'
  if (!fecha || typeof fecha !== 'object' || isNaN(fecha.getTime())) return 'PENDIENTE'
  if (diasAviso === undefined || diasAviso === null) diasAviso = 0
  var h = new Date()
  var fechaVence
  if (mesesMax >= 1 && mesesMax === Math.round(mesesMax)) {
    fechaVence = new Date(fecha)
    fechaVence.setMonth(fechaVence.getMonth() + mesesMax)
  } else {
    fechaVence = new Date(fecha.getTime() + mesesMax * 30.44 * 86400000)
  }
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
  if (mesesMax === null || mesesMax === undefined) return 'N/A'
  if (v != null) {
    var tE = String(v).trim().toUpperCase()
    if (tE === 'N/A' || tE === 'NA') return 'N/A'
  }
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
  // RUN sin guion de 7 dígitos: falta el dígito verificador, no se puede
  // inventar (formatearRUT('1629792') sería un RUT corrupto '162979-2').
  if (/^\d{7}$/.test(s)) return s
  var num = s.slice(0, -1).replace(/[^0-9]/g, '')
  var dv = s.slice(-1).toUpperCase().replace(/[^0-9K]/, '')
  if (!num || !dv) return rut
  return num + '-' + dv
}

// Calcula el dígito verificador de un RUN (solo la parte numérica).
function _calcularDV(numero) {
  var suma = 0, factor = 2
  for (var i = numero.length - 1; i >= 0; i--) {
    suma += parseInt(numero.charAt(i), 10) * factor
    factor = factor === 7 ? 2 : factor + 1
  }
  var dvCalc = 11 - (suma % 11)
  if (dvCalc === 11) return '0'
  if (dvCalc === 10) return 'K'
  return String(dvCalc)
}

function _validarDigitoRUT(rut) {
  if (!rut || typeof rut !== 'string') return false

  var s = rut.replace(/\./g, '').replace(/\s/g, '').toUpperCase()
  if (!/^\d{1,8}-[0-9K]$/.test(s)) return false
  var parts = s.split('-')
  return parts[1] === _calcularDV(parts[0])
}

// Nota (tooltip) para una celda RUN: null si es válido o incompleto; si el
// dígito verificador no coincide, mensaje con el dígito correcto.
function _notaRUN(rut) {
  if (!rut || typeof rut !== 'string') return null
  var s = rut.replace(/\./g, '').replace(/\s/g, '').toUpperCase()
  if (!/^\d{1,8}-[0-9K]$/.test(s)) return null
  if (_validarDigitoRUT(s)) return null
  return '\u26a0\ufe0f RUN inv\u00e1lido: el d\u00edgito verificador deber\u00eda ser ' + _calcularDV(s.split('-')[0])
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
  rng.clearDataValidations()
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
  var data = lr > 0 ? sh.getRange(1, 1, lr, 4).getValues() : []
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
    if (isNaN(val) || val <= 0) continue
    var unidad = String(data[r][2] || '').trim().toUpperCase()
    if (unidad === 'N/A') {
      p['_DESACTIVADO_' + key] = true
      continue
    }
    if (unidad === 'DIAS') {
      p[key] = val / 30.44
      p['_DIAS_' + key] = val
    } else {
      p[key] = val
    }
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
  try { _asegurarColumnaInsulino(sh) } catch(e) {}
  if (sh.getMaxColumns() < 112) {
    return { ok: false, msg: 'La hoja "' + HOJA_PAC + '" no tiene la estructura esperada (' + sh.getMaxColumns() + ' columnas; se esperan al menos 112). Ejecuta 🩺 Pacientes → 🛠️ Datos → "🎨 Formatear hoja".' }
  }
  return { ok: true, sh: sh }
}

// Inserta la columna checkbox 'INSULINO DEPENDIENTE' (columna 56) si aún no
// existe en la hoja. Al insertarla en medio se desplazan las columnas siguientes
// una posición hacia la derecha conservando los datos. Devuelve true si insertó.
function _asegurarColumnaInsulino(sh) {
  if (!sh) return false
  var n = sh.getMaxColumns()
  if (n < 55) return false
  try {
    var hdr = String(sh.getRange(3, 56).getValue() || '').trim().toUpperCase()
    if (hdr.indexOf('INSULINO') >= 0) return false
  } catch(e) { return false }
  sh.insertColumnAfter(55)
  try { sh.getRange(3, 56).setValue('INSULINO DEPENDIENTE') } catch(e2) {}
  return true
}

// ─── COMPARTIDOS ENTRE ARCHIVOS ─────────────────────────────────────────────
var _COLS_TEXTO_LIBRE = [3, 4, 5, 11, 12, 16, 20, 48, 50, 63, 64, 76, 98, 100, 111]
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

// ─── LOG EN HOJA 'Log' (trazabilidad sin abrir el editor — AGENTS.md §3) ──────
// Registra (timestamp, módulo, evento, resultado, detalle) por lote con lock.
// La pestaña 'Log' se crea sola la primera vez y se recorta al crecer.
// Nunca lanza errores: el logging jamás debe romper el flujo principal.

var _LOG_MAX_FILAS = 2000

function _log(ss, pagina, evento, resultado, detalle) {
  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet()
    if (!ss) return
    var sh = ss.getSheetByName('Log')
    if (!sh) {
      sh = ss.insertSheet('Log')
      sh.setFrozenRows(1)
      sh.getRange(1, 1, 1, 5).setValues([['FECHA', 'MODULO', 'EVENTO', 'RESULTADO', 'DETALLE']])
      sh.setColumnWidths(1, 5, 110)
    }
    var lock = LockService.getScriptLock()
    if (!lock.tryLock(3000)) { console.warn('_log: sin lock, se omite el registro'); return }
    var ts = ''
    try { ts = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm:ss') } catch (eT) { ts = String(new Date()) }
    sh.appendRow([ts, String(pagina || ''), String(evento || ''), String(resultado || ''), String(detalle || '')])
    var sobra = sh.getLastRow() - _LOG_MAX_FILAS
    if (sobra > 1) sh.deleteRows(2, sobra)
    lock.releaseLock()
  } catch (eLog) {
    console.error('_log: ' + eLog.message)
  }
}
