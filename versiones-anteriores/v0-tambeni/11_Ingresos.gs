// ─── INGRESOS: CREAR HOJA + FORMATO + PASO A PACIENTES ──────────────────────
//
// ⚠️ REGLAS DE SEGURIDAD (obligatorias, no eliminar):
//  1. ESTE CÓDIGO NUNCA ELIMINA HOJAS. Nunca. Ni esta ni ninguna otra.
//  2. Si se crea una hoja en blanco dentro de este código, solo ESA hoja puede
//     manipularse/eliminarse, y solo mientras esté en blanco. Las hojas con
//     datos (INGRESOS, LISTA DE TRABAJO, DESVIO MEDICO, AGENDA REUNIONES,
//     KINESIOLOGIA EVALUACION, Pacientes, Agenda Profesionales, etc.) NO se
//     eliminan jamás.
//  3. Este código NO inserta ni borra columnas de hojas existentes, NO
//     reordena columnas y solo reescribe valores en la utilidad "Corregir y
//     mejorar formato", que se limita a NORMALIZAR: RUT sin puntos ni
//     caracteres sueltos (ej. 1629792-5| → 1629792-5), espacios y
//     mayúsculas en nombres, y fechas escritas como texto → fecha real.
//     Nunca borra ni altera el contenido. Agrega UNA fila nueva al final
//     de Pacientes si el RUT no existe ahí. La entrada de datos en
//     INGRESOS sigue siendo MANUAL.
//  4. FLUJO AL PONER INGRESA EN OBSERVACION (disparador instalable):
//     a. Aparece una ventana de confirmación: "¿Enviar a la hoja Pacientes?"
//        — avisa que la fila se eliminará de INGRESOS al confirmar.
//     b. Se verifica el RUT con _buscarFilaPaciente: si ya existe en
//        Pacientes NO se duplica y la fila NO se elimina (queda con nota).
//     c. Si se confirma y el RUT es nuevo: se crea el paciente en Pacientes
//        y la fila se ELIMINA de INGRESOS (se elimina la fila, nunca una
//        hoja). El formato de las filas restantes se reacomoda.
//     d. Si el paciente enviado tiene OTRAS filas con el mismo RUN en
//        INGRESOS (repetidas, pendientes o marcadas NO INGRESA), se pregunta
//        si eliminarlas también (el paciente ya quedó ingresado en
//        Pacientes). Lo mismo aplica al envío masivo.
//     Esto es INDEPENDIENTE del flujo del Formulario Usuario / Profesional.
//  5. La ventana de confirmación NO puede abrirse desde el onEdit simple de
//     Apps Script: requiere el disparador instalable onEditIngresos, que se
//     activa desde el menú 📥 Ingresos → "Activar confirmación de INGRESA".
//     Sin activar, el onEdit simple solo aplica formato (no envía ni elimina).
//
//  ESTRUCTURA REAL DE LA HOJA (11 columnas, datos manuales):
//  ESTADO · FECHA SOLICITUD INGRESO · OBSERVACION · APELLIDO PATERNO ·
//  APELLIDO MATERNO · NOMBRE · RUN · DIRECCION · TELEFONO · DERIVADO POR ·
//  ANTECEDENTES
//
//  FORMATO AUTOMÁTICO: las filas en blanco que siguen a los datos quedan
//  pre-formateadas (bandas, bordes, validaciones) y cada fila editada se
//  formatea sola vía onEdit. Así las filas nuevas nunca quedan sin formato.

// ─── DETECCIÓN DE ESTRUCTURA (no asume columnas fijas) ──────────────────────
// Escanea filas 1-10 buscando la fila de encabezados que tenga RUN/RUT,
// NOMBRE y una columna de decisión (OBSERVACION o ESTADO).
// Devuelve { headerRow, cols } o null.

function _ingDetectarColumnas(sh) {
  var lc = sh.getLastColumn()
  if (lc < 2) return null
  var maxRow = Math.min(10, sh.getMaxRows())
  var data = sh.getRange(1, 1, maxRow, lc).getValues()
  var best = null
  var bestScore = 0

  for (var r = 0; r < data.length; r++) {
    var score = 0
    var cols = {
      run: 0, nombre: 0, apPaterno: 0, apMaterno: 0,
      estado: 0, accion: 0, fechaSolicitud: 0,
      direccion: 0, telefono: 0, derivado: 0, antecedentes: 0
    }
    for (var c = 0; c < data[r].length; c++) {
      var h = String(data[r][c] || '').toUpperCase()
      if (!h) continue
      if (/RUT|RUN/.test(h) && !cols.run) { cols.run = c + 1; score++ }
      if (/NOMBRE/.test(h) && !cols.nombre) { cols.nombre = c + 1; score++ }
      if (/APELLIDO/.test(h)) {
        if (!cols.apPaterno) { cols.apPaterno = c + 1; score++ }
        else if (!cols.apMaterno) { cols.apMaterno = c + 1; score++ }
      }
      if (/^ESTADO/.test(h) && !cols.estado) { cols.estado = c + 1; score++ }
      if (/OBSERVACION|OBS/.test(h) && !cols.accion) { cols.accion = c + 1; score++ }
      if (/FECHA|F\./.test(h) && /SOLICITUD|REGISTRO/.test(h) && !cols.fechaSolicitud) { cols.fechaSolicitud = c + 1; score++ }
      if (/DIRECCION/.test(h) && !cols.direccion) { cols.direccion = c + 1 }
      if (/TELEFONO|FONO/.test(h) && !cols.telefono) { cols.telefono = c + 1 }
      if (/DERIVADO/.test(h) && !cols.derivado) { cols.derivado = c + 1 }
      if (/ANTECEDENTE/.test(h) && !cols.antecedentes) { cols.antecedentes = c + 1 }
    }
    if (score > bestScore && cols.run && cols.nombre && (cols.accion || cols.estado)) {
      bestScore = score
      best = { headerRow: r + 1, cols: cols }
    }
  }
  return best
}

function _ingVacio(v) {
  return v == null || String(v).trim() === ''
}

// Valores que disparan el envío a Pacientes (columna OBSERVACION)
function _ingEsIngresa(v) {
  var s = String(v || '').trim().toUpperCase()
  return s === 'INGRESA' || s === 'INGRESO' || s === 'INGRESADO'
}

// Color de la celda OBSERVACION según su valor
function _ingColorAccion(v) {
  var s = String(v || '').trim().toUpperCase()
  if (s === 'INGRESA' || s === 'INGRESO' || s === 'INGRESADO') return '#C8E6C9'
  if (s === 'PENDIENTE') return '#FFF9C4'
  if (s.indexOf('NO INGRESA') === 0) return '#FFCDD2'
  return ''
}

// Filas de INGRESOS con el mismo RUN del paciente enviado (repetidas,
// pendientes o marcadas NO INGRESA), excluyendo las filas de exceptRows
// (array). Se usan para ofrecer eliminarlas al ingresar el paciente: el
// paciente ya quedó en Pacientes, así que las demás filas con su RUN sobran.

function _ingRepetidasMismoRut(sh, d, runN, exceptRows) {
  var res = []
  var hr = d.headerRow
  var lr = sh.getLastRow()
  if (lr <= hr || !d.cols.run) return res
  var exc = {}
  for (var x = 0; x < exceptRows.length; x++) exc[exceptRows[x]] = true
  var runs = sh.getRange(hr + 1, d.cols.run, lr - hr, 1).getValues()
  for (var i = 0; i < runs.length; i++) {
    var r = hr + 1 + i
    if (exc[r]) continue
    var rRaw = String(runs[i][0] || '').trim()
    if (!/[0-9]/.test(rRaw)) continue
    if (formatearRUT(rRaw).toUpperCase() !== String(runN).toUpperCase()) continue
    res.push(r)
  }
  return res
}

// ─── FORMATO MEJORADO (compartido por crear y mejorar la hoja) ──────────────
// Aplica: encabezados, congelado, bandas alternadas, bordes, anchos,
// validaciones, alturas y color de pestaña. Pre-formatea también las filas
// en blanco que siguen a los datos (formato automático para filas nuevas).

var _ING_PREFILAS = 500  // filas en blanco pre-formateadas bajo los datos

function _ingAplicarFormato(sh, d) {
  var hr = d.headerRow
  var lr = sh.getLastRow()
  var lc = sh.getLastColumn()
  var cols = d.cols

  // última columna con texto en la fila de encabezados (no se formatea más allá)
  var lcF = 0
  for (var c = 1; c <= lc; c++) {
    if (!_ingVacio(sh.getRange(hr, c).getValue())) lcF = c
  }
  if (!lcF) lcF = lc
  d._lcF = lcF

  // 1) Encabezados
  sh.getRange(hr, 1, 1, lcF)
    .setBackground('#2E7D5B').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontFamily('Calibri').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  sh.setRowHeight(hr, 28)
  sh.setFrozenRows(hr)

  // 2) Filas a formatear: datos + filas en blanco siguientes
  var desde = hr + 1
  var hasta = Math.min(sh.getMaxRows(), hr + _ING_PREFILAS)
  var total = hasta - desde + 1
  if (total < 1) return

  // bandas alternadas (una sola llamada setBackgrounds)
  var bgArr = []
  for (var r = desde; r <= hasta; r++) {
    var bg = (r - hr) % 2 === 1 ? '#FFFFFF' : '#F1F8E9'
    var rowArr = []
    for (var c2 = 0; c2 < lcF; c2++) rowArr.push(bg)
    bgArr.push(rowArr)
  }
  sh.getRange(desde, 1, total, lcF).setBackgrounds(bgArr)

  sh.getRange(desde, 1, total, lcF)
    .setFontFamily('Calibri').setFontSize(10)
    .setVerticalAlignment('middle').setFontColor('#202124')
  sh.setRowHeights(desde, total, 22)

  // 3) Centrar columnas cortas (RUN, OBSERVACION, ESTADO, FECHA)
  if (cols.run) sh.getRange(desde, cols.run, total, 1).setHorizontalAlignment('center')
  if (cols.accion) sh.getRange(desde, cols.accion, total, 1).setHorizontalAlignment('center')
  if (cols.estado) sh.getRange(desde, cols.estado, total, 1).setHorizontalAlignment('center')
  if (cols.fechaSolicitud) sh.getRange(desde, cols.fechaSolicitud, total, 1).setHorizontalAlignment('center')

  // 4) Validaciones (solo si no existen; permiten escritura libre)
  if (cols.accion) {
    var t1 = false
    try { t1 = sh.getRange(desde, cols.accion).getDataValidation() != null } catch (e1) {}
    if (!t1) {
      sh.getRange(desde, cols.accion, total, 1).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['INGRESA', 'NO INGRESA', 'PENDIENTE'], true).setAllowInvalid(true).build())
    }
  }
  if (cols.estado) {
    var t2 = false
    try { t2 = sh.getRange(desde, cols.estado).getDataValidation() != null } catch (e2) {}
    if (!t2) {
      sh.getRange(desde, cols.estado, total, 1).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['GESTIONADOS', 'PENDIENTES', 'EN ESPERA'], true).setAllowInvalid(true).build())
    }
  }

  // 5) Color de la celda OBSERVACION según valor (solo filas con datos)
  var hastaDatos = Math.min(lr, hasta)
  if (cols.accion && hastaDatos >= desde) {
    var vals = sh.getRange(desde, cols.accion, hastaDatos - desde + 1, 1).getValues()
    for (var r2 = 0; r2 < vals.length; r2++) {
      var bg2 = _ingColorAccion(vals[r2][0])
      if (bg2) sh.getRange(desde + r2, cols.accion).setBackground(bg2)
    }
  }

  // 6) Bordes finos en todas las filas formateadas
  sh.getRange(desde, 1, total, lcF)
    .setBorder(true, true, true, true, true, true, '#D7D7D7', SpreadsheetApp.BorderStyle.SOLID)

  // 7) Anchos de columna
  var w = {}
  if (cols.run) w[cols.run] = 110
  if (cols.nombre) w[cols.nombre] = 180
  if (cols.apPaterno) w[cols.apPaterno] = 160
  if (cols.apMaterno) w[cols.apMaterno] = 160
  if (cols.estado) w[cols.estado] = 110
  if (cols.accion) w[cols.accion] = 120
  if (cols.fechaSolicitud) w[cols.fechaSolicitud] = 130
  if (cols.direccion) w[cols.direccion] = 240
  if (cols.telefono) w[cols.telefono] = 130
  if (cols.derivado) w[cols.derivado] = 140
  if (cols.antecedentes) w[cols.antecedentes] = 320
  for (var cw in w) {
    if (w.hasOwnProperty(cw)) sh.setColumnWidth(Number(cw), w[cw])
  }

  // 8) Color de pestaña
  sh.setTabColor('#2E7D5B')
}

// ─── FORMATO DE UNA FILA EDITADA (llamado por onEdit) ───────────────────────
// Da formato a la fila al momento de editar, para que las filas nuevas
// nunca queden sin formato aunque estén fuera del rango pre-formateado.

function _ingFormatearFila(sh, row, numRows, d) {
  var lcF = d._lcF || 11
  var hr = d.headerRow
  var cols = d.cols
  var maxFilas = Math.min(numRows, 100)

  for (var i = 0; i < maxFilas; i++) {
    var r = row + i
    var bg = (r - hr) % 2 === 1 ? '#FFFFFF' : '#F1F8E9'
    var rng = sh.getRange(r, 1, 1, lcF)
    rng.setBackground(bg)
      .setFontFamily('Calibri').setFontSize(10)
      .setVerticalAlignment('middle').setFontColor('#202124')
    if (cols.run) sh.getRange(r, cols.run).setHorizontalAlignment('center')
    if (cols.estado) sh.getRange(r, cols.estado).setHorizontalAlignment('center')
    if (cols.fechaSolicitud) sh.getRange(r, cols.fechaSolicitud).setHorizontalAlignment('center')
    if (cols.accion) {
      sh.getRange(r, cols.accion).setHorizontalAlignment('center')
      var bgAcc = _ingColorAccion(sh.getRange(r, cols.accion).getValue())
      if (bgAcc) sh.getRange(r, cols.accion).setBackground(bgAcc)
    }
    sh.getRange(r, 1, 1, lcF)
      .setBorder(true, true, true, true, true, true, '#D7D7D7', SpreadsheetApp.BorderStyle.SOLID)
  }
}

// ─── CREAR HOJA INGRESOS (solo si no existe; nunca borra la existente) ───────

var _ING_HEADERS = ['ESTADO', 'FECHA SOLICITUD INGRESO', 'OBSERVACION', 'APELLIDO PATERNO',
  'APELLIDO MATERNO', 'NOMBRE', 'RUN', 'DIRECCION', 'TELEFONO', 'DERIVADO POR', 'ANTECEDENTES']
var _ING_ANCHOS = [110, 130, 120, 160, 160, 180, 110, 240, 130, 140, 320]

function crearHojaIngresos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('INGRESOS')
  if (sh) {
    ss.toast('La hoja INGRESOS ya existe: no se borra ni se recrea. Se corrige y aplica el formato mejorado.', 'INGRESOS', 6)
    corregirYMejorarIngresos()
    return
  }

  var ui = SpreadsheetApp.getUi()
  var resp = ui.alert('Crear hoja INGRESOS',
    'Se creará la hoja "INGRESOS" en blanco con 11 columnas:\n' + _ING_HEADERS.join(' · ') +
    '\n\nEn OBSERVACION: INGRESA = enviar a Pacientes automáticamente.\nEl formato mejorado y las filas nuevas ya vienen formateadas.',
    ui.ButtonSet.YES_NO)
  if (resp !== ui.Button.YES) return

  sh = ss.insertSheet('INGRESOS')
  var nCols = _ING_HEADERS.length

  // Hoja recién creada por este código: se puede recortar a su estructura
  if (sh.getMaxColumns() > nCols) sh.deleteColumns(nCols + 1, sh.getMaxColumns() - nCols)

  // Fila 1: título
  sh.getRange(1, 1, 1, nCols).merge()
  sh.getRange(1, 1).setValue('INGRESOS — LISTA DE ADMISIÓN')
  sh.getRange(1, 1).setBackground('#1F6FB2').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
  sh.setRowHeight(1, 32)

  // Fila 2: instrucciones
  sh.getRange(2, 1).setValue('REGISTRO MANUAL · OBSERVACION = INGRESA: pide confirmación y la fila pasa a Pacientes (verificación de RUT, sin duplicados) y se elimina de INGRESOS · Actívalo en el menú Ingresos · Nunca se borran hojas')
  sh.getRange(2, 1, 1, nCols).setFontStyle('italic').setFontColor('#5F6368').setFontSize(9)
  sh.setRowHeight(2, 20)

  // Fila 3: encabezados
  sh.getRange(3, 1, 1, nCols).setValues([_ING_HEADERS])

  // Formato mejorado completo (bandas, bordes, validaciones, anchos, etc.)
  var d = _ingDetectarColumnas(sh)
  if (d) {
    _ingAplicarFormato(sh, d)
  } else {
    // Respaldo mínimo (no debería ocurrir)
    sh.setFrozenRows(3)
    sh.setTabColor('#2E7D5B')
  }

  ss.toast('Hoja INGRESOS creada con formato mejorado (11 columnas, filas nuevas formateadas)', 'INGRESOS', 6)
}

// ─── CORREGIR Y MEJORAR FORMATO (menú) ───────────────────────────────────────
// 1) Normaliza errores comunes de datos (RUT, espacios, mayúsculas, fechas
//    de texto). 2) Aplica el formato mejorado. Nunca borra ni elimina nada.

// Corrige/normaliza celdas de datos. Devuelve cuántas celdas cambiaron.
// Solo normaliza: no borra contenido ni cambia el sentido de los datos.

function _ingCorregirDatos(sh, d) {
  var hr = d.headerRow
  var lr = sh.getLastRow()
  if (lr <= hr) return 0
  var cols = d.cols
  var lc = sh.getLastColumn()

  var data = sh.getRange(hr + 1, 1, lr - hr, lc).getValues()
  var corregidas = 0

  // columnas que se pasan a MAYÚSCULAS (nombres, decisiones, estados)
  var upCols = {}
  if (cols.nombre) upCols[cols.nombre] = true
  if (cols.apPaterno) upCols[cols.apPaterno] = true
  if (cols.apMaterno) upCols[cols.apMaterno] = true
  if (cols.accion) upCols[cols.accion] = true
  if (cols.estado) upCols[cols.estado] = true
  if (cols.derivado) upCols[cols.derivado] = true

  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < data[r].length; c++) {
      var v = data[r][c]
      if (v == null) continue
      var col = c + 1
      var nuevo = null

      if (typeof v === 'string') {
        // trim + colapsar espacios internos múltiples ("soto " → "soto")
        var s = v.trim().replace(/\s+/g, ' ')
        if (col === cols.run && s !== '') {
          // RUT: quita puntos, caracteres sueltos al final y deja DV en
          // mayúscula ("1629792-5|" → "1629792-5", "4.592.119-0." → "4592119-0")
          var f = formatearRUT(s)
          if (f !== s) nuevo = f
        } else if (upCols[col] && s !== '') {
          var u = s.toUpperCase()
          if (u !== s) nuevo = u
        } else if (col === cols.fechaSolicitud && s !== '') {
          // fechas escritas como texto ("6/3/2026") → fecha real
          // (notas como "PUNTA MIRA" no se tocan: _parseDate devuelve null)
          var fd = _parseDate(s)
          if (fd) nuevo = fd
        } else if (s !== v) {
          nuevo = s
        }
      }

      if (nuevo !== null && nuevo !== v) {
        data[r][c] = nuevo
        corregidas++
      }
    }
  }

  if (corregidas > 0) {
    sh.getRange(hr + 1, 1, data.length, lc).setValues(data)
  }
  return corregidas
}

function corregirYMejorarIngresos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('INGRESOS')
  if (!sh) {
    ss.toast('No existe la hoja "INGRESOS". Usa el menú "Crear hoja INGRESOS". No se tocó nada.', 'INGRESOS', 6)
    return
  }

  var d = _ingDetectarColumnas(sh)
  if (!d) {
    ss.toast('No encontré fila de encabezados (RUN/RUT · NOMBRE · OBSERVACION). No hice cambios. Envíame una captura de tu hoja para ajustar.', 'INGRESOS', 8)
    return
  }

  // 1) Corregir datos (solo normalización, nunca borra)
  var corregidas = _ingCorregirDatos(sh, d)

  // 2) Formato mejorado (recalcula colores según los valores corregidos)
  _ingAplicarFormato(sh, d)

  if (corregidas > 0) {
    ss.toast(corregidas + ' celda(s) corregida(s) y formato mejorado aplicado', 'INGRESOS', 6)
    SpreadsheetApp.getUi().alert('Corrección y formato',
      corregidas + ' celda(s) fueron corregidas:\n· RUT: puntos, espacios o caracteres sueltos al final (ej. 1629792-5| → 1629792-5)\n· Espacios y mayúsculas en nombres/estados\n· Fechas escritas como texto → fecha real\n\nEl formato mejorado quedó aplicado.',
      SpreadsheetApp.getUi().ButtonSet.OK)
  } else {
    ss.toast('Sin errores que corregir · formato mejorado aplicado', 'INGRESOS', 6)
  }
}

// Alias (compatibilidad): mejorarHojaIngresos = corregir + formato
function mejorarHojaIngresos() { corregirYMejorarIngresos() }

// ─── ORDENAR FILAS POR FECHA (opcional, pide confirmación) ───────────────────

function ordenarIngresosPorFecha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('INGRESOS')
  if (!sh) { ss.toast('No existe la hoja "INGRESOS"', 'INGRESOS', 4); return }
  var d = _ingDetectarColumnas(sh)
  if (!d) { ss.toast('No pude detectar encabezados: no se ordenó nada', 'INGRESOS', 5); return }

  var col = d.cols.fechaSolicitud
  if (!col) { ss.toast('No hay columna de fecha para ordenar', 'INGRESOS', 5); return }

  var hr = d.headerRow
  var lr = sh.getLastRow()
  if (lr <= hr) return

  var ui = SpreadsheetApp.getUi()
  var resp = ui.alert('Ordenar INGRESOS',
    '¿Ordenar las ' + (lr - hr) + ' filas de datos por la columna de fecha (columna ' + col + ')?',
    ui.ButtonSet.YES_NO)
  if (resp !== ui.Button.YES) return

  sh.getRange(hr + 1, 1, lr - hr, sh.getLastColumn())
    .sort({ column: col, ascending: true })
  ss.toast('INGRESOS ordenado por fecha', '', 4)
}

// ─── OBSERVACION = INGRESA → ENVIAR A PACIENTES ──────────────────────────────
// Llamada por onEdit (04_Eventos.gs). Independiente del flujo del formulario.
// Solo agrega UNA fila nueva al final de Pacientes si el RUT no existe ahí.
// Devuelve: 'enviado' | 'duplicado' | 'sin_rut' | 'no_estructura'

function _ingresarDesdeLista(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('INGRESOS')
  var pac = ss.getSheetByName('Pacientes')
  if (!sh || !pac) return 'no_estructura'

  var d = _ingDetectarColumnas(sh)
  if (!d || !d.cols.run || !d.cols.nombre) return 'no_estructura'
  var cols = d.cols
  var celAccion = sh.getRange(row, cols.accion || cols.estado || 1)

  var run = String(sh.getRange(row, cols.run).getValue() || '').trim()
  if (!/[0-9]/.test(run)) {
    celAccion.setNote('RUN no válido (sin dígitos): no se envió a Pacientes')
    ss.toast('RUN no válido en la fila ' + row + ': no se envió a Pacientes', 'INGRESOS', 5)
    return 'sin_rut'
  }
  var runN = formatearRUT(run)

  // 1) ¿El RUT ya existe en Pacientes? → no duplicar (verificación por RUT)
  var filaExistente = _buscarFilaPaciente(pac, runN)
  if (filaExistente > 0) {
    celAccion.setNote('Ya existe en Pacientes (fila ' + filaExistente + ') — no se duplicó')
    ss.toast('El RUT ya está en Pacientes (fila ' + filaExistente + '): no se duplicó', 'INGRESOS', 6)
    return 'duplicado'
  }

  // 2) Crear paciente nuevo (agregar fila al final de Pacientes, nunca borrar)
  var lr = pac.getLastRow()
  var fila = lr + 1
  var id = 1
  if (lr > 3) {
    var ids = pac.getRange(4, 1, lr - 3, 1).getValues()
    for (var i = 0; i < ids.length; i++) {
      var v = Number(ids[i][0])
      if (!isNaN(v) && v >= id) id = v + 1
    }
  }

  pac.insertRowsAfter(pac.getMaxRows(), 1)
  var newRow = []
  for (var c = 1; c <= 111; c++) newRow.push('')
  pac.getRange(fila, 1, 1, 111).setValues([newRow])

  var nombre = String(sh.getRange(row, cols.nombre).getValue() || '').trim().toUpperCase()
  var apP = cols.apPaterno ? String(sh.getRange(row, cols.apPaterno).getValue() || '').trim().toUpperCase() : ''
  var apM = cols.apMaterno ? String(sh.getRange(row, cols.apMaterno).getValue() || '').trim().toUpperCase() : ''
  var dir = cols.direccion ? String(sh.getRange(row, cols.direccion).getValue() || '').trim() : ''
  var tel = cols.telefono ? String(sh.getRange(row, cols.telefono).getValue() || '').trim() : ''
  var deriv = cols.derivado ? String(sh.getRange(row, cols.derivado).getValue() || '').trim() : ''
  var antec = cols.antecedentes ? String(sh.getRange(row, cols.antecedentes).getValue() || '').trim() : ''

  pac.getRange(fila, 1).setValue(id)                 // N°
  pac.getRange(fila, 6).setValue('VIGENTE')          // ESTADO
  pac.getRange(fila, 8).setValue(runN)               // RUN
  if (nombre) pac.getRange(fila, 3).setValue(nombre)
  if (apP) pac.getRange(fila, 4).setValue(apP)
  if (apM) pac.getRange(fila, 5).setValue(apM)
  if (dir) pac.getRange(fila, 11).setValue(dir.toUpperCase())  // DIRECCION
  if (tel) pac.getRange(fila, 12).setValue(tel)                // TELEFONO

  pac.getRange(fila, 13).setValue(new Date())        // F. INGRESO PADI = hoy

  // ANTECEDENTES → MORBILIDAD (col 57), solo si la fila nueva está vacía
  if (antec) pac.getRange(fila, 57).setValue(antec.toUpperCase())

  // OBSERVACIONES (col 110): origen del registro, solo si está vacía
  var notaObs = '[INGRESO desde hoja INGRESOS] ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy')
  if (deriv) notaObs += ' · Derivado por: ' + deriv
  pac.getRange(fila, 110).setValue(notaObs)

  try { pac.getRange(fila, 111).setValue(Session.getActiveUser().getEmail()) } catch (e) {} // EDITOR

  pac.getRange(fila, 1, 1, 111)
    .setFontColor('#000000').setFontWeight('normal').setFontSize(9).setVerticalAlignment('middle')

  try { _actualizarEstadosFila(fila) } catch (e) {}

  // 3) Marcar en INGRESOS (nota, sin borrar nada)
  celAccion.setNote(
    'Enviado a Pacientes · fila ' + fila + ' · ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'))

  SpreadsheetApp.flush()

  // Limpieza: la fila insertada por insertRowsAfter puede quedar vacía si no
  // se usó; se eliminan SOLO filas 100% vacías (sin datos ni fórmulas)
  try { _borrarFilasVacias(pac, 4) } catch (eV) {}

  ss.toast('Ingreso confirmado → Pacientes fila ' + fila + ' (ID ' + id + ')', 'INGRESOS', 6)
  return 'enviado'
}

// ─── CONFIRMACIÓN AL PONER INGRESA (disparador instalable) ───────────────────
// Apps Script NO permite ventanas de diálogo en el onEdit simple, por eso
// esta lógica vive en un disparador instalable activado desde el menú.
// Flujo: confirmación → verificación de RUT (sin duplicados) → envío a
// Pacientes → eliminación de la fila de INGRESOS → re-bandeado del formato.

function configurarTriggerIngresos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  var triggers = ScriptApp.getProjectTriggers()
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEditIngresos') {
      ui.alert('Ya está activo', 'La confirmación de INGRESA ya está activa.', ui.ButtonSet.OK)
      return
    }
  }

  var resp = ui.alert('Activar confirmación de INGRESA',
    'Al activar, cada vez que escribas INGRESA en OBSERVACION aparecerá la ventana "¿Enviar a la hoja Pacientes?".\nAl confirmar: se verifica el RUT (sin duplicados) y la fila se elimina de INGRESOS.\n\n¿Activar?',
    ui.ButtonSet.YES_NO)
  if (resp !== ui.Button.YES) return

  ScriptApp.newTrigger('onEditIngresos').forSpreadsheet(ss).onEdit().create()
  ui.alert('Activado', 'Confirmación activada. Ya puedes escribir INGRESA en OBSERVACION.', ui.ButtonSet.OK)
}

function onEditIngresos(e) {
  if (!e) return
  var sh = e.range.getSheet()
  if (sh.getName() !== 'INGRESOS') return

  var row = e.range.getRow()
  var numRows = e.range.getNumRows()
  if (numRows !== 1 || e.value == null) return

  var d = _ingDetectarColumnas(sh)
  var colAcc = d && (d.cols.accion || d.cols.estado)
  if (!d || !colAcc || row <= d.headerRow) return
  if (e.range.getColumn() !== colAcc) return
  if (!_ingEsIngresa(e.value)) return

  // Re-chequeo en vivo: si la celda cambió o la fila ya no existe, no hacer nada
  if (row > sh.getLastRow()) return
  var valActual = String(sh.getRange(row, colAcc).getValue() || '').trim().toUpperCase()
  if (valActual !== String(e.value).trim().toUpperCase()) return

  // Datos para mostrar en la ventana
  var nombre = d.cols.nombre ? String(sh.getRange(row, d.cols.nombre).getValue() || '').trim() : ''
  var runRaw = d.cols.run ? String(sh.getRange(row, d.cols.run).getValue() || '').trim() : ''
  var etiqueta = (nombre ? nombre.toUpperCase() : '') + (runRaw ? ' · ' + runRaw : '')

  var ui = SpreadsheetApp.getUi()
  var resp = ui.alert('Enviar a hoja Pacientes',
    '¿Enviar a la hoja Pacientes?\n' +
    (etiqueta ? 'Paciente: ' + etiqueta + '\n' : '') +
    'Al confirmar:\n· Se verificará el RUT (no se duplican pacientes).\n· La fila se ELIMINARÁ de INGRESOS.\n\n¿Continuar?',
    ui.ButtonSet.YES_NO)
  if (resp !== ui.Button.YES) return

  // Verificar RUT antes de eliminar nada
  runRaw = String(sh.getRange(row, d.cols.run).getValue() || '').trim()
  if (!/[0-9]/.test(runRaw)) {
    ui.alert('RUN no válido',
      'La fila no tiene un RUN válido (sin dígitos).\nNo se envió ni se eliminó la fila.', ui.ButtonSet.OK)
    return
  }
  var runN = formatearRUT(runRaw)
  var pac = e.source.getSheetByName('Pacientes')
  var existente = pac ? _buscarFilaPaciente(pac, runN) : -1
  if (existente > 0) {
    sh.getRange(row, colAcc).setNote('Ya existe en Pacientes (fila ' + existente + ') — no se envió ni se eliminó la fila')
    ui.alert('Paciente duplicado',
      'El RUT ' + runN + ' ya existe en Pacientes (fila ' + existente + ').\nNo se duplicó: la fila de INGRESOS se mantiene.', ui.ButtonSet.OK)
    return
  }

  var res = _ingresarDesdeLista(row)
  if (res !== 'enviado') {
    ui.alert('No se pudo enviar', 'No se eliminó la fila.', ui.ButtonSet.OK)
    return
  }

  // Eliminar SOLO la fila (nunca una hoja) y marcar para que el onEdit
  // simple no formatee la fila desplazada
  try {
    CacheService.getScriptCache().put('ING_DEL_' + row, '1', 30)
    sh.deleteRow(row)
  } catch (eDel) {
    ui.alert('Enviado (aviso)',
      'El paciente se agregó a Pacientes, pero no se pudo eliminar la fila (' + eDel.message + '). Revisa manualmente.', ui.ButtonSet.OK)
    return
  }

  // Reacomodar el formato (bandas) de las filas restantes
  try {
    if (sh.getLastRow() >= d.headerRow + 1) _ingReBandear(sh, d, row)
  } catch (eBand) {}

  // Filas repetidas/pendientes con el mismo RUN: el paciente ya quedó
  // ingresado en Pacientes, así que se ofrece eliminarlas también
  // (solo filas, nunca hojas)
  var elimPend = 0
  try {
    var reps = _ingRepetidasMismoRut(sh, d, runN, [])
    if (reps.length > 0) {
      var respP = ui.alert('Filas repetidas del mismo paciente',
        'El paciente ya fue ingresado en Pacientes y hay ' + reps.length +
        ' fila(s) más con el mismo RUN en INGRESOS (repetidas, pendientes o marcadas NO INGRESA).\n\n¿Eliminarlas también?',
        ui.ButtonSet.YES_NO)
      if (respP === ui.Button.YES) {
        reps.sort(function(a, b) { return b - a })
        var minPend = reps[reps.length - 1]
        for (var pk = 0; pk < reps.length; pk++) {
          try { sh.deleteRow(reps[pk]); elimPend++ } catch (ePD) {}
        }
        try {
          if (sh.getLastRow() >= d.headerRow + 1) _ingReBandear(sh, d, minPend)
        } catch (ePB) {}
      }
    }
  } catch (ePend) {}

  ui.alert('Enviado',
    'El paciente se agregó a Pacientes y su fila fue eliminada de INGRESOS.' +
    (elimPend > 0 ? '\n' + elimPend + ' fila(s) repetida(s) del mismo paciente también fueron eliminadas.' : ''),
    ui.ButtonSet.OK)
}

// Re-aplica bandas y colores de OBSERVACION desde una fila hacia abajo
// (útil después de eliminar filas, para que el formato quede parejo)

function _ingReBandear(sh, d, desdeFila) {
  var hr = d.headerRow
  var lcF = d._lcF || 11
  var lr = sh.getLastRow()
  if (lr < desdeFila || lr < hr + 1) return

  var bgArr = []
  for (var r = desdeFila; r <= lr; r++) {
    var bg = (r - hr) % 2 === 1 ? '#FFFFFF' : '#F1F8E9'
    var rowArr = []
    for (var c = 0; c < lcF; c++) rowArr.push(bg)
    bgArr.push(rowArr)
  }
  sh.getRange(desdeFila, 1, lr - desdeFila + 1, lcF).setBackgrounds(bgArr)

  if (d.cols.accion) {
    var vals = sh.getRange(desdeFila, d.cols.accion, lr - desdeFila + 1, 1).getValues()
    for (var i = 0; i < vals.length; i++) {
      var bgA = _ingColorAccion(vals[i][0])
      if (bgA) sh.getRange(desdeFila + i, d.cols.accion).setBackground(bgA)
    }
  }
}

// ─── ENVÍO MASIVO: todas las filas con OBSERVACION = INGRESA ─────────────────
// Útil para los datos ya cargados. Pide confirmación, no duplica y elimina
// las filas enviadas (con segunda confirmación).

function enviarIngresasAPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('INGRESOS')
  var pac = ss.getSheetByName('Pacientes')
  if (!sh || !pac) { ss.toast('Faltan las hojas INGRESOS o Pacientes', '', 5); return }

  var d = _ingDetectarColumnas(sh)
  var colAcc = d && (d.cols.accion || d.cols.estado)
  if (!d || !colAcc || !d.cols.run) {
    ss.toast('No pude detectar la estructura de INGRESOS', 'INGRESOS', 6)
    return
  }

  var hr = d.headerRow
  var lr = sh.getLastRow()
  if (lr <= hr) { ss.toast('No hay datos que enviar', 'INGRESOS', 5); return }

  var vals = sh.getRange(hr + 1, colAcc, lr - hr, 1).getValues()
  var candidatas = []
  for (var i = 0; i < vals.length; i++) {
    if (_ingEsIngresa(vals[i][0])) candidatas.push(hr + 1 + i)
  }
  if (!candidatas.length) { ss.toast('No hay filas con OBSERVACION = INGRESA', 'INGRESOS', 5); return }

  var ui = SpreadsheetApp.getUi()
  var resp = ui.alert('Enviar a Pacientes',
    'Se enviarán ' + candidatas.length + ' fila(s) con OBSERVACION = INGRESA a Pacientes.\n' +
    'Las que ya existan no se duplican.\nLas filas enviadas se ELIMINARÁN de INGRESOS (se preguntará de nuevo antes de eliminar).\n\n¿Continuar?',
    ui.ButtonSet.YES_NO)
  if (resp !== ui.Button.YES) return

  var enviadas = 0, dup = 0, sinRut = 0
  var filasEnviadas = []
  var rutsEnviados = []
  for (var j = 0; j < candidatas.length; j++) {
    var res = _ingresarDesdeLista(candidatas[j])
    if (res === 'enviado') {
      enviadas++; filasEnviadas.push(candidatas[j])
      var rRaw = String(sh.getRange(candidatas[j], d.cols.run).getValue() || '').trim()
      if (rRaw && /[0-9]/.test(rRaw)) rutsEnviados.push(formatearRUT(rRaw).toUpperCase())
    }
    else if (res === 'duplicado') dup++
    else if (res === 'sin_rut') sinRut++
  }

  var msg = 'Enviadas: ' + enviadas + ' · ya existían: ' + dup
  if (sinRut) msg += ' · sin RUT válido: ' + sinRut
  ss.toast(msg, 'INGRESOS', 8)

  if (enviadas > 0) {
    var resp2 = ui.alert('Eliminar filas',
      enviadas + ' paciente(s) enviado(s). ¿Eliminar esas filas de INGRESOS?',
      ui.ButtonSet.YES_NO)
    if (resp2 === ui.Button.YES) {
      // de abajo hacia arriba para no desalinear índices
      filasEnviadas.sort(function(a, b) { return b - a })
      for (var k = 0; k < filasEnviadas.length; k++) {
        try { sh.deleteRow(filasEnviadas[k]) } catch (eDel) {}
      }
      try {
        var d2 = _ingDetectarColumnas(sh)
        if (d2) _ingReBandear(sh, d2, filasEnviadas[filasEnviadas.length - 1])
      } catch (eBand) {}
      msg += ' · filas eliminadas: ' + enviadas
    }
  }

  // Filas repetidas/pendientes con el mismo RUN (los pacientes ya quedaron
  // ingresados en Pacientes): ofrecer eliminarlas también (solo filas,
  // nunca hojas). Se excluyen las filas enviadas de esta misma pasada.
  var elimPend = 0
  try {
    var pendsTodas = []
    for (var pr = 0; pr < rutsEnviados.length; pr++) {
      var pend = _ingRepetidasMismoRut(sh, d, rutsEnviados[pr], filasEnviadas)
      for (var pi = 0; pi < pend.length; pi++) {
        if (pendsTodas.indexOf(pend[pi]) < 0) pendsTodas.push(pend[pi])
      }
    }
    if (pendsTodas.length > 0) {
      var respP = ui.alert('Filas repetidas del mismo paciente',
        'Los pacientes enviados tienen ' + pendsTodas.length +
        ' fila(s) más con el mismo RUN en INGRESOS (repetidas, pendientes o marcadas NO INGRESA).\n\n¿Eliminarlas también?',
        ui.ButtonSet.YES_NO)
      if (respP === ui.Button.YES) {
        pendsTodas.sort(function(a, b) { return b - a })
        var minPend = pendsTodas[pendsTodas.length - 1]
        for (var pk2 = 0; pk2 < pendsTodas.length; pk2++) {
          try { sh.deleteRow(pendsTodas[pk2]); elimPend++ } catch (ePD2) {}
        }
        try {
          var d3 = _ingDetectarColumnas(sh)
          if (d3) _ingReBandear(sh, d3, minPend)
        } catch (ePB2) {}
      }
    }
  } catch (ePend2) {}
  if (elimPend > 0) msg += ' · pendientes eliminadas: ' + elimPend

  ui.alert('Resultado', msg, ui.ButtonSet.OK)
}

