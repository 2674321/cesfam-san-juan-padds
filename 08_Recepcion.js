
// ════════════════════════════════════════════════════════
// ÍNDICE ▏08_Recepcion.gs │ formulario profesional, aprobación de formularios, estadísticas

// ─────────────────────────────────────────────────────────

// Normaliza un RUN para comparación: quita puntos/espacios; si tiene guion lo
// re-formatea (12345678-5); si son solo dígitos (falta DV) se deja tal cual
// (formatearRUT lo convertiría en un DV incorrecto).
function _normRUN(v) {
  var s = String(v || '').replace(/\./g, '').replace(/\s/g, '').toUpperCase()
  if (!s) return ''
  if (s.indexOf('-') >= 0) return formatearRUT(s)
  if (/^\d{7,8}$/.test(s)) return s
  return formatearRUT(s)
}

var _ESTADO_CSS = {
  'Pendiente': { badge: '#C2410C', tint: '#FFEDD5', fg: '#FFFFFF' },
  'Gestionado': { badge: '#15803D', tint: '#DCFCE7', fg: '#FFFFFF' },
  'Rechazado': { badge: '#B91C1C', tint: '#FEE2E2', fg: '#FFFFFF' },
}

// Normaliza el valor de ESTADO (columna C) de la hoja de recepción: tolera
// minúsculas, mayúsculas, espacios y estados legados ("Aprobado"). Devuelve el
// canónico ('Pendiente' | 'Gestionado' | 'Rechazado') o el texto original
// recortado si no es ninguno de ellos. Sin esto, escribir el estado a mano
// (p. ej. "gestionado") hacía que Aprobar/Rechazar/Limpiar no detectaran la fila.
function _normEstadoForm(v) {
  var s = String(v == null ? '' : v).trim()
  if (!s) return s
  var up = s.toUpperCase()
    .replace(/[ÁÀÄÂ]/g, 'A').replace(/[ÉÈËÊ]/g, 'E').replace(/[ÍÌÏÎ]/g, 'I')
    .replace(/[ÓÒÖÔ]/g, 'O').replace(/[ÚÙÜÛ]/g, 'U')
  if (up.indexOf('GESTIONAD') === 0 || up.indexOf('APROB') === 0) return 'Gestionado'
  if (up.indexOf('RECHAZAD') === 0) return 'Rechazado'
  if (up.indexOf('PEND') === 0) return 'Pendiente'
  return s
}

var _FORM_SEC_TINTS = {
  1:'#F1F5F9', 2:'#E2E8F0', 3:'#F1F5F9', 4:'#E2E8F0', 5:'#F1F5F9',
  6:'#E0F2FE', 7:'#F0FDFA', 8:'#E0F2FE', 9:'#F0FDFA',
  10:'#DCFCE7',11:'#F0FDFA',12:'#DCFCE7',13:'#F0FDFA',14:'#DCFCE7',15:'#F0FDFA',16:'#DCFCE7',
  17:'#F3E8FF', 18:'#F3E8FF', 19:'#F5F3FF',
}

var _ALIGN = {
  1: 'center', 2: 'center', 3: 'center', 4: 'left', 5: 'center',
  6: 'center', 7: 'center', 8: 'left', 9: 'left', 10: 'left', 11: 'left',
  12: 'left', 13: 'center', 14: 'left', 15: 'left', 16: 'center', 17: 'left',
  18: 'center', 19: 'center',
}

function _estiloFila(sh, row, lc) {
  var even = row % 2 === 0
  var rng = sh.getRange(row, 1, 1, lc)
  var aligns = []
  var fmts = []
  for (var c = 1; c <= lc; c++) {
    aligns.push(_ALIGN[c] || 'center')
    fmts.push(c === 1 || c === 5 ? 'dd/mm/yyyy hh:mm' : (c === 6 || c === 16 ? 'dd/mm/yyyy' : 'General'))
  }
  rng
    .setFontFamily(_UI.font).setFontSize(10).setFontColor('#212121')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, _UI.border, SpreadsheetApp.BorderStyle.SOLID)
    .setBackground(even ? '#FFFFFF' : '#F8FAFC')
    .setHorizontalAlignments([aligns])
    .setNumberFormats([fmts])
  sh.setRowHeight(row, 26)
  sh.getRange(row, 17).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
}

function _colorEstado(sh, row, estado, lc) {
  var css = _ESTADO_CSS[estado]
  if (!css) return
  sh.getRange(row, 3).setBackground(css.badge).setFontColor(css.fg).setFontWeight('bold')
}

function _miniResumenForm(sh, force) {
  // El panel ya construido no necesita reconstruirse (sus contadores son
  // fórmulas vivas): onFormSubmit y los flujos de aprobar/rechazar/limpiar
  // lo saltan; solo force=true (creación/formateo) lo reconstruye.
  if (!force && PropertiesService.getScriptProperties().getProperty('_recPanelOk') === '1') return
  var lc = sh.getLastColumn() || 19
  var SEP = getFormulaSep()
  var BS_MED = SpreadsheetApp.BorderStyle.SOLID_MEDIUM

  try { sh.getRange(1, 1, 3, Math.max(lc, 20)).breakApart() } catch (eB) {}

  var t
  try { t = sh.getRange(1, 1, 1, lc).merge() } catch (eM) { t = sh.getRange(1, 1) }
  t.setBackground(_UI.hdrBg)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  var txtTit = 'REGISTRO DE SOLICITUDES\nPrograma PADDS · Recepción y gestión de formularios profesionales'
  var nlTit = txtTit.indexOf('\n')
  var okRT = false
  try {
    var rtTit = SpreadsheetApp.newRichTextValue().setText(txtTit)
      .setTextStyle(0, nlTit, SpreadsheetApp.newTextStyle().setForegroundColor('#FFFFFF')
        .setFontSize(17).setBold(true).setFontFamily(_UI.font).build())
      .setTextStyle(nlTit + 1, txtTit.length, SpreadsheetApp.newTextStyle().setForegroundColor(_UI.hdrSub)
        .setFontSize(10).setFontFamily(_UI.font).build())
      .build()
    t.setRichTextValue(rtTit)
    okRT = true
  } catch (eRT) {}
  if (!okRT) {
    t.setValue('REGISTRO DE SOLICITUDES')
      .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(15).setFontFamily(_UI.font)
  }
  t.setBorder(true, true, true, true, false, false, '#1E293B', BS_MED)
  t.setNote('Los formularios llegan automáticamente. Revise los datos, corrija si es necesario, ' +
    'luego use el menú Formulario: "Aprobar" (copia a Pacientes y elimina la fila) o "Rechazar".')
  sh.setRowHeight(1, 44)

  var cards = [
    { label: 'TOTAL SOLICITUDES', col1: 1,  col2: 5,  color: '#0F766E', tint: '#EEF2F7',
      f: 'COUNTA(G5:G)' },
    { label: 'PENDIENTES',        col1: 6,  col2: 10, color: '#C2410C', tint: '#FFEDD5',
      f: 'COUNTIF(C5:C' + SEP + '"Pendiente")' },
    { label: 'GESTIONADOS',        col1: 11, col2: 15, color: '#15803D', tint: '#DCFCE7',
      f: 'COUNTIF(C5:C' + SEP + '"Gestionado")' },
    { label: 'RECHAZADOS',        col1: 16, col2: 19, color: '#B91C1C', tint: '#FEE2E2',
      f: 'COUNTIF(C5:C' + SEP + '"Rechazado")' },
  ]
  for (var i = 0; i < cards.length; i++) {
    var ca = cards[i]
    if (ca.col1 > lc) continue
    var nCols = Math.min(ca.col2, lc) - ca.col1 + 1
    var rLab = sh.getRange(2, ca.col1, 1, nCols).merge()
    var labTxt = '● ' + ca.label
    var rtLab = SpreadsheetApp.newRichTextValue().setText(labTxt)
      .setTextStyle(0, 1, SpreadsheetApp.newTextStyle().setForegroundColor(ca.color).setBold(true)
        .setFontSize(8).setFontFamily(_UI.font).build())
      .setTextStyle(2, labTxt.length, SpreadsheetApp.newTextStyle().setForegroundColor('#64748B').setBold(true)
        .setFontSize(8).setFontFamily(_UI.font).build())
      .build()
    rLab.setRichTextValue(rtLab)
    rLab.setBackground(ca.tint).setHorizontalAlignment('center').setVerticalAlignment('middle')
    var rNum = sh.getRange(3, ca.col1, 1, nCols).merge()
    var aNum = sh.getRange(3, ca.col1)
    try { aNum.setFormula(ca.f) } catch (eF) { aNum.setValue(0) }
    aNum.setNumberFormat('0')
      .setFontFamily(_UI.font).setFontSize(20).setFontWeight('bold')
      .setFontColor(ca.color).setBackground('#FFFFFF')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
    sh.getRange(2, ca.col1, 2, nCols)
      .setBorder(true, true, true, true, false, false, ca.color, BS_MED)
  }
  sh.setRowHeight(2, 20)
  sh.setRowHeight(3, 36)
  PropertiesService.getScriptProperties().setProperty('_recPanelOk', '1')
}

function _obtenerEstadisticas() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_FORM)
  if (!sh) return null
  var lr = sh.getLastRow()
  if (lr < 5) return { total: 0, pend: 0, aprob: 0, rech: 0 }
  var data = sh.getRange(5, 1, lr - 4, 3).getValues()
  var stats = { total: 0, pend: 0, aprob: 0, rech: 0 }
  for (var r = 0; r < data.length; r++) {
    var est = _normEstadoForm(data[r][2])
    if (!est) continue
    stats.total++
    if (est === 'Pendiente') stats.pend++
    else if (est === 'Gestionado') stats.aprob++
    else if (est === 'Rechazado') stats.rech++
  }
  return stats
}

function mostrarSidebarEstadisticas() {
  var stats = _obtenerEstadisticas()
  if (!stats) { SpreadsheetApp.getUi().alert('Primero crea la hoja de recepción (📝 Formulario → "📋 Crear hoja de recepción").'); return }
  var html = '<html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:0;padding:16px;color:#202124;font-size:14px;background:#fff}' +
    'h2{margin:0 0 4px 0;font-size:16px;font-weight:500;color:#1E293B}' +
    '.sub{font-size:12px;color:#64748B;margin-bottom:16px}' +
    '.card{border:1px solid #E2E8F0;border-radius:8px;padding:12px;margin-bottom:12px}' +
    '.total-num{font-size:36px;font-weight:300;color:#1E293B}' +
    '.total-label{font-size:12px;color:#64748B}' +
    '.row{display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #F1F5F9}' +
    '.row:last-child{border-bottom:none}' +
    '.count{font-size:18px;font-weight:500}' +
    '.tag{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:8px}' +
    '.tag-pend{background:#C2410C} .tag-aprob{background:#15803D} .tag-rech{background:#B91C1C}' +
    '.label-name{color:#3c4043;display:flex;align-items:center}' +
    '.footer{font-size:11px;color:#94A3B8;text-align:center;padding-top:12px;border-top:1px solid #E2E8F0;margin-top:12px}' +
    '</style></head><body>' +
    '<h2>Recepción formularios PADDS</h2>' +
    '<div class="sub" id="updated">Actualizado: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss') + '</div>' +
    '<div class="card"><div class="total-num" id="totalNum">' + stats.total + '</div><div class="total-label">Formularios recibidos</div></div>' +
    '<div class="card"><div style="font-size:12px;color:#64748B;margin-bottom:8px">Estado</div>' +
    '<div class="row"><span class="label-name"><span class="tag tag-pend"></span>Pendiente</span><span class="count" id="pendCount">' + stats.pend + '</span></div>' +
    '<div class="row"><span class="label-name"><span class="tag tag-aprob"></span>Gestionado</span><span class="count" id="aprobCount">' + stats.aprob + '</span></div>' +
    '<div class="row"><span class="label-name"><span class="tag tag-rech"></span>Rechazado</span><span class="count" id="rechCount">' + stats.rech + '</span></div></div>' +
    '<div class="footer" id="footer">Los datos se actualizan automáticamente</div>' +
    '<script>' +
    'function refreshStats(){google.script.run.withSuccessHandler(function(s){' +
    'document.getElementById("totalNum").textContent=s.total||0;' +
    'document.getElementById("pendCount").textContent=s.pend||0;' +
    'document.getElementById("aprobCount").textContent=s.aprob||0;' +
    'document.getElementById("rechCount").textContent=s.rech||0;' +
    'var d=new Date();var h=d.getHours().toString().padStart(2,"0");var m=d.getMinutes().toString().padStart(2,"0");var s2=d.getSeconds().toString().padStart(2,"0");' +
    'document.getElementById("updated").textContent="Actualizado: "+h+":"+m+":"+s2;' +
    '})._obtenerEstadisticas();}' +
    'setInterval(refreshStats,5000);' +
    '</script></body></html>'
  var output = HtmlService.createHtmlOutput(html).setTitle('Recepción Formulario').setWidth(300)
  SpreadsheetApp.getUi().showSidebar(output)
}

function _configurarTriggerFormulario() {
  var triggers = ScriptApp.getProjectTriggers()
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(triggers[i])
    }
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create()
}

function configurarTriggerFormulario() {
  try {
    _configurarTriggerFormulario()
    SpreadsheetApp.getUi().alert('Recepción automática activada. Los formularios se recibirán solos al ser enviados.')
  } catch(e) {
    SpreadsheetApp.getUi().alert('Error al configurar: ' + e.message + '\nAsegúrate de tener permisos y que el formulario este vinculado a este spreadsheet.')
  }
}

function crearFormularioBase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var legacy = ss.getSheetByName('Formulario Usuario / Profesional')
  if (legacy) legacy.setName(HOJA_FORM)
  var sh = ss.getSheetByName(HOJA_FORM)
  var existe = !!sh
  if (!existe) {
    ss.toast('Creando hoja de recepción…', 'PADDS', 1)
    sh = ss.insertSheet(HOJA_FORM)
  } else {
    ss.toast('Aplicando mejoras sin borrar datos…', 'PADDS', 1)
  }
  sh.setTabColor('#64748B')

  sh.setFrozenColumns(0)
  sh.setFrozenRows(0)

  var lc = 19, WHITE = '#ffffff'
  var SEC_COLORS = ['#475569', '#0F766E', '#15803D', '#7E22CE']
  var SEC_TINTS = ['#EEF2F7', '#E0F2FE', '#DCFCE7', '#F3E8FF']
  var BS = SpreadsheetApp.BorderStyle.SOLID
  var BS_MED = SpreadsheetApp.BorderStyle.SOLID_MEDIUM

  var FORM_TOOLTIPS = [
    'Fecha y hora de recepción del formulario (se llena automáticamente).',
    'Tipo de formulario recibido (ej: "Registro Atención Diaria").',
    'Estado de revisión: Pendiente → Gestionado (se envía a Pacientes) → Rechazado. Cambie el valor usando el menú.',
    'Correo del profesional que aprobó el formulario (se llena automáticamente).',
    'Fecha y hora de aprobación (se llena automáticamente).',
    'Fecha de la atención registrada. Se usa para actualizar la columna de servicio en Pacientes.',
    'RUN del usuario (sin puntos, con guion). Usado para buscar al paciente en Pacientes.',
    'Nombre del usuario.',
    'Apellido del usuario.',
    'Profesional que realizó la prestación.',
    'Prestación o servicio realizado. Determina a qué columna de Pacientes se copia la fecha.',
    'Clasificación del estado nutricional (solo aplica cuando la prestación es de Nutrióloga).',
    'Indica si la visita fue perdida (SI/NO). Si es SI, se registra en Observaciones.',
    'Resultado del índice Barthel (dependencia funcional).',
    'Resultado del test Zarit (sobrecarga del cuidador).',
    'Fecha sugerida para el próximo control.',
    'Observaciones de la atención. Se acumulan en Pacientes en formato [fecha] con los detalles.',
    'Paciente oncológico (SI/NO). Se copia a Pacientes (col. ONCOLÓGICO).',
    'Estado de la postulación a estipendio. Se copia a Pacientes (col. ESTIPENDIO).',
  ]

  sh.setRowHeight(4, 30)
  var HEADERS = [
    'FECHA RECEPCIÓN', 'FORMULARIO', 'ESTADO', 'APROBADO POR', 'FECHA APROB.',
    'FECHA ATENCIÓN', 'RUT USUARIO', 'NOMBRE', 'APELLIDO', 'PROFESIONAL',
    'PRESTACIÓN', 'ESTADO NUTRICIONAL', 'VISITA PERDIDA', 'RESULTADO BARTHEL',
    'RESULTADO ZARIT', 'PRÓXIMO CONTROL', 'OBSERVACIONES',
    'USUARIO ONCOLÓGICO', 'POSTULACIÓN ESTIPENDIO',
  ]
  for (var c = 0; c < lc; c++) {
    var secIdx = -1
    for (var si = 0; si < FORM_SECCIONES.length; si++) {
      if (c + 1 >= FORM_SECCIONES[si].ini && c + 1 <= FORM_SECCIONES[si].fin) { secIdx = si; break }
    }
    var hdrBg = secIdx >= 0 ? SEC_TINTS[secIdx] : '#EEF1F5'
    var hdrFg = secIdx >= 0 ? _ajustarHex(SEC_COLORS[secIdx], -40, -40, -40) : '#1E293B'
    sh.getRange(4, c + 1)
      .setValue(HEADERS[c])
      .setBackground(hdrBg).setFontColor(hdrFg)
      .setFontFamily(_UI.font).setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, WHITE, BS_MED)
      .setNote([
        HEADERS[c],
        FORM_TOOLTIPS[c] || ''
      ].join('\n'))
  }

  var NEW_ANCHOS = [150, 120, 100, 180, 140, 110, 100, 170, 150, 160, 180, 150, 90, 140, 130, 110, 280, 110, 150]
  for (var i = 0; i < Math.min(NEW_ANCHOS.length, lc); i++)
    sh.setColumnWidth(i + 1, NEW_ANCHOS[i])

  _miniResumenForm(sh, true)

  var lrForm = sh.getLastRow()
  var maxRows = Math.max(300, lrForm + 5)
  var statusCol = sh.getRange(5, 3, maxRows - 4, 1)
  var rules = []
  var estados = ['Rechazado', 'Gestionado', 'Pendiente']
  for (var ei = 0; ei < estados.length; ei++) {
    var css = _ESTADO_CSS[estados[ei]]
rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(estados[ei])
    .setBackground(css.badge).setFontColor(css.fg).setBold(true)
    .setRanges([statusCol]).build())
  }
  // El ESTADO se muestra solo como badge en la columna C (reglas de arriba):
  // pintar filas completas (naranja/verde/gris) chocaba con el formato de filas
  // y dejaba la hoja "toda naranja". Las filas de datos llevan cebra uniforme.

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($G5<>"",COUNTIF($G$5:$G,$G5)>1)')
    .setBackground('#FEF3C7').setFontColor('#C2410C').setBold(true)
    .setRanges([sh.getRange(5, 7, maxRows - 4, 1)]).build())
  sh.setConditionalFormatRules(rules)

  // Formato uniforme para las filas de datos existentes (cebra): reemplaza los
  // tintes por sección pintados por versiones antiguas, que se veían mezclados
  // con el formato actual. El color de estado ahora solo vive en la columna C.
  var _lrD = sh.getLastRow()
  if (_lrD >= 5) {
    var _nD = _lrD - 4
    var _rowsBg = []
    for (var _d = 0; _d < _nD; _d++) {
      var _one = []
      var _bCebra = (_d % 2 === 0) ? '#FFFFFF' : '#F8FAFC'
      for (var _cX = 0; _cX < lc; _cX++) _one.push(_bCebra)
      _rowsBg.push(_one)
    }
    sh.getRange(5, 1, _nD, lc).setBackgrounds(_rowsBg)
  }

  var hasta = existe ? Math.min(lrForm, 500) : 25
  for (var fr = 5; fr <= Math.max(hasta, 25); fr++) _estiloFila(sh, fr, lc)

  sh.getRange(5, 3, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(FORM_ESTADOS, true).setAllowInvalid(true).build())

  // API de Sheets en reglas de validación (bug conocido), y el RUN se normaliza y

  sh.getRange(4, 7).setNote('Formato RUN: 12345678-5 (sin puntos, con guión). Se normaliza y valida automáticamente al procesar la fila.')

  sh.getRange(5, 13, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['SI', 'NO', 'PADI SAN JUAN'], true)
      .setAllowInvalid(true).setHelpText('SI = visita perdida (se registra en Observaciones). PADI SAN JUAN = registro del centro, no se marca como perdida.').build())

  sh.getRange(5, 18, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['SI', 'NO'], true)
      .setAllowInvalid(true).setHelpText('Paciente oncológico (SI/NO). Se copia a Pacientes.').build())

  sh.getRange(5, 19, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['BENEFICIARIO', 'INGRESADO', 'PENDIENTE', 'NO APLICA', 'EN ESPERA', 'N/A'], true)
      .setAllowInvalid(true).setHelpText('Postulación a estipendio. Se copia a Pacientes.').build())

  sh.getRange(5, 16, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build())
  sh.getRange(5, 1, maxRows - 4, 1).setNumberFormat('dd/mm/yyyy hh:mm')
  sh.getRange(5, 5, maxRows - 4, 1).setNumberFormat('dd/mm/yyyy hh:mm')
  sh.getRange(5, 6, maxRows - 4, 1).setNumberFormat('dd/mm/yyyy')
  sh.getRange(5, 16, maxRows - 4, 1).setNumberFormat('dd/mm/yyyy')
  sh.getRange(5, 17, maxRows - 4, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
  sh.setRowHeights(5, maxRows - 4, 26)

  var _prestList = _PRESTACIONES_LIST.map(function(n) {
    return n.toLowerCase().replace(/\b\w/g, function(ch) { return ch.toUpperCase() })
  }).sort()
  sh.getRange(5, 11, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(_prestList, true)
      .setAllowInvalid(true).setHelpText('Elija la prestación realizada (se copia a la columna correspondiente de Pacientes)').build())

  sh.getRange(5, 12, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['NORMAL', 'SOBRE PESO', 'SOBREPESO', 'OBESIDAD', 'BAJO PESO', 'N/A'], true)
      .setAllowInvalid(true).build())

  sh.getRange(5, 10, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['MEDICO', 'NUTRICIONISTA', 'ENFERMERIA', 'TRABAJADOR SOCIAL',
        'KINESIOLOGIA', 'FONOAUDIOLOGIA', 'TENS', 'PODOLOGO', 'PSICOLOGA', 'OTRO'], true)
      .setAllowInvalid(true).build())

  sh.getRange(5, 14, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['DEPENDENCIA LEVE', 'DEPENDENCIA MODERADA', 'DEPENDENCIA SEVERA', 'N/A'], true)
      .setAllowInvalid(true).setHelpText('Opciones del formulario (el texto "Otros" también se acepta). Se normaliza al guardar en Pacientes.').build())

  sh.getRange(5, 15, maxRows - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['SOBRECARGA INTENSA', 'SOBRECARGA LEVE', 'SIN SOBRECARGA', 'AUSENCIA', 'N/A', 'PENDIENTE'], true)
      .setAllowInvalid(true).setHelpText('Resultado de la escala Zarit (se copia a Pacientes).').build())

  sh.setFrozenColumns(0)
  _unmergeQueCruzaFila(sh, 4)
  try { sh.setFrozenRows(4) } catch(eFz) {}
  var _frmF = sh.getFilter()
  if (_frmF) _frmF.remove()
  sh.getRange(4, 1, 1, lc).createFilter()

  try { ss.setSpreadsheetTimeZone('America/Santiago') } catch(e) {}

  try { _configurarTriggerFormulario() } catch(e) {}

  // Recortar solo filas completamente vacías al final (nunca se tocan datos)
  var _fmr = sh.getMaxRows()
  var _minimo = Math.max(lrForm + 10, maxRows + 5)
  if (_fmr > _minimo) {
    var _sobra = sh.getRange(_minimo + 1, 1, _fmr - _minimo, lc).getValues()
    var _todoVacio = true
    for (var _v = 0; _v < _sobra.length && _todoVacio; _v++) {
      for (var _vc = 0; _vc < lc; _vc++) {
        if (_sobra[_v][_vc] !== '' && _sobra[_v][_vc] !== null && _sobra[_v][_vc] !== undefined) { _todoVacio = false; break }
      }
    }
    if (_todoVacio) sh.deleteRows(_minimo + 1, _fmr - _minimo)
  }

  try { configurarResaltadoFila() } catch(e) {}

  try { sh.setFrozenRows(4) } catch(eFz) {}

  // Migración no destructiva: estados antiguos "Aprobado" pasan a "Gestionado"
  var _estData = sh.getRange(5, 3, Math.max(sh.getLastRow() - 4, 0), 1).getValues()
  var _huboCambio = false
  for (var _ec = 0; _ec < _estData.length; _ec++) {
    if (String(_estData[_ec][0] || '').trim() === 'Aprobado') { _estData[_ec][0] = 'Gestionado'; _huboCambio = true }
  }
  if (_huboCambio) sh.getRange(5, 3, _estData.length, 1).setValues(_estData)

  ss.toast(existe
    ? 'Hoja de recepción actualizada: formato, validaciones y resumen aplicados sin borrar ningún formulario'
    : 'Hoja de recepción creada con 19 columnas (recepción automática activada)', 'Recepción', 4)
}

function onFormSubmit(e) {
  try {
    if (!e || !e.namedValues) return
    var ss = e.source
    if (!ss) return
    var stagingSh = ss.getSheetByName(HOJA_FORM)
    if (!stagingSh) return

    var lc = 19
    var newRow = []
    for (var c = 0; c < lc; c++) newRow.push('')

    newRow[0] = e.namedValues['Marca temporal'] ? e.namedValues['Marca temporal'][0] : new Date()
    newRow[2] = 'Pendiente'
    newRow[1] = 'Registro Atención Diaria'

    var trimmedKeys = {}
    for (var _ef in e.namedValues) {
      if (e.namedValues.hasOwnProperty(_ef)) trimmedKeys[_ef.trim()] = e.namedValues[_ef]
    }

    for (var field in _FORM_FIELD_MAP) {
      if (_FORM_FIELD_MAP.hasOwnProperty(field) && trimmedKeys[field] && trimmedKeys[field][0]) {
        newRow[_FORM_FIELD_MAP[field] - 1] = String(trimmedKeys[field][0]).trim()
      }
    }

    // Nombres en MAYÚSCULAS (convención de Pacientes), aunque el formulario
    // los envíe en minúsculas.
    newRow[7] = _mayusNombre(newRow[7])
    newRow[8] = _mayusNombre(newRow[8])

    var rutRecibido = _normRUN(newRow[6])
    var rutRecibidoOK = /^\d{7,8}-[0-9K]$/.test(rutRecibido) && _validarDigitoRUT(rutRecibido)
    if (rutRecibido) newRow[6] = rutRecibido

    var lr = stagingSh.getLastRow()
    var nr = Math.max(lr + 1, 5)
    if (nr > stagingSh.getMaxRows()) stagingSh.insertRowsAfter(stagingSh.getMaxRows(), nr - stagingSh.getMaxRows() + 10)
    stagingSh.getRange(nr, 1, 1, lc).setValues([newRow])
    _estiloFila(stagingSh, nr, lc)
    _colorEstado(stagingSh, nr, 'Pendiente', lc)
    _miniResumenForm(stagingSh)
    if (rutRecibido && !rutRecibidoOK) {
      stagingSh.getRange(nr, 7).setBackground('#FEE2E2').setFontColor('#B91C1C').setFontWeight('bold')
      stagingSh.getRange(nr, 7).setNote('⚠️ RUN inválido: revisa el dígito verificador (formato 12345678-5)')
    }
    _log(ss, 'Recepción', 'onFormSubmit', 'ok', 'fila ' + nr)
  } catch(err) {
    console.log('onFormSubmit: ' + err.message)
    try { _log(ss, 'Recepción', 'onFormSubmit', 'ERROR', err.message) } catch (eLogF) {}
  }
}

function aprobarFormularios() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.toast('Procesando formularios…', 'PADDS', 1)
  var sh = ss.getSheetByName(HOJA_FORM)
  if (!sh) { SpreadsheetApp.getUi().alert('Primero crea la hoja de recepción (menú Formulario).'); return }

  var lr = sh.getLastRow()
  var lc = sh.getLastColumn()
  if (lr < 5) { ss.toast('No hay formularios para procesar', 'Formulario', 3); return }

  var lock = LockService.getScriptLock()
  try { if (!lock.tryLock(10000)) { ss.toast('No se pudo adquirir el bloqueo, intenta de nuevo', 'Formulario', 3); return } } catch(ee) { return }

  try {
    var data = sh.getRange(5, 1, lr - 4, lc).getValues()
    var email = ''
    try { email = Session.getActiveUser().getEmail() } catch(e) {}
    var now = new Date()

    var aprobRows = []

    var rechRows = []

    var avisosAp = []

    for (var r = data.length - 1; r >= 0; r--) {
      var rowNum = r + 5
      var estRaw = String(data[r][2] || '').trim()
      var estado = _normEstadoForm(estRaw)
      if (estado !== estRaw) {
        sh.getRange(rowNum, 3).setValue(estado)
        data[r][2] = estado
      }
      if (estado === 'Rechazado') {
        rechRows.push(rowNum)
      } else if (estado === 'Gestionado') {
        if (!data[r][6] || String(data[r][6]).trim() === '') {
          sh.getRange(rowNum, 7).setNote('⚠️ Falta el RUN: complétalo y vuelve a aprobar')
          sh.getRange(rowNum, 7).setBackground('#FEF3C7').setFontColor('#A16207').setFontWeight('bold')
          avisosAp.push('Fila ' + rowNum + ': falta el RUN (col. G).')
          continue
        }
        if (!data[r][7] || String(data[r][7]).trim() === '') {
          sh.getRange(rowNum, 8).setNote('⚠️ Falta el nombre: complétalo y vuelve a aprobar')
          sh.getRange(rowNum, 8).setBackground('#FEF3C7').setFontColor('#A16207').setFontWeight('bold')
          avisosAp.push('Fila ' + rowNum + ': falta el NOMBRE (col. H).')
          continue
        }
        var rutNorm = _normRUN(data[r][6])
        var rutOK = /^\d{7,8}-[0-9K]$/.test(rutNorm) && _validarDigitoRUT(rutNorm)
        if (!rutOK) {
          sh.getRange(rowNum, 7).setNote('⚠️ RUN inválido (formato 12345678-5 con dígito verificador correcto). Corrígelo y vuelve a aprobar.')
          sh.getRange(rowNum, 7).setBackground('#FEE2E2').setFontColor('#B91C1C').setFontWeight('bold')
          avisosAp.push('Fila ' + rowNum + ': RUN inválido (formato 12345678-5 con DV correcto).')
          continue
        }
        aprobRows.push({ row: rowNum, data: data[r] })
      }
    }

    // (las filas no se eliminan aún: primero se transfieren a Pacientes)
    for (var i = 0; i < aprobRows.length; i++) {
      var rn = aprobRows[i].row
      var cur = aprobRows[i].data
      if (!String(cur[3] || '').trim()) sh.getRange(rn, 4).setValue(email)
      if (!cur[4]) sh.getRange(rn, 5).setValue(now).setNumberFormat('dd/mm/yyyy hh:mm')
    }

    var transferidos = aprobRows.length ? _batchCopiarFormularios(aprobRows) : { rows: [], filaDestino: 0 }
    if (transferidos.rows.length) recalcularTodo()

    // Eliminar SOLO filas (nunca hojas): Aprobados (ya copiados a Pacientes) y Rechazados

    var aEliminar = rechRows.slice()
    for (var i = 0; i < transferidos.rows.length; i++) aEliminar.push(transferidos.rows[i])
    aEliminar.sort(function(a, b) { return b - a })
    var eliminados = 0
    for (var i = 0; i < aEliminar.length; i++) {
      try { sh.deleteRow(aEliminar[i]); eliminados++ } catch (eDel) {}
    }

    var msg = []
    if (transferidos.rows.length) msg.push(transferidos.rows.length + ' transferidos a Pacientes')
    if (rechRows.length) msg.push(rechRows.length + ' rechazados eliminados')
    ss.toast(msg.length ? msg.join(', ') : 'No hay Gestionados ni Rechazados para procesar. Pendientes se saltan.', 'PADDS', 4)

    // Ir al paciente transferido para que se vea dónde quedó.
    try {
      if (transferidos.filaDestino >= 4) {
        var pacN = ss.getSheetByName(HOJA_PAC)
        if (pacN && transferidos.filaDestino <= pacN.getLastRow()) {
          pacN.setActiveRange(pacN.getRange(transferidos.filaDestino, 1, 1, pacN.getLastColumn()))
        }
      }
    } catch (eNav) {}
    if (avisosAp.length) {
      SpreadsheetApp.getUi().alert('Formularios con datos incompletos',
        'Estas filas están marcadas como Gestionado pero faltan datos o son inválidos, por lo que NO se enviaron a Pacientes:\n\n' +
        avisosAp.join('\n') + '\n\nCorrige los datos y vuelve a aprobar.',
        SpreadsheetApp.getUi().ButtonSet.OK)
    }
    _miniResumenForm(sh)
  } finally {
    lock.releaseLock()
  }
}

function _batchCopiarFormularios(formRows) {
  if (!formRows.length) return []
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var pac = ss.getSheetByName(HOJA_PAC)
  if (!pac) return []

  try { _borrarFilasVacias(pac, 4) } catch (eB) {}
  var lr = pac.getLastRow()
  var lc = pac.getLastColumn()

  var runMap = {}
  var pacData = []
  var maxId = 0
  if (lr >= 4) {
    pacData = pac.getRange(4, 1, lr - 3, lc).getValues()
    for (var i = 0; i < pacData.length; i++) {
      var rut = _normRUN(pacData[i][COL.RUN - 1])
      if (rut) {
        runMap[rut] = { row: i + 4, data: pacData[i] }
        if (rut.indexOf('-') >= 0) runMap[rut.replace('-', '')] = runMap[rut]
      }
      var id = Number(pacData[i][COL.ID - 1])
      if (id > maxId) maxId = id
    }
  }

  var newRows = []
  var updated = 0
  var hechos = []
  var filaDestino = 0

  for (var r = 0; r < formRows.length; r++) {
    var formData = formRows[r].data
    var formRut = _normRUN(formData[6])
    if (!formRut) continue

    var found = runMap[formRut] || runMap[formRut.replace('-', '')]

    var fechaAtencion = formData[5]
    var fechaOk = fechaAtencion != null && _parseDate(fechaAtencion) != null
    var prestacion = String(formData[10] || '').trim()
    var prestNorm = prestacion.toUpperCase()
      .replace(/[ÁÀÄÂ]/g,'A').replace(/[ÉÈËÊ]/g,'E').replace(/[ÍÌÏÎ]/g,'I')
      .replace(/[ÓÒÖÔ]/g,'O').replace(/[ÚÙÜÛ]/g,'U')
    var servicioCol = _SERVICIO_COL_MAP[prestNorm] || _SERVICIO_SI_MAP[prestNorm]
    if (!servicioCol && prestNorm.indexOf('SONDA') >= 0) servicioCol = COL.F_CAMBIO_SONDA
    var observaciones = String(formData[16] || '').trim()
    var profesional = String(formData[9] || '').trim()
    var visitaPerdida = String(formData[12] || '').trim()
    var proxControlRaw = formData[15]
    var proxControl = (proxControlRaw instanceof Date && !isNaN(proxControlRaw.getTime()))
      ? proxControlRaw : _parseDate(proxControlRaw)

    var fechaStr = (fechaAtencion instanceof Date && !isNaN(fechaAtencion.getTime()))
      ? fmtFecha(fechaAtencion)
      : fmtFecha(new Date())

    if (found) {
      // ─── Existing patient: update service date and cumulative logs ────────
      // Todas las actualizaciones se acumulan en memoria y se escriben con un
      // solo setValues (antes era ~8-10 escrituras por fila).
      var pacRow = found.row
      var rowArr = found.data.slice()
      var tocObs = false, tocMc = false

      if (servicioCol && servicioCol <= lc) {
        var sv = _SERVICIO_SI_MAP[prestNorm] ? 'SI' : (fechaOk ? fechaAtencion : '')
        if (sv) rowArr[servicioCol - 1] = sv
      }

      var obsParts = []
      if (prestacion) obsParts.push('Prestación: ' + prestacion)
      if (observaciones) obsParts.push('Obs: ' + observaciones)
      if (visitaPerdida === 'SI') obsParts.push('Visita perdida')
      if (proxControl) obsParts.push('Próximo control: ' + fmtFecha(proxControl))
      if (prestNorm.indexOf('CONTROL CUIDADOR') >= 0) obsParts.push('Control cuidador')
      if (formData.length > 17 && String(formData[17] || '').trim().toUpperCase() === 'SI') obsParts.push('Paciente oncológico')
      if (prestacion && !fechaOk) obsParts.push('⚠️ Fecha de atención inválida')
      if (obsParts.length) {
        var obsLine = '[' + fechaStr + '] ' + obsParts.join('\n    ')
        var obsOld = String(rowArr[COL.OBSERVACIONES - 1] || '').trim()
        var obsFlat = obsOld.replace(/\s+/g, ' ').trim()
        var lineFlat = obsLine.replace(/\s+/g, ' ').trim()
        if (!obsFlat || obsFlat.indexOf(lineFlat) === -1) {
          rowArr[COL.OBSERVACIONES - 1] = obsOld ? obsOld + '\n\n' + obsLine : obsLine
          tocObs = true
        }
      }

      if (prestacion) {
        var mcOld = String(rowArr[COL.CONTROLES_MISCELANEOS - 1] || '').trim()
        var mcLine = '[' + fechaStr + '] ' + prestacion
        var mcFlat = mcOld.replace(/\s+/g, ' ').trim()
        var mcLineFlat = mcLine.replace(/\s+/g, ' ').trim()
        if (!mcFlat || mcFlat.indexOf(mcLineFlat) === -1) {
          rowArr[COL.CONTROLES_MISCELANEOS - 1] = mcOld ? mcOld + '\n' + mcLine : mcLine
          tocMc = true
        }
      }

      if (profesional) {
        rowArr[COL.EDITOR - 1] = profesional
      }

      // Nombre/apellidos normalizados desde el formulario (sin editar a mano)
      // y sector PENDIENTE si el paciente aún no tiene uno asignado.
      var _nU = _dividirNombreApellidos(String(formData[7] || ''), String(formData[8] || ''), '')
      if (_nU.nombre) rowArr[COL.NOMBRE - 1] = _nU.nombre
      if (_nU.apellido) rowArr[COL.APELLIDO - 1] = _nU.apellido
      if (_nU.apellido2) rowArr[COL.APELLIDO2 - 1] = _nU.apellido2
      if (String(rowArr[COL.SECTOR - 1] || '').trim() === '') {
        rowArr[COL.SECTOR - 1] = 'PENDIENTE'
      }

      for (var fi = 0; fi < FORM_A_PAC.length; fi++) {
        var fpCol = FORM_A_PAC[fi][0] - 1, ppCol = FORM_A_PAC[fi][1] - 1
        if (fpCol < formData.length && ppCol >= 0 && ppCol < lc
            && fpCol !== 5 && ppCol !== COL.OBSERVACIONES - 1) {
          var fv = formData[fpCol]
          if (fv !== undefined && fv !== null && String(fv).trim() !== '') {
            var fvs = String(fv).trim()
            fvs = _formValorNormalizado(ppCol + 1, fvs)
            if (ppCol === COL.NOMBRE - 1 || ppCol === COL.APELLIDO - 1) fvs = _mayusNombre(fvs)
            rowArr[ppCol] = fvs
          }
        }
      }

      pac.getRange(pacRow, 1, 1, lc).setValues([rowArr])
      if (tocObs) pac.getRange(pacRow, COL.OBSERVACIONES, 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      if (tocMc) pac.getRange(pacRow, COL.CONTROLES_MISCELANEOS, 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)

      updated++
      filaDestino = pacRow
    } else {
      // ─── New patient: create row ────────────────────────────────────────
      maxId++
      var blank = []
      for (var c = 0; c < lc; c++) {
        blank.push(c === COL.ID - 1 ? maxId : c === COL.VITAL - 1 ? 'VIGENTE' : '')
      }

      blank[COL.RUN - 1] = formRut

      for (var i = 0; i < FORM_A_PAC.length; i++) {
        var fCol = FORM_A_PAC[i][0] - 1, pCol = FORM_A_PAC[i][1] - 1
        if (fCol < formData.length && pCol >= 0 && pCol < lc) {
          var val = formData[fCol]
          if (val !== undefined && val !== null && String(val).trim() !== '') {
            var vs = String(val).trim()
            vs = _formValorNormalizado(pCol + 1, vs)
            if (pCol === COL.NOMBRE - 1 || pCol === COL.APELLIDO - 1) vs = _mayusNombre(vs)
            blank[pCol] = vs
          }
        }
      }

      if (servicioCol && servicioCol <= lc && fechaOk) {
        blank[servicioCol - 1] = _SERVICIO_SI_MAP[prestNorm] ? 'SI' : fechaAtencion
      }

      // Sector nuevo queda PENDIENTE (gris) hasta que se asigne.
      if (!blank[COL.SECTOR - 1]) blank[COL.SECTOR - 1] = 'PENDIENTE'
      // Separar el nombre en NOMBRE/APELLIDO/APELLIDO 2 para no editarlo a mano.
      var _nS = _dividirNombreApellidos(String(formData[7] || ''), String(formData[8] || ''), '')
      if (_nS.nombre) blank[COL.NOMBRE - 1] = _nS.nombre
      if (_nS.apellido) blank[COL.APELLIDO - 1] = _nS.apellido
      if (_nS.apellido2) blank[COL.APELLIDO2 - 1] = _nS.apellido2

      var obsParts2 = []
      if (prestacion) obsParts2.push('Prestación: ' + prestacion)
      if (observaciones) obsParts2.push('Obs: ' + observaciones)
      if (visitaPerdida === 'SI') obsParts2.push('Visita perdida')
      if (proxControl) obsParts2.push('Próximo control: ' + fmtFecha(proxControl))
      if (formData.length > 17 && String(formData[17] || '').trim().toUpperCase() === 'SI') obsParts2.push('Paciente oncológico')
      if (prestacion && !fechaOk) obsParts2.push('⚠️ Fecha de atención inválida')
      if (obsParts2.length) blank[COL.OBSERVACIONES - 1] = '[' + fechaStr + '] ' + obsParts2.join(' | ')

      if (prestacion) blank[COL.CONTROLES_MISCELANEOS - 1] = '[' + fechaStr + '] ' + prestacion

      newRows.push(blank)
    }
    hechos.push(formRows[r].row)
  }

  var nr = 0, insertadas = 0
  if (newRows.length > 0) {
    nr = lr + 1
    insertadas = 0
    if (nr + newRows.length - 1 > pac.getMaxRows()) {
      insertadas = nr + newRows.length - pac.getMaxRows()
      pac.insertRowsAfter(pac.getMaxRows(), insertadas)
    }
    pac.getRange(nr, 1, newRows.length, lc).setValues(newRows)
    filaDestino = nr
    var rng = pac.getRange(nr, 1, newRows.length, lc)
    rng.setFontColor('#000000')
    rng.setFontWeight('normal')
    rng.setFontStyle('normal')
    pac.getRange(nr, COL.OBSERVACIONES, newRows.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    pac.getRange(nr, COL.CONTROLES_MISCELANEOS, newRows.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    pac.setRowHeights(nr, newRows.length, 26)
  }

  try {
    var _nPre2 = _compactarPacientes(pac) + _limpiarFilasVaciasLoop(pac, 4)
    _log(ss, 'Pacientes', '_batchCopiarFormularios', 'ok',
      'nr=' + nr + ' lr=' + lr + ' nuevos=' + newRows.length + ' insertadas=' + insertadas +
      ' maxRows=' + pac.getMaxRows() + ' vaciasEliminadas=' + _nPre2)
  } catch (eV2) {}

  ss.toast(updated + ' pacientes actualizados, ' + newRows.length + ' pacientes creados desde formularios', 'PADDS', 4)
  return { rows: hechos, filaDestino: filaDestino }
}

function rechazarFormularios() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    ss.toast('Rechazando formularios…', 'PADDS', 1)
    var sh = ss.getSheetByName(HOJA_FORM)
    if (!sh) return

    var lr = sh.getLastRow()
    if (lr < 5) { ss.toast('No hay formularios para rechazar', 'Formulario', 3); return }

    var data = sh.getRange(5, 1, lr - 4, sh.getLastColumn()).getValues()
    var email = ''
    try { email = Session.getActiveUser().getEmail() } catch(e) { email = 'usuario' }

    var rechRows = []
    for (var r = 0; r < data.length; r++) {
      if (_normEstadoForm(data[r][2]) !== 'Pendiente') continue
      rechRows.push(r + 5)
    }

    if (rechRows.length === 0) { ss.toast('No hay Pendientes para rechazar', 'Formulario', 3); return }

    var _quien = email + ' · ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    // Los Rechazados NO son contiguos: no se escribe un bloque único (sobrescribiría
    // el ESTADO de filas intermedias). Se agrupan los tramos contiguos y se escribe
    // por tramo en lote (estado, nota, badges).
    var _rechG = _agruparContiguos(rechRows.slice())
    var _rcss = _ESTADO_CSS['Rechazado']
    for (var i = 0; i < _rechG.length; i++) {
      var r0 = _rechG[i][0], rn = _rechG[i][1]
      var rgC = sh.getRange(r0, 3, rn, 1)
      var vals = [], notas = [], bgs = [], fgs = [], wts = []
      for (var v = 0; v < rn; v++) {
        vals.push(['Rechazado'])
        notas.push(['Rechazado por: ' + _quien])
        bgs.push([_rcss.badge])
        fgs.push([_rcss.fg])
        wts.push(['bold'])
      }
      rgC.setValues(vals).setNotes(notas).setBackgrounds(bgs).setFontColors(fgs).setFontWeights(wts)
    }

    ss.toast(rechRows.length + ' formularios marcados como Rechazado (usa "🧹 Limpiar procesados" para eliminar)', 'Formulario', 4)
    _miniResumenForm(sh)
  } catch(e) {
    SpreadsheetApp.getUi().alert('Error al rechazar: ' + e.message)
  }
}

function limpiarFormulariosAprobados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(HOJA_FORM)
  if (!sh) return
  var lr = sh.getLastRow()
  var lc = sh.getLastColumn()
  if (lr < 5) return

  ss.toast('Buscando formularios procesados…', 'PADDS', 1)
  var data = sh.getRange(5, 1, lr - 4, lc).getValues()
  var ui = SpreadsheetApp.getUi()
  var aLimpiar = 0

  for (var r = data.length - 1; r >= 0; r--) {
    var estLim = _normEstadoForm(data[r][2])
    if (estLim === 'Gestionado' || estLim === 'Rechazado') {
      aLimpiar++
    }
  }

  if (aLimpiar === 0) { ss.toast('No hay formularios gestionados o rechazados para limpiar', 'Formulario', 3); return }
  if (ui.alert('Limpiar', '¿Eliminar ' + aLimpiar + ' formularios gestionados o rechazados?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return

  var grupos = [], start = -1
  for (var r = data.length - 1; r >= 0; r--) {
    var estGr = _normEstadoForm(data[r][2])
    if (estGr === 'Gestionado' || estGr === 'Rechazado') {
      if (start < 0) { start = r + 5; grupos.push([start, 1]) }
      else { grupos[grupos.length - 1][1]++ }
    } else { start = -1 }
  }
  for (var g = 0; g < grupos.length; g++) {
    sh.deleteRows(grupos[g][0], grupos[g][1])
  }

  _miniResumenForm(sh)
  SpreadsheetApp.getActiveSpreadsheet().toast(aLimpiar + ' formularios eliminados de la hoja', 'PADDS', 3)
}
