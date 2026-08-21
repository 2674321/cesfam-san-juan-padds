function onOpen() {
  var ui = SpreadsheetApp.getUi()

  ui.createMenu('📅 Agenda')
    .addItem('➕ Agregar semana (fecha actual)', 'generarSemana')
    .addItem('➕ Agregar semana en fecha específica', 'generarSemanaEspecifica')
    .addItem('➕ Agregar 3 semanas seguidas', 'generarTresSemanas')
    .addSeparator()
    .addItem('❌ Eliminar semana', 'eliminarSemana')
    .addItem('🧹 Limpiar semanas pasadas', 'limpiarSemanasPasadas')
    .addSeparator()
    .addItem('📅 Ir al día de hoy', 'irAHoy')
    .addItem('📅 Ir a una fecha', 'irAFecha')
    .addItem('🔎 Buscar paciente', 'buscarPaciente')
    .addSeparator()
    .addItem('📊 Resumen de atenciones', 'resumen')
    .addToUi()

  ui.createMenu('🩺 Pacientes')
    .addItem('🔎 Buscar paciente', 'buscarEnPacientes')
    .addItem('➕ Agregar paciente', 'agregarPaciente')
    .addItem('➖ Eliminar paciente', 'eliminarPaciente')
    .addSeparator()
    .addItem('🔤 Ordenar A-Z', 'ordenarPacientes')
    .addItem('🔢 Renumerar IDs', 'reindexarPacientes')
    .addSeparator()
    .addItem('🧹 Limpiar datos (espacios, mayúsculas, RUN, teléfonos)', 'limpiarPacientesCompleto')
    .addSeparator()
    .addItem('🔄 Recalcular estados', 'recalcularTodo')
    .addSeparator()
    .addItem('📄 Ficha Resumen', 'verFichaPaciente')
    .addToUi()

  ui.createMenu('📥 Ingresos')
    .addItem('📋 Crear hoja INGRESOS', 'crearHojaIngresos')
    .addItem('🧹 Corregir y mejorar formato', 'corregirYMejorarIngresos')
    .addItem('🔤 Ordenar filas por fecha', 'ordenarIngresosPorFecha')
    .addSeparator()
    .addItem('⚙️ Activar confirmación de INGRESA', 'configurarTriggerIngresos')
    .addItem('📤 Enviar filas INGRESA a Pacientes', 'enviarIngresasAPacientes')
    .addToUi()

  ui.createMenu('📝 Formulario')
    .addItem('📋 Crear hoja de recepción', 'crearFormularioBase')
    .addItem('📊 Estadísticas', 'mostrarSidebarEstadisticas')
    .addSeparator()
    .addItem('✅ Sincronizar Aprobados a Pacientes', 'aprobarFormularios')
    .addItem('❌ Rechazar pendientes', 'rechazarFormularios')
    .addItem('🧹 Limpiar procesados', 'limpiarFormulariosAprobados')
    .addSeparator()
    .addItem('⚙️ Activar recepción automática', 'configurarTriggerFormulario')
    .addToUi()

  ui.createMenu('🔧 Herramientas')
    .addItem('🎨 Ordenar y colorear hojas', 'organizarHojas')
    .addItem('🔄 Reconstruir Pacientes', 'repintarPacientes')
    .addItem('🎨 Formatear Pacientes (visual)', 'formatearPacientesVisual')
    .addSeparator()
    .addItem('📋 Parámetros de vigencia', 'crearParametros')
    .addItem('💊 Crear columnas de recetas controladas', 'crearColumnasRecetas')
    .addItem('📊 Dashboard', 'crearDashboard')
    .addItem('🩹 Reparar fórmulas del Dashboard', 'repararFormulas')
    .addSeparator()
    .addItem('🧪 Probar aviso de vencimiento', 'probarDiasAviso')
    .addItem('🎯 Activar resaltado de fila activa', 'configurarResaltadoFila')
    .addSeparator()
    .addItem('🧪 Ejecutar pruebas', 'ejecutarPruebas')
    .addToUi()

  ui.createMenu('🚨 Alertas')
    .addItem('📊 Ver alertas', 'mostrarAlertas')
    .addToUi()

  ui.createMenu('ℹ️ Ayuda')
    .addItem('❓ Ayuda rápida', 'mostrarAyudaRapida')
    .addItem('📋 Referencia de columnas', 'crearInstrucciones')
    .addSeparator()
    .addItem('ℹ️ Acerca del sistema', 'acercaDelSistema')
    .addToUi()
}

function mostrarAyudaRapida() {
  var html = '<html><head><base target="_top"><style>' +
    'body{font-family:Arial,sans-serif;margin:0;padding:14px;color:#333;font-size:12px;line-height:1.5}' +
    'h2{font-size:16px;color:#1a237e;margin:0 0 4px}' +
    'h3{font-size:13px;color:#fff;margin:10px 0 4px;padding:5px 8px;border-radius:3px}' +
    'h3.a{background:#1a3c5e} h3.p{background:#2e7d5b} h3.f{background:#1f6fb2} h3.h{background:#5c4ee5} h3.al{background:#e65100}' +
    'p{margin:2px 0}' +
    'li{margin:1px 0 1px 14px}' +
    'b{color:#222}' +
    '</style></head><body>' +
    '<h2>Ayuda rápida</h2>' +
    '<h3 class="a">📅 Agenda</h3>' +
    '<b>¿Qué hace?</b> Organiza visitas domiciliarias por semana. Cada profesional tiene su columna.<br>' +
    '<b>Cómo empezar:</b> Menú Agenda → "Agregar semana de esta fecha".<br>' +
    '<b>Datos:</b> Cabecera azul → nombre del profesional. Horario fijo 8-16. Atención: VDI/abreviadas/REGISTRO.<br>' +
    '<b>Limpiar:</b> ✕/📦 → selecciona y Delete para limpiar la semana. O usa "Limpiar semanas anteriores a hoy".<br>' +
    '<b>Buscar:</b> "Buscar paciente en la agenda", "Ir al día de hoy" o "Ir a una fecha".' +
    '<h3 class="p">🩺 Pacientes</h3>' +
    '<b>¿Qué hace?</b> Registro completo de pacientes: datos personales, patologías, controles, cuidador (107 campos).<br>' +
    '<b>Cómo usar:</b> Buscar por RUN o apellido. Agregar nuevo (ID automático). Ordenar A-Z (renumerar después).<br>' +
    '<b>Formatear:</b> Teléfonos → +56 y espacios. RUN → sin puntos. Mayúsculas → nombres en MAYÚSCULAS.<br>' +
    '<b>Estados:</b> Cada control se calcula solo AL DÍA / POR VENCER / VENCIDO. Usa "Recalcular estados" si editaste fechas manualmente. Plazos en Parámetros.' +
    '<h3 class="h">🔧 Herramientas</h3>' +
    '<b>Reconstruir Pacientes:</b> restaura colores, validaciones y tooltips sin perder datos.<br>' +
    '<b>Parámetros:</b> configura plazos (meses) de cada control y días de aviso para "POR VENCER".<br>' +
    '<b>Dashboard:</b> crea hoja con gráficos y estadísticas en vivo.<br>' +
    '<b>Reparar fórmulas:</b> corrige #REF! en Dashboard.<br>' +
    '<b>Verificar aviso:</b> muestra qué pacientes están por vencer.' +
    '<h3 class="f">📝 Formulario</h3>' +
    '1. Menú → "Crear hoja de recepción"<br>' +
    '2. Los formularios llegan solos (recepción automática)<br>' +
    '3. Revisa los datos<br>' +
    '4. "Aprobar pendientes" → pasa a Pacientes<br>' +
    '5. "Rechazar pendientes" · 6. "Limpiar procesados"' +
    '<h3 class="al">🚨 Alertas</h3>' +
    'Muestra resumen visual de pacientes con controles VENCIDOS (rojo), POR VENCER (naranja), PENDIENTES (amarillo) y prioritarios. Haz clic para ir a la fila.' +
    '</body></html>'

  var output = HtmlService.createHtmlOutput(html).setTitle('Ayuda Rápida - PADDS').setWidth(400)
  SpreadsheetApp.getUi().showSidebar(output)
}

function acercaDelSistema() {
  var html = '<html><head><base target="_top"><style>' +
    'body{font-family:Arial,sans-serif;margin:0;padding:14px;color:#333;font-size:13px;line-height:1.6}' +
    'h2{font-size:18px;color:#1a237e;margin:0 0 2px}' +
    'h3{font-size:13px;color:#1a237e;margin:12px 0 4px;border-bottom:1px solid #e0e0e0}' +
    '.ver{font-size:11px;color:#888;margin-bottom:8px}' +
    '.mod{margin:4px 0;padding:4px 8px;border-left:3px solid #ccc;background:#f8f9fa}' +
    '</style></head><body>' +
    '<h2>Sistema PADDS</h2>' +
    '<div class="ver">Planilla de Atención Domiciliaria para CESFAM · v3.0</div>' +
    '<h3>🟦 Pacientes</h3><div class="mod">Ficha única con datos personales, patologías, cuidador, controles y priorización automática. Más de 100 campos organizados por secciones de colores.</div>' +
    '<h3>🟧 Agenda</h3><div class="mod">Planificación semanal de visitas. Múltiples profesionales en columnas paralelas. Slots de atención con colores por tipo.</div>' +
    '<h3>🟩 Formulario</h3><div class="mod">Recepción de solicitudes desde formularios digitales. Aprobación o rechazo en lote con pase automático a Pacientes.</div>' +
    '<h3>🟨 Alertas</h3><div class="mod">Panel visual con controles vencidos, por vencer, pendientes y pacientes prioritarios. Navegación con un clic.</div>' +
    '<h3>🟪 Dashboard</h3><div class="mod">Estadísticas en vivo: totales, demografía, patologías, vigencias por área, alertas resumen.</div>' +
    '<div style="margin-top:16px;font-size:11px;color:#aaa;text-align:center">© 2026 Patricio A. Varela C. · Interno TENS</div>' +
    '</body></html>'

  var output = HtmlService.createHtmlOutput(html).setTitle('Acerca del Sistema').setWidth(380)
  SpreadsheetApp.getUi().showSidebar(output)
}

// ─── ORDENAR Y COLOREAR HOJAS ────────────────────────────────────────────────
// Ordena las pestañas por flujo de trabajo y les asigna color según su
// función. NO crea ni elimina hojas: solo reordena y colorea las que ya
// existen. Las hojas desconocidas quedan donde estén, sin color.

function organizarHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  // [nombre, color de pestaña] — en el orden deseado (flujo de trabajo)
  var orden = [
    ['INGRESOS',                     '#2E7D5B'],  // entrada / admisión
    ['Pacientes',                    '#1F6FB2'],  // base de datos central
    ['Agenda Profesionales',         '#F9A825'],  // agendamiento semanal
    ['Formulario Usuario / Profesional', '#6A1B9A'], // recepción de solicitudes
    ['LISTA DE TRABAJO',             '#C62828'],  // operativo diario
    ['DESVIO MEDICO',                '#E65100'],  // desvíos
    ['AGENDA REUNIONES',             '#00838F'],  // reuniones
    ['KINESIOLOGIA EVALUACION',      '#5C4EE5'],  // evaluaciones kine
    ['Dashboard',                    '#2E7D32'],  // estadísticas
    ['Referencia Columnas',          '#8D6E63'],  // guía de columnas
    ['Parametros',                   '#546E7A'],  // configuración
    ['Config',                       '#78909C'],  // configuración interna
    ['_PlantillaSemana',             '#9E9E9E'],  // plantilla interna
    ['_Resalte',                     '#BDBDBD'],  // soporte interno
  ]

  var activa = ''
  try { activa = ss.getActiveSheet().getName() } catch (e) {}

  // 1) Colorear las que existen (no se toca nada más)
  var presentes = 0
  for (var i = 0; i < orden.length; i++) {
    var sh = ss.getSheetByName(orden[i][0])
    if (sh) {
      sh.setTabColor(orden[i][1])
      presentes++
    }
  }

  // 2) Reordenar solo las visibles que existen; las ocultas se dejan quietas
  var pos = 1
  for (var i = 0; i < orden.length; i++) {
    var sh2 = ss.getSheetByName(orden[i][0])
    if (!sh2) continue
    if (sh2.isSheetHidden()) continue
    try {
      ss.setActiveSheet(sh2, false)
      ss.moveActiveSheet(pos)
      pos++
    } catch (eMov) {
      console.error('organizarHojas: ' + orden[i][0] + ' -> ' + eMov.message)
    }
  }

  // 3) Volver a la hoja que estaba activa
  try {
    var back = ss.getSheetByName(activa)
    if (back) ss.setActiveSheet(back, false)
  } catch (e) {}

  ss.toast('Hojas ordenadas y coloreadas (' + presentes + ' hojas)', '', 5)
}

// ─── CREAR COLUMNAS DE RECETAS CONTROLADAS ──────────────────────────────────
// Agrega al final de Pacientes: col 112 = F. RECETA (fecha) y col 113 =
// V. RECETA (vigencia automática). NO reestructura columnas existentes.

function crearColumnasRecetas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) { ss.toast('Hoja Pacientes no encontrada', '', 3); return }
  var ui = SpreadsheetApp.getUi()

  var maxC = sh.getMaxColumns()
  var h112 = maxC >= 112 ? String(sh.getRange(3, 112).getValue() || '').trim() : ''
  var h113 = maxC >= 113 ? String(sh.getRange(3, 113).getValue() || '').trim() : ''
  var ok112 = h112.toUpperCase().indexOf('RECETA') >= 0
  var ok113 = h113.toUpperCase().indexOf('RECETA') >= 0

  if (ok112 && ok113) {
    ui.alert('Recetas controladas', 'Las columnas ya existen:\n\nCol 112: ' + h112 + '\nCol 113: ' + h113 + '\n\nNo se modifica nada.')
    return
  }
  if (h112 && !ok112) {
    ui.alert('Atención', 'La columna 112 ya existe con otro encabezado: "' + h112 + '".\nNo se modifica para no perder datos.')
    return
  }
  if (h113 && !ok113) {
    ui.alert('Atención', 'La columna 113 ya existe con otro encabezado: "' + h113 + '".\nNo se modifica para no perder datos.')
    return
  }

  var r = ui.alert('Recetas controladas',
    'Se agregarán al final de Pacientes:\n\n' +
    'Col 112: F. RECETA — fecha de la última receta\n' +
    'Col 113: V. RECETA — vigencia automática\n\n' +
    '¿Continuar?',
    ui.ButtonSet.YES_NO)
  if (r !== ui.Button.YES) return

  if (maxC < 112) sh.insertColumns(112, Math.min(2, 113 - maxC))
  else if (maxC < 113) sh.insertColumns(113, 1)

  sh.getRange(3, 112).setValue('F. RECETA')
  sh.getRange(3, 113).setValue('V. RECETA')

  var lr = sh.getLastRow()
  if (lr >= 4) {
    var filas = lr - 3
    sh.getRange(4, 112, filas, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build())
      .setNumberFormat('dd/mm/yyyy')
    sh.getRange(4, 113, filas, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(['V. RECETA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'], true)
        .setAllowInvalid(true).build())
  }

  sh.setColumnWidth(112, 110)
  sh.setColumnWidth(113, 110)
  sh.getRange(3, 112, 1, 2).setFontWeight('bold')
    .setFontColor('#ffffff').setBackground('#4A148C')

  recalcularTodo()
  ss.toast('Columnas de recetas controladas creadas', '', 4)
}

