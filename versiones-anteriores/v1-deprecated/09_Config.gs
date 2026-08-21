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
    .addItem('🔄 Reconstruir Pacientes', 'repintarPacientes')
    .addItem('🎨 Formatear Pacientes (visual)', 'formatearPacientesVisual')
    .addSeparator()
    .addItem('📋 Parámetros de vigencia', 'crearParametros')
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

