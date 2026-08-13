// ════════════════════════════════════════════════════════
// ÍNDICE ▏09_Config.gs │ UI: menú onOpen, rehacer menús, ayuda y organización de hojas

// ─────────────────────────────────────────────────────────
// ════════════════════════════════════════════════════════
function onOpen() {
  var ui = SpreadsheetApp.getUi()

  ui.createMenu('📅 Agenda')
    .addItem('➕ Agregar semana', 'generarSemana')
    .addSubMenu(ui.createMenu('➕ Más semanas')
      .addItem('📅 En otra fecha', 'generarSemanaEspecifica')
      .addItem('🗓️ 3 semanas seguidas', 'generarTresSemanas'))
    .addSeparator()
    .addItem('❌ Eliminar semana', 'eliminarSemana')
    .addItem('🧹 Borrar semanas pasadas', 'limpiarSemanasPasadas')
    .addSeparator()
    .addItem('🎯 Ir a hoy', 'irAHoy')
    .addItem('📅 Ir a fecha', 'irAFecha')
    .addItem('🔎 Buscar paciente', 'buscarPaciente')
    .addSeparator()
    .addItem('📊 Resumen', 'resumen')
    .addToUi()

  ui.createMenu('🩺 Pacientes')
    .addItem('➕ Nuevo paciente', 'agregarPaciente')
    .addItem('🔎 Buscar paciente', 'buscarEnPacientes')
    .addItem('📄 Ficha resumen', 'verFichaPaciente')
    .addSeparator()
    .addSubMenu(ui.createMenu('🛠️ Datos')
      .addItem('✨ Corregir datos', 'limpiarPacientesCompleto')
      .addSeparator()
      .addItem('🪪 Verificar RUN', 'verificarIntegridadRUN')
      .addItem('🪪 Formatear RUN', 'formatearRUTPacientes')
      .addSeparator()
      .addItem('🔡 Ordenar (A-Z)', 'ordenarPacientes')
      .addItem('🔢 Renumerar filas', 'reindexarPacientes')
      .addItem('🧹 Depurar duplicados', 'depurarDuplicados')
      .addSeparator()
      .addItem('🎨 Formatear hoja', 'formatearDatos')
      .addItem('🛡️ Reparar formato', 'repararFormatoPacientes')
      .addSeparator()
      .addItem('➖ Eliminar paciente', 'eliminarPaciente'))
    .addSubMenu(ui.createMenu('🔄 Vigencias')
      .addItem('🔄 Recalcular', 'recalcularTodo')
      .addItem('🚨 Alertas', 'mostrarAlertas'))
    .addToUi()

  ui.createMenu('📥 Ingresos')
    .addItem('🎨 Formatear hoja', 'corregirYMejorarIngresos')
    .addSeparator()
    .addItem('⚙️ Confirmar "INGRESA"', 'configurarTriggerIngresos')
    .addSeparator()
    .addItem('📤 Enviar a Pacientes', 'enviarIngresasAPacientes')
    .addToUi()

  ui.createMenu('📝 Formulario')
    .addItem('📋 Crear hoja', 'crearFormularioBase')
    .addItem('📊 Estadísticas', 'mostrarSidebarEstadisticas')
    .addSeparator()
    .addItem('✅ Aprobar', 'aprobarFormularios')
    .addItem('❌ Rechazar', 'rechazarFormularios')
    .addItem('🧹 Limpiar', 'limpiarFormulariosAprobados')
    .addSeparator()
    .addItem('⚙️ Recepción automática', 'configurarTriggerFormulario')
    .addToUi()

  ui.createMenu('🔧 Herramientas')
    .addItem('🚀 Inicializar', 'inicializarPADDS')
    .addSeparator()
    .addSubMenu(ui.createMenu('🗂️ Organización')
      .addItem('📑 Organizar hojas', 'organizarHojas')
      .addItem('🎯 Resaltar fila activa', 'configurarResaltadoFila'))
    .addItem('⚙️ Parámetros', 'crearParametros')
    .addSubMenu(ui.createMenu('📊 Dashboard')
      .addItem('📊 Crear', 'crearDashboard')
      .addItem('🔄 Actualizar', 'actualizarDashboard')
      .addItem('🩹 Reparar fórmulas', 'repararFormulas'))
    .addSubMenu(ui.createMenu('💾 Backups')
      .addItem('⚙️ Configurar', 'configurarBackupAutomatico')
      .addItem('📦 Hacer ahora', 'ejecutarBackup'))
    .addSeparator()
    .addItem('🧪 Diagnóstico', 'ejecutarPruebas')
    .addToUi()

  ui.createMenu('ℹ️ Ayuda')
    .addItem('❓ Ayuda', 'mostrarAyudaRapida')
    .addItem('📋 Referencia columnas', 'crearInstrucciones')
    .addItem('🔄 Recargar menús', 'rehacerMenus')
    .addSeparator()
    .addItem('ℹ️ Acerca de', 'acercaDelSistema')
    .addToUi()
}

function rehacerMenus() {
  try { onOpen() } catch (e) {
    SpreadsheetApp.getUi().alert('No se pudo recargar: ' + e.message)
  }
}

function mostrarAyudaRapida() {
  var css = '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:0;padding:14px;background:#F1F5F9;color:#202124;font-size:13px;line-height:1.55}' +
    'h2{font-size:18px;color:#1E293B;margin:0 0 2px;letter-spacing:.2px}' +
    '.sub{font-size:11.5px;color:#64748B;margin:0 0 12px}' +
    '.card{background:#fff;border-radius:12px;box-shadow:0 1px 3px rgba(0,0,0,.10);margin-bottom:10px;overflow:hidden}' +
    '.card>summary{list-style:none;cursor:pointer;padding:11px 14px;font-weight:600;font-size:13.5px;display:flex;align-items:center;gap:8px;border-left:6px solid var(--c);user-select:none}' +
    '.card>summary::-webkit-details-marker{display:none}' +
    '.card>summary::after{content:"▾";margin-left:auto;color:#94A3B8;font-size:12px;transition:transform .15s}' +
    '.card:not([open])>summary::after{transform:rotate(-90deg)}' +
    '.card[open]>summary{border-bottom:1px solid #F1F5F9}' +
    '.bd{padding:6px 14px 12px;font-size:12.5px}' +
    '.bd p{margin:7px 0}' +
    'b{color:#1E293B}' +
    '.k{display:inline-block;background:#F8FAFC;border:1px solid #CBD5E1;border-bottom-width:2px;border-radius:5px;padding:0 5px;font-size:11px;font-family:Consolas,monospace}' +
    '.chip{display:inline-block;background:#E0F2FE;color:#0369A1;border-radius:20px;padding:1px 9px;font-size:11px;margin:0 4px 4px 0;white-space:nowrap}' +
    '.paso{display:flex;gap:8px;margin:6px 0;align-items:baseline}' +
    '.paso .n{flex:0 0 18px;height:18px;border-radius:50%;background:var(--c);color:#fff;font-size:11px;font-weight:700;text-align:center;line-height:18px;margin-top:2px}' +
    '.tip{background:#FEF3C7;border:1px solid #FDE68A;color:#A16207;border-radius:8px;padding:6px 10px;font-size:12px;margin:8px 0}' +
    '.aviso{background:#FEE2E2;border:1px solid #FCA5A5;color:#B91C1C;border-radius:8px;padding:6px 10px;font-size:12px;margin:8px 0}' +
    '.ley{display:flex;flex-wrap:wrap;gap:4px;margin:6px 0}' +
    '.ley .l{display:inline-flex;align-items:center;gap:5px;background:#F8FAFC;border:1px solid #E2E8F0;border-radius:6px;padding:2px 7px;font-size:11px}' +
    '.dot{width:9px;height:9px;border-radius:3px;display:inline-block}' +
    '.men{color:#64748B;font-size:11.5px;font-style:italic}' +
    '</style>'

  var secs = [
    { c: '#B45309', t: '📅 Agenda', open: false,
      h: '<p><b>¿Qué es?</b> Planifica visitas domiciliarias por semana; cada profesional tiene su columna y el horario es fijo de 8:00 a 16:00.</p>' +
         '<p><span class="chip">➕ Agregar semana</span><span class="chip">➕ Más semanas ▸</span><span class="chip">❌ Eliminar semana</span><span class="chip">🧹 Borrar semanas pasadas</span></p>' +
         '<p><b>Escribir atención:</b> escribe <b>VDI</b>, abreviada o <b>REGISTRO</b> en la celda del horario. Cabecera oscura = nombre del profesional.</p>' +
         '<p><b>Limpiar una semana:</b> selecciona las celdas (✕/📦) y presiona <span class="k">Delete</span>.</p>' +
         '<p><b>Navegar:</b> <span class="k">🎯 Ir a hoy</span>, <span class="k">📅 Ir a fecha</span> o <span class="k">🔎 Buscar paciente</span>.</p>' },
    { c: '#0369A1', t: '🩺 Pacientes', open: true,
      h: '<p><b>¿Qué es?</b> La ficha central: datos personales, patologías, controles y cuidador (111 campos por secciones de color).</p>' +
         '<p><b>Buscador vivo (fila 2):</b> escribe en <b>B2</b> nombre, RUN o teléfono y presiona <span class="k">Enter</span> — las filas que no coinciden se ocultan; <b>D2</b> muestra el conteo. <b>Borra B2</b> para ver todo.</p>' +
         '<p><b>Mostrar secciones:</b> a la derecha del buscador, en <b>F2</b> (una celda) un <b>menú desplegable</b> elige qué sección ver: las demás columnas se ocultan temporalmente y se ven solo las de esa sección. <b>"TODAS"</b> restaura todo. La zona del buscador (columnas 1-17) siempre queda visible.</p>' +
         '<p><b>Estados de vigencia:</b> cada control es UNA fecha; el color dice su estado:</p>' +
         '<div class="ley"><span class="l"><i class="dot" style="background:#15803D"></i>AL DÍA</span><span class="l"><i class="dot" style="background:#C2410C"></i>POR VENCER</span><span class="l"><i class="dot" style="background:#B91C1C"></i>VENCIDO</span><span class="l"><i class="dot" style="background:#B45309"></i>PENDIENTE</span><span class="l"><i class="dot" style="background:#CBD5E1"></i>N/A</span></div>' +
         '<div class="tip">💡 Plazos y días de aviso: 🔧 Herramientas → ⚙️ Parámetros de vigencia.</div>' },
    { c: '#0F766E', t: '📥 Ingresos', open: false,
      h: '<div class="paso"><span class="n">1</span><span>🎨 <b>Formatear hoja</b> (crea la hoja si no existe)</span></div>' +
         '<div class="paso"><span class="n">2</span><span>Escribe la fecha y el <b>RUN</b> del paciente</span></div>' +
         '<div class="paso"><span class="n">3</span><span>En OBSERVACION escribe <b>INGRESA</b> para confirmar el ingreso</span></div>' +
         '<div class="paso"><span class="n">4</span><span>📤 <b>Enviar a Pacientes</b>: verifica RUT, evita duplicados y pasa la fila a Pacientes</span></div>' +
         '<div class="aviso">⚠️ La confirmación automática al escribir INGRESA se activa con "⚙️ Confirmar INGRESA".</div>' +
         '<div class="tip">💡 La hoja incluye panel de estadísticas en vivo (total, pendientes, en espera, gestionados, por ingresar) y filtros automáticos con flechas en los encabezados.</div>' },
    { c: '#7E22CE', t: '📝 Formulario', open: false,
      h: '<div class="paso"><span class="n">1</span><span>📋 <b>Crear hoja</b> (primera vez)</span></div>' +
         '<div class="paso"><span class="n">2</span><span>Los formularios digitales llegan solos (recepción automática)</span></div>' +
         '<div class="paso"><span class="n">3</span><span>Revisa los datos en la hoja</span></div>' +
         '<div class="paso"><span class="n">4</span><span>✅ <b>Aprobar</b> (pasa a Pacientes)</span></div>' +
         '<div class="paso"><span class="n">5</span><span>❌ <b>Rechazar</b> · 🧹 <b>Limpiar</b> solicitudes ya procesadas</span></div>' },
    { c: '#475569', t: '🔧 Herramientas', open: false,
      h: '<p><b>🚀 Inicializar:</b> al primer uso, crea todas las hojas del sistema (Parámetros, Ingresos, Recepción, Dashboard, Referencia Columnas, Agenda, _Resalte) y formatea ' + HOJA_PAC + ' de una sola vez.</p>' +
         '<p><b>🗂️ Organización:</b> ordena y colorea pestañas · activa el resaltado de la fila activa.</p>' +
         '<p><b>🩺 Pacientes → 🛠️ Datos:</b> 🎨 <b>Formatear hoja</b> aplica el diseño completo de la plantilla (encabezados, secciones, buscador, validaciones, colores) sin perder datos. 🧹 <b>Reparar formato</b> se usa cuando los datos quedan bien pero colores/casillas/fondos quedaron en columnas equivocadas (por pegar con formato): limpia y re-aplica el diseño a las posiciones correctas, sin mover datos.</p>' +
         '<p><b>⚙️ Parámetros:</b> plazos de cada control en MESES, DÍAS o N/A (desactivado) y días de aviso para "POR VENCER".</p>' +
         '<p><b>📊 Dashboard:</b> estadísticas en vivo con tarjetas (totales, demografía, prioridad, sector, dependencia, EMPA/EMPAM), tablas de vacunación y patologías con faltante y %, vigencias por área y <b>gráficos</b> (sector en barras verticales y sexo). Si aparece #REF!, usa "🩹 Reparar fórmulas".</p>' },
    { c: '#B91C1C', t: '🚨 Alertas', open: false,
      h: '<p>Panel visual de pacientes con controles <b>VENCIDOS</b>, <b>POR VENCER</b>, <b>PENDIENTES</b> y <b>PRIORITARIOS</b>. Haz clic en un paciente para ir a su fila.</p>' +
         '<p class="men">Acceso: 🩺 Pacientes → 🔄 Vigencias → 🚨 Alertas.</p>' },
    { c: '#334155', t: '👨‍⚕️ Créditos', open: false,
      h: '<p><b>Sistema PADDS</b> — Planilla de Atención Domiciliaria para CESFAM.</p>' +
         '<p><b>Creador:</b> Patricio Varela C. · <b>Cargo:</b> Interno TENS</p>' +
         '<p><b>Contacto:</b> patriciovarelacontreras@gmail.com</p>' +
         '<p><b>Versión:</b> v3.1 · <b>Última actualización:</b> 12/08/2026</p>' },
  ]

  var cards = ''
  for (var i = 0; i < secs.length; i++) {
    cards += '<details class="card" style="--c:' + secs[i].c + '"' + (secs[i].open ? ' open' : '') + '>' +
      '<summary>' + secs[i].t + '</summary><div class="bd">' + secs[i].h + '</div></details>'
  }

  var html = '<html><head><base target="_top">' + css + '</head><body>' +
    '<h2>Ayuda rápida</h2>' +
    '<p class="sub">Sistema PADDS · Haz clic en cada sección para abrirla</p>' + cards +
    '</body></html>'

  var output = HtmlService.createHtmlOutput(html).setTitle('Ayuda Rápida - PADDS').setWidth(420)
  SpreadsheetApp.getUi().showSidebar(output)
}

function acercaDelSistema() {
  var html = '<html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:0;padding:16px;background:#F1F5F9;color:#202124;font-size:13px;line-height:1.55}' +
    'h2{font-size:19px;color:#1E293B;margin:0}' +
    '.ver{font-size:11.5px;color:#64748B;margin:2px 0 14px}' +
    '.m{display:flex;gap:10px;background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);padding:10px 12px;margin-bottom:8px}' +
    '.m .ic{flex:0 0 30px;height:30px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;color:#fff}' +
    '.m b{display:block;font-size:13px;margin-bottom:1px}' +
    '.m span{font-size:12px;color:#64748B}' +
    '.f{font-size:11px;color:#94A3B8;text-align:center;margin-top:14px}' +
    '</style></head><body>' +
    '<h2>Sistema PADDS</h2>' +
    '<div class="ver">Planilla de Atención Domiciliaria para CESFAM · v3.1</div>' +
    '<div class="m"><div class="ic" style="background:#1E293B">👨‍⚕️</div><div><b>Creador</b><span>Patricio Varela C. · Cargo: Interno TENS</span></div></div>' +
    '<div class="m"><div class="ic" style="background:#475569">📧</div><div><b>Contacto</b><span>patriciovarelacontreras@gmail.com</span></div></div>' +
    '<div class="m"><div class="ic" style="background:#0F766E">📅</div><div><b>Última actualización</b><span>12/08/2026 · versión 3.1</span></div></div>' +
    '<div class="m"><div class="ic" style="background:#0F766E">🩺</div><div><b>Pacientes</b><span>Ficha única con datos, patologías, cuidador, controles, signos vitales y priorización automática. 111 campos por secciones de color.</span></div></div>' +
    '<div class="m"><div class="ic" style="background:#B45309">📅</div><div><b>Agenda</b><span>Planificación semanal de visitas, con profesionales en columnas paralelas y tipos de atención.</span></div></div>' +
    '<div class="m"><div class="ic" style="background:#7E22CE">📝</div><div><b>Formulario</b><span>Recepción de solicitudes digitales con aprobación o rechazo en lote y pase automático a Pacientes.</span></div></div>' +
    '<div class="m"><div class="ic" style="background:#B91C1C">🚨</div><div><b>Alertas</b><span>Panel de controles vencidos, por vencer, pendientes y pacientes prioritarios. Navegación con un clic.</span></div></div>' +
    '<div class="m"><div class="ic" style="background:#15803D">📊</div><div><b>Dashboard</b><span>Estadísticas en vivo: totales, demografía, patologías y resumen de vigencias.</span></div></div>' +
    '<div class="f">© 2026 Patricio A. Varela C. · Interno TENS · Contacto: patriciovarelacontreras@gmail.com</div>' +
    '</body></html>'

  var output = HtmlService.createHtmlOutput(html).setTitle('Acerca del Sistema').setWidth(400)
  SpreadsheetApp.getUi().showSidebar(output)
}

// ─── ORDENAR Y COLOREAR HOJAS ────────────────────────────────────────────────

function organizarHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  var resp = ui.alert('Ordenar y colorear pestañas',
    'Esto SOLO ordena, renombra (nombres estándar en título) y colorea\n' +
    'las pestañas listadas. Las fórmulas que apunten a una pestaña\n' +
    'renombrada se actualizan automáticamente.\n' +
    'NO toca el diseño de la hoja ' + HOJA_PAC + ' ni sus datos.\n' +
    'NO crea ni elimina ninguna hoja: las que no están listadas quedan\n' +
    'donde están, sin color y sin tocarse.\n\n' +
    'Para el diseño completo de la plantilla (buscador, secciones, validaciones) ' +
    'usa 🩺 Pacientes → 🛠️ Datos → "🎨 Formatear hoja".\n\n¿Continuar?',
    ui.ButtonSet.YES_NO)
  if (resp !== ui.Button.YES) return
  _organizarHojasSinConfirmar(ss)
}

function _organizarHojasSinConfirmar(ss) {
  var orden = [
    ['Ingresos',                     '#0F766E'],
    [HOJA_PAC,                       '#64748B'],
    ['Agenda Profesionales',         '#B5A67C'],
    [HOJA_FORM,                      '#64748B'],
    ['Dashboard',                    '#0F766E'],
    ['Referencia Columnas',          '#A3927F'],
    ['Parámetros',                   '#8C959D'],
    ['Config',                       '#A5ABB0'],
    ['_PlantillaSemana',             '#B2B2B2'],
    ['_Resalte',                     '#C2C2C2'],
  ]

  var activa = ''
  try { activa = ss.getActiveSheet().getName() } catch (e) {}

  var hojas = ss.getSheets()
  var idx = {}
  for (var hi = 0; hi < hojas.length; hi++) {
    idx[_normalizarNombre(hojas[hi].getName())] = hojas[hi]
  }

  var renombradas = 0
  for (var ri = 0; ri < orden.length; ri++) {
    var canon = orden[ri][0]
    var shR = idx[_normalizarNombre(canon)]
    if (!shR) continue
    var actual = shR.getName()
    if (actual !== canon) {
      shR.setName(canon)
      _reescribirRefsPestana(ss, actual, canon)
      renombradas++
    }
  }

  var presentes = 0
  for (var i = 0; i < orden.length; i++) {
    var sh = idx[_normalizarNombre(orden[i][0])]
    if (sh) {
      sh.setTabColor(orden[i][1])
      presentes++
    }
  }

  // se dejan en su sitio. AQUÍ NUNCA SE BORRA NI SE CREA UNA HOJA.
  var pos = 1
  for (var i = 0; i < orden.length; i++) {
    var sh2 = idx[_normalizarNombre(orden[i][0])]
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

  try {
    var back = idx[_normalizarNombre(activa)]
    if (back) ss.setActiveSheet(back, false)
  } catch (e) {}

  ss.toast((renombradas > 0 ? renombradas + ' pestañas renombradas · ' : '') +
    presentes + ' hojas coloreadas. Las demás no se tocaron', '', 5)
}

// nuevo nombre canónico. Solo toca fórmulas: nunca borra ni crea hojas.
function _reescribirRefsPestana(ss, viejo, nuevo) {
  var frag = viejo.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  var re = new RegExp("'?" + frag + "'?(?=!)", 'gi')
  var total = 0
  var hojas = ss.getSheets()
  for (var i = 0; i < hojas.length; i++) {
    var sh = hojas[i]
    var rng = sh.getDataRange()
    var f = rng.getFormulas()
    var cambio = false
    for (var r = 0; r < f.length; r++) {
      for (var c = 0; c < f[r].length; c++) {
        if (!f[r][c]) continue
        var nf = f[r][c].replace(re, "'" + nuevo + "'")
        if (nf !== f[r][c]) { f[r][c] = nf; cambio = true; total++ }
      }
    }
    if (cambio) rng.setFormulas(f)
  }
  return total
}

function _normalizarNombre(nombre) {
  return _quitarAcentos(String(nombre || '').toLowerCase())
}

// ─── INICIALIZACIÓN COMPLETA (crear hojas + formatear Pacientes) ──────────────

function inicializarPADDS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()

  var shPac = ss.getSheetByName(HOJA_PAC)
  var avisoPac = shPac
    ? ''
    : '\n\n⚠️ No existe la hoja "' + HOJA_PAC + '". Créala (Insertar → Hoja), nómbrala "' + HOJA_PAC +
      '" con tus datos o renómbrala, y vuelve a ejecutar esta opción.'

  var resp = ui.alert('🚀 Inicializar PADDS',
    'Esto prepara la planilla para su primer uso:\n' +
    '· ' + HOJA_PAC + ': agrega columnas faltantes y aplica el diseño completo\n' +
    '  (validaciones, secciones, buscador, colores, anchos; elimina filas vacías)\n' +
    '· Crea/actualiza: Parámetros · Ingresos · ' + HOJA_FORM + ' · Dashboard\n' +
    '  · Referencia Columnas · Agenda Profesionales · _Resalte\n' +
    '· Ordena y colorea las pestañas\n\n' +
    '⚠️ Parámetros, Dashboard y Referencia Columnas se REGENERAN (valores por defecto).\n' + avisoPac,
    ui.ButtonSet.YES_NO)
  if (resp !== ui.Button.YES) return
  if (!shPac) {
    ui.alert('Inicializar PADDS', 'No se puede inicializar sin la hoja "' + HOJA_PAC + '".\n' + avisoPac, ui.ButtonSet.OK)
    return
  }

  var resumen = []

  try {
    _agregarColumnasFaltantes(ss, shPac, _COLUMNAS._count || 112)
    resumen.push('✔ ' + HOJA_PAC + ': columnas listas (' + (_COLUMNAS._count || 112) + ')')
  } catch (e) { resumen.push('✘ ' + HOJA_PAC + ' columnas: ' + e.message) }

  try { crearParametros(); resumen.push('✔ Parámetros regenerada') }
  catch (e) { resumen.push('✘ Parámetros: ' + e.message) }

  try { _crearHojaIngresosSinConfirmar(ss); resumen.push('✔ Ingresos lista') }
  catch (e) { resumen.push('✘ Ingresos: ' + e.message) }

  try { crearFormularioBase(); resumen.push('✔ ' + HOJA_FORM + ' lista') }
  catch (e) { resumen.push('✘ ' + HOJA_FORM + ': ' + e.message) }

  try {
    if (!ss.getSheetByName(HOJA)) ss.insertSheet(HOJA)
    resumen.push('✔ ' + HOJA + ' lista (la primera semana se agrega con 📅 Agenda → ➕)')
  } catch (e) { resumen.push('✘ ' + HOJA + ': ' + e.message) }

  try { _formatearDatosSinConfirmar(shPac, ss); resumen.push('✔ Formato de ' + HOJA_PAC + ' aplicado') }
  catch (e) { resumen.push('✘ Formato ' + HOJA_PAC + ': ' + e.message) }

  try { crearDashboard(); resumen.push('✔ Dashboard creada') }
  catch (e) { resumen.push('✘ Dashboard: ' + e.message) }

  try { crearInstrucciones(); resumen.push('✔ Referencia Columnas creada') }
  catch (e) { resumen.push('✘ Referencia Columnas: ' + e.message) }

  try { configurarResaltadoFila(); resumen.push('✔ Resaltado de fila activa') }
  catch (e) { resumen.push('✘ _Resalte: ' + e.message) }

  try { _organizarHojasSinConfirmar(ss); resumen.push('✔ Pestañas ordenadas y coloreadas') }
  catch (e) { resumen.push('✘ Organizar hojas: ' + e.message) }

  try { onOpen() } catch (e) {}

  var ok = 0, fail = 0
  for (var i = 0; i < resumen.length; i++) { if (resumen[i].indexOf('✘') === 0) fail++; else ok++ }

  ss.toast('Inicialización PADDS: ' + ok + ' pasos OK' + (fail ? ', ' + fail + ' con error' : ''), 'PADDS', 8)
  ui.alert('Inicialización PADDS',
    (fail ? '⚠️ Algunos pasos fallaron:\n\n' : '✅ Inicialización completada:\n\n') + resumen.join('\n'),
    ui.ButtonSet.OK)
}
