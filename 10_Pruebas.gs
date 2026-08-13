// ════════════════════════════════════════════════════════
// ÍNDICE ▏10_Pruebas.gs │ autodiagnóstico PADDS (solo lectura)
//
// ⚠️ REGLA DE SEGURIDAD (obligatoria, no eliminar):
// ESTE ARCHIVO NUNCA MODIFICA NADA. Solo revisa y reporta:
// estructura de hojas, constantes, lógica pura y menús.
// La implementación real de la hoja Ingresos vive en
// 11_Ingresos.gs (este archivo NO la duplica).

function ejecutarPruebas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ui = SpreadsheetApp.getUi()
  var res = []

  function chk(nombre, cond, det) { res.push({ n: nombre, ok: !!cond, d: det || '' }) }
  function fnExiste(name) {
    try { return typeof globalThis[name] === 'function' } catch (e) { return true }
  }

  // ── 1. ESTRUCTURA DE HOJAS ─────────────────────────────
  var esperadas = [HOJA_PAC, 'Ingresos', HOJA, HOJA_FORM, 'Dashboard', 'Parámetros', 'Referencia Columnas']
  for (var i = 0; i < esperadas.length; i++) {
    chk('Hoja presente: ' + esperadas[i], ss.getSheetByName(esperadas[i]) != null)
  }

  var pac = ss.getSheetByName(HOJA_PAC)
  if (pac) {
    var lcPac = pac.getLastColumn()
    chk('Pacientes: ' + lcPac + ' columnas (esperadas ' + (_COLUMNAS._count || 111) + ')',
      lcPac >= (_COLUMNAS._count || 111))
    if (lcPac >= 3) {
      var heads = pac.getRange(3, 1, 1, lcPac).getValues()[0]
      var clavesCab = ['NOMBRE', 'RUN', 'ESTADO', 'SECTOR', 'OBSERVACIONES', 'PRIORIDAD']
      var faltanCab = []
      for (var hk = 0; hk < clavesCab.length; hk++) {
        var hallado = false
        for (var h = 0; h < lcPac; h++) {
          if (String(heads[h] || '').toUpperCase().indexOf(clavesCab[hk]) >= 0) { hallado = true; break }
        }
        if (!hallado) faltanCab.push(clavesCab[hk])
      }
      chk('Pacientes: encabezados clave presentes', faltanCab.length === 0,
        faltanCab.length ? 'Faltan: ' + faltanCab.join(', ') : '')
    }
    chk('Pacientes: con datos (fila 4+)', pac.getLastRow() >= 4)
  }

  // ── 2. CONSTANTES Y MAPAS ──────────────────────────────
  chk('Guía de columnas: 111 registradas', (_COLUMNAS._count || 0) === 111)
  var claves = ['ID', 'SECTOR', 'NOMBRE', 'RUN', 'VITAL', 'PRIORIDAD', 'OBSERVACIONES', 'EDITOR']
  for (var c = 0; c < claves.length; c++) {
    chk('COL.' + claves[c] + ' definida', COL[claves[c]] !== undefined)
  }
  chk('Controles: _FECHA_BY_COL cubre todos', Object.keys(_FECHA_BY_COL).length >= _CONTROL_FECHAS.length)

  var okServ = true, detServ = ''
  for (var sk in _SERVICIO_COL_MAP) {
    var pc = _SERVICIO_COL_MAP[sk]
    if (!(pc >= 1 && pc <= 111)) { okServ = false; detServ += sk + '→' + pc + ' ' }
  }
  chk('Servicios del formulario → columna válida', okServ, detServ)

  var okPac = true, detPac = ''
  for (var fi = 0; fi < FORM_A_PAC.length; fi++) {
    var fp = FORM_A_PAC[fi]
    if (!(fp[1] == null || (fp[1] >= 1 && fp[1] <= 111))) { okPac = false; detPac += fi + ' ' }
  }
  chk('Mapa formulario→Pacientes válido', okPac, detPac)

  // ── 3. LÓGICA PURA (no toca la hoja) ───────────────────
  chk('RUT sin guión → con guión', formatearRUT('16297925') === '1629792-5')
  chk('RUT con puntos → normal', formatearRUT('15.451.872-K') === '15451872-K')
  chk('Dígito verificador: 1629792-5 válido', _validarDigitoRUT('1629792-5') === true)
  chk('Dígito verificador: 1629792-4 inválido', _validarDigitoRUT('1629792-4') === false)
  chk('_parseDate: fecha válida', _parseDate('15/06/2026') instanceof Date)
  chk('_parseDate: fecha inválida', _parseDate('40/13/2026') === null)
  chk('_calcularEdad: devuelve número', !isNaN(Number(_calcularEdad('15/05/1950'))))
  chk('EMPA: 30 años → EMPA', _asignarEMPA(30) === 'EMPA')
  chk('EMPA: 70 años → EMPAM', _asignarEMPA(70) === 'EMPAM')
  chk('EMPA: 10 años → N/A', _asignarEMPA(10) === 'N/A')
  chk('Sexo: "Femenino" → F', _normalizarSexo('Femenino') === 'F')
  chk('Sexo: "Hombre" → M', _normalizarSexo('Hombre') === 'M')
  chk('Estado: "fallEcIdO" → FALLECIDO', _normalizarVitalEstado('fallEcIdO') === 'FALLECIDO')
  chk('Estado: "egreso del programa" → EGRESO', _normalizarVitalEstado('egreso del programa') === 'EGRESO')
  chk('Acentos: "Mañana Núñez" → normal', _quitarAcentos('Mañana Árbol Núñez') === 'Manana Arbol Nunez')
  chk('_pad2(5) → "05"', _pad2(5) === '05')
  chk('fmtFecha: 12/08/2026', fmtFecha(new Date(2026, 7, 12)) === '12/08/2026')
  chk('colToLetter(1) → A', colToLetter(1) === 'A')
  chk('colToLetter(27) → AA', colToLetter(27) === 'AA')
  chk('colToLetter(111) → DG', colToLetter(111) === 'DG')
  chk('Teléfono móvil: 9 8765 4321', fmtNum('987654321') === '9 8765 4321')
  chk('Teléfono fijo: 2234 5678', fmtNum('22345678') === '2234 5678')
  chk('_estadoFecha: N/A', _estadoFecha('N/A', 6, 30) === 'N/A')
  chk('_estadoFecha: vacío → PENDIENTE', _estadoFecha('', 6, 30) === 'PENDIENTE')
  chk('_estadoFecha: texto → PENDIENTE', _estadoFecha('sin fecha', 6, 30) === 'PENDIENTE')

  var hNow = new Date()
  var fPas = new Date(hNow.getFullYear() - 1, hNow.getMonth(), hNow.getDate())
  var fPv = new Date(hNow.getFullYear(), hNow.getMonth() - 6, hNow.getDate() + 10)
  var fDia = new Date(hNow.getTime() + 60 * 86400000)
  chk('calcStatus: fecha pasada → VENCIDO', calcStatus(fPas, 6, '', 30) === 'VENCIDO')
  chk('calcStatus: próxima semana → POR VENCER', calcStatus(fPv, 6, '', 30) === 'POR VENCER')
  chk('calcStatus: en 2 meses → AL DIA', calcStatus(fDia, 6, '', 30) === 'AL DIA')

  // ── 4. MENÚS: funciones referenciadas existen ──────────
  var handlers = [
    'generarSemana', 'generarSemanaEspecifica', 'generarTresSemanas', 'eliminarSemana',
    'limpiarSemanasPasadas', 'irAHoy', 'irAFecha', 'buscarPaciente', 'resumen',
    'agregarPaciente', 'buscarEnPacientes', 'eliminarPaciente', 'ordenarPacientes',
    'reindexarPacientes', 'limpiarPacientesCompleto', 'formatearDatos',
    'repararFormatoPacientes', 'verificarIntegridadRUN', 'depurarDuplicados',
    'recalcularTodo', 'mostrarAlertas', 'verFichaPaciente', 'corregirYMejorarIngresos',
    'configurarTriggerIngresos', 'enviarIngresasAPacientes',
    'crearFormularioBase', 'mostrarSidebarEstadisticas', 'aprobarFormularios',
    'rechazarFormularios', 'limpiarFormulariosAprobados', 'configurarTriggerFormulario',
    'organizarHojas', 'configurarResaltadoFila', 'crearParametros',
    'crearDashboard', 'actualizarDashboard', 'repararFormulas', 'ejecutarPruebas',
    'inicializarPADDS',
    'mostrarAyudaRapida', 'crearInstrucciones', 'rehacerMenus', 'acercaDelSistema'
  ]
  var faltantes = []
  for (var m = 0; m < handlers.length; m++) {
    if (!fnExiste(handlers[m])) faltantes.push(handlers[m])
  }
  chk('Menús: todas las funciones existen', faltantes.length === 0,
    faltantes.length ? 'Faltan: ' + faltantes.join(', ') : '')

  // ── REPORTE ────────────────────────────────────────────
  var okN = 0, failN = 0
  for (var r = 0; r < res.length; r++) {
    if (res[r].ok) okN++; else failN++
  }
  var rows = ''
  for (var r2 = 0; r2 < res.length; r2++) {
    rows += '<tr><td style="color:' + (res[r2].ok ? '#2e7d32' : '#c62828') + '">' +
      (res[r2].ok ? '&#10003;' : '&#10007;') + '</td><td>' + _esc(res[r2].n) + '</td>' +
      (res[r2].ok ? '<td></td>' : '<td style="color:#c62828">' + _esc(res[r2].d) + '</td>') + '</tr>'
  }
  var html = '<html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:0;padding:14px;background:#f3f4f8;color:#202124;font-size:12.5px}' +
    'h2{margin:0 0 2px;font-size:17px;color:#1a237e}' +
    '.sub{font-size:11px;color:#5f6368;margin-bottom:10px}' +
    '.sum{display:flex;gap:8px;margin-bottom:10px}' +
    '.chip{flex:1;border-radius:10px;padding:8px;text-align:center;font-weight:700;color:#fff;font-size:13px}' +
    'table{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden}' +
    'td{padding:4px 8px;border-bottom:1px solid #f0f0f0}' +
    'td:first-child{width:24px;text-align:center}' +
    '</style></head><body>' +
    '<h2>Diagnóstico del sistema</h2>' +
    '<div class="sub">' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + ' · solo revisión, no modifica nada</div>' +
    '<div class="sum">' +
    '<div class="chip" style="background:' + (okN === res.length ? '#2e7d32' : '#e65100') + '">' + okN + ' OK</div>' +
    '<div class="chip" style="background:' + (failN ? '#c62828' : '#7b7b7b') + '">' + failN + ' pendientes</div>' +
    '<div class="chip" style="background:#1a237e">' + res.length + ' total</div></div>' +
    '<table>' + rows + '</table>' +
    '</body></html>'
  ui.showSidebar(HtmlService.createHtmlOutput(html).setTitle('Diagnóstico PADDS').setWidth(400))
  ss.toast(failN === 0
    ? 'Diagnóstico: ' + okN + '/' + res.length + ' verificaciones OK'
    : 'Diagnóstico: ' + failN + ' puntos a revisar de ' + res.length, 'PADDS', 6)
}
