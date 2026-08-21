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
    chk('Pacientes: ' + lcPac + ' columnas (esperadas ' + (_COLUMNAS._count || 112) + ')',
      lcPac >= (_COLUMNAS._count || 112))
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
      chk('Pacientes: columna INSULINO DEPENDIENTE presente',
        lcPac >= 56 && String(heads[55] || '').toUpperCase().indexOf('INSULINO') >= 0)
    }
    chk('Pacientes: con datos (fila 4+)', pac.getLastRow() >= 4)
  }

  // ── 2. CONSTANTES Y MAPAS ──────────────────────────────
  chk('Guía de columnas: 112 registradas', (_COLUMNAS._count || 0) === 112)
  var claves = ['ID', 'SECTOR', 'NOMBRE', 'RUN', 'VITAL', 'PRIORIDAD', 'OBSERVACIONES', 'EDITOR']
  for (var c = 0; c < claves.length; c++) {
    chk('COL.' + claves[c] + ' definida', COL[claves[c]] !== undefined)
  }
  chk('Controles: _FECHA_BY_COL cubre todos', Object.keys(_FECHA_BY_COL).length >= _CONTROL_FECHAS.length)

  var okVac = true, detVac = ''
  for (var vc = 0; vc < _VACUNA_COLS.length; vc++) {
    if (PAC_VALIDACIONES[_VACUNA_COLS[vc]] !== _VACUNA_VALS) { okVac = false; detVac += _VACUNA_COLS[vc] + ' ' }
  }
  chk('Inmunización 81-84: dropdown SI·NO·N/A·R·P', okVac, detVac)

  var okLibre = true, detLibre = ''
  for (var li = 0; li < PAC_LIBRES.length; li++) {
    var lib = PAC_LIBRES[li]
    if (PAC_VALIDACIONES[lib] || _FECHAS_VA.indexOf(lib) >= 0 || _CHECKBOX_COLS.indexOf(lib) >= 0) {
      okLibre = false; detLibre += lib + ' '
    }
  }
  chk('PAC_LIBRES: sin columnas de dropdown/fecha/casilla', okLibre, detLibre)

  var okSis = true, detSis = ''
  if (HOJAS_SISTEMA.indexOf(HOJA) >= 0) { okSis = false; detSis += HOJA + ' no debe estar listada ' }
  if (HOJAS_SISTEMA.indexOf(HOJA_PAC) < 0) { okSis = false; detSis += HOJA_PAC + ' debe estar listada ' }
  if (HOJAS_SISTEMA.indexOf(HOJA_FORM) < 0) { okSis = false; detSis += HOJA_FORM + ' debe estar listada ' }
  chk('HOJAS_SISTEMA: excluye Agenda, incluye hojas de sistema', okSis, detSis)

  var okServ = true, detServ = ''
  for (var sk in _SERVICIO_COL_MAP) {
    var pc = _SERVICIO_COL_MAP[sk]
    if (!(pc >= 1 && pc <= 112)) { okServ = false; detServ += sk + '→' + pc + ' ' }
  }
  chk('Servicios del formulario → columna válida', okServ, detServ)

  var okPac = true, detPac = ''
  for (var fi = 0; fi < FORM_A_PAC.length; fi++) {
    var fp = FORM_A_PAC[fi]
    if (!(fp[1] == null || (fp[1] >= 1 && fp[1] <= 112))) { okPac = false; detPac += fi + ' ' }
  }
  chk('Mapa formulario→Pacientes válido', okPac, detPac)

  var okFfm = true, detFfm = ''
  for (var fmk in _FORM_FIELD_MAP) {
    var fmc = _FORM_FIELD_MAP[fmk]
    if (!(fmc >= 1 && fmc <= 19)) { okFfm = false; detFfm += fmk + '→' + fmc + ' ' }
  }
  chk('Mapa formulario→columnas de recepción (1-19) válido', okFfm, detFfm)

  var okOp = true, detOp = ''
  var colsOp = Object.keys(PAC_VALIDACIONES)
  for (var oi = 0; oi < colsOp.length; oi++) {
    var listaOp = PAC_VALIDACIONES[colsOp[oi]]
    for (var oj = 0; oj < listaOp.length; oj++) {
      if (!_bgOpcionCol(Number(colsOp[oi]), _claveColorOpcion(listaOp[oj]))) {
        okOp = false; detOp += colsOp[oi] + ':' + listaOp[oj] + ' '
      }
    }
  }
  chk('Dropdown: todas las opciones tienen color', okOp, detOp)
chk('Sector: "PENDIENTE" distinto de "AMARILLO" (gris azulado)', _SECTOR_COLORS['PENDIENTE'] !== undefined && _SECTOR_COLORS['PENDIENTE'][0] === '#F1F5F9' && _SECTOR_COLORS['PENDIENTE'][1] !== _SECTOR_COLORS['AMARILLO'][1])
  chk('Sector: chip "PENDIENTE" por columna (col 2 → #64748B)', _bgOpcionCol(2, 'PENDIENTE') === '#64748B' && _bgOpcionCol(6, 'PENDIENTE') === '#B45309')

  // ── 3. LÓGICA PURA (no toca la hoja) ───────────────────
  chk('RUT sin guión → con guión', formatearRUT('16297925') === '1629792-5')
  chk('RUT con puntos → normal', formatearRUT('15.451.872-K') === '15451872-K')
  chk('Dígito verificador: 1629792-5 válido', _validarDigitoRUT('1629792-5') === true)
  chk('Dígito verificador: 1629792-4 inválido', _validarDigitoRUT('1629792-4') === false)
  chk('_parseDate: fecha válida', _parseDate('15/06/2026') instanceof Date)
  chk('_parseDate: fecha inválida', _parseDate('40/13/2026') === null)
  chk('_calcularEdad: devuelve número', !isNaN(Number(_calcularEdad('15/05/1950'))))
  chk('_edadConMeses: <1 año → "N mes(es)"', /^\d+ mes(es)?$/.test(_edadConMeses('01/01/2026', '')) === true)
  chk('_edadConMeses: sin fecha → edad en años', _edadConMeses('', 33) === '33 años')
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
  chk('colToLetter(112) → DH', colToLetter(112) === 'DH')
  chk('Teléfono móvil: 9 8765 4321', fmtNum('987654321') === '9 8765 4321')
  chk('Teléfono fijo: 2234 5678', fmtNum('22345678') === '2234 5678')
  chk('_estadoFecha: N/A', _estadoFecha('N/A', 6, 30) === 'N/A')
  chk('_estadoFecha: n/a minúscula → N/A', _estadoFecha('n/a', 6, 30) === 'N/A')
  chk('_estadoFecha: vacío → PENDIENTE', _estadoFecha('', 6, 30) === 'PENDIENTE')
  chk('_estadoFecha: texto → PENDIENTE', _estadoFecha('sin fecha', 6, 30) === 'PENDIENTE')
  chk('RUT sin DV (7 dígitos) no se corrompe', formatearRUT('1629792') === '1629792')
  chk('_normRUN: 7 dígitos sin guion se deja tal cual', _normRUN('1629792') === '1629792')
  chk('_fechaCorrupta: fecha válida → false', _fechaCorrupta('15/06/2026', false) === false)
  chk('_fechaCorrupta: 40/13/2026 → mal', _fechaCorrupta('40/13/2026', false) === 'mal')
  chk('_fechaCorrupta: 31/12/1969 → epoch', _fechaCorrupta('31/12/1969', false) === 'epoch')
  chk('_fechaCorrupta: año 2500 → futuro', _fechaCorrupta('01/01/2500', false) === 'futuro')
  chk('_fechaCorrupta: texto plano → false', _fechaCorrupta('PENDIENTE', false) === false)

  var hNow = new Date()
  var fPas = new Date(hNow.getFullYear() - 1, hNow.getMonth(), hNow.getDate())
  var fPv = new Date(hNow.getFullYear(), hNow.getMonth() - 6, hNow.getDate() + 10)
  var fDia = new Date(hNow.getTime() + 60 * 86400000)
  chk('calcStatus: fecha pasada → VENCIDO', calcStatus(fPas, 6, '', 30) === 'VENCIDO')
  chk('calcStatus: próxima semana → POR VENCER', calcStatus(fPv, 6, '', 30) === 'POR VENCER')
  chk('calcStatus: en 2 meses → AL DIA', calcStatus(fDia, 6, '', 30) === 'AL DIA')

  // ── 4. MENÚS: funciones referenciadas existen ──────────
  chk('Ingresos: fuentes de verdad de dropdowns', _ING_ACCIONES.join() === 'INGRESA,NO INGRESA,PENDIENTE' && _ING_ESTADOS.join() === 'GESTIONADOS,PENDIENTES,EN ESPERA')
  chk('Ingresos: validaciones centralizadas', typeof _ingAplicarValidaciones === 'function')
  chk('Ingresos: cebra centralizada', typeof _ingPintarCebra === 'function')
  var handlers = [
    'generarSemana', 'generarSemanaEspecifica', 'generarTresSemanas', 'eliminarSemana',
    'limpiarSemanasPasadas', 'irAHoy', 'irAFecha', 'buscarPaciente', 'resumen',
    'agregarPaciente', 'buscarPacienteGlobal', 'eliminarPaciente', 'ordenarPacientes',
    'reindexarPacientes', 'limpiarPacientesCompleto', 'formatearDatos',
    'repararFormatoPacientes', 'verificarIntegridadRUN', 'formatearRUTPacientes',
    'depurarDuplicados', 'limpiarFilasVaciasPacientes', 'diagnosticarFilasExtra',
    'recalcularTodo', 'verFichaPaciente', 'corregirYMejorarIngresos',
    'configurarTriggerIngresos', 'enviarIngresasAPacientes',
    'crearFormularioBase', 'mostrarSidebarEstadisticas', 'aprobarFormularios',
    'rechazarFormularios', 'limpiarFormulariosAprobados', 'configurarTriggerFormulario',
    'organizarHojas', 'configurarResaltadoFila', 'crearParametros',
    'crearDashboard', 'actualizarDashboard', 'repararFormulas', 'ejecutarPruebas',
    'restablecerPADDS', 'configurarBackupAutomatico', 'ejecutarBackup',
    'mostrarAyudaRapida', 'crearInstrucciones', 'rehacerMenus', 'acercaDelSistema',
    'abrirCentroControl', 'buscarPacienteGlobal', 'verFichaPacienteV4',
    'abrirDashboardV4', 'mostrarAlertasV4', 'mostrarTareas',
    'mostrarAuditoriaDatos', 'mostrarGuiaColumnas', 'actualizarSistemaCompleto'
  ]
  var faltantes = []
  for (var m = 0; m < handlers.length; m++) {
    if (!fnExiste(handlers[m])) faltantes.push(handlers[m])
  }
  chk('Menús: todas las funciones existen', faltantes.length === 0,
    faltantes.length ? 'Faltan: ' + faltantes.join(', ') : '')

  // ── 5. UNICODE / ENCODING (regresión de mojibake) ──────
  var MOJI = /[ÃÂðâœ\x00-\x1F\x80-\x9F]/
  var digUI = []
  try { digUI = digestUIStrings() } catch (eDG) { digUI = [] }
  chk('Unicode: digestUIStrings disponible', digUI.length > 0)
  var okDig = true, detDig = ''
  for (var dg = 0; dg < digUI.length; dg++) {
    if (MOJI.test(digUI[dg] || '')) { okDig = false; detDig += digUI[dg] + ' ' }
  }
  chk('Unicode: digestUIStrings sin mojibake', okDig, detDig)
  var okHtml = true, detHtml = ''
  try {
    var uiHTML = String(_uiV4Shell('Prueba · PADDS', '<div>🔎 á é í ó ú ñ · → ✓</div>').getContent())
    if (MOJI.test(uiHTML)) { okHtml = false; detHtml = uiHTML }
  } catch (eH) { okHtml = false; detHtml = 'no generó HTML: ' + eH.message }
  chk('Unicode: HTML generado por _uiV4Shell limpio', okHtml, detHtml)
  chk('Unicode: quitarAcentos "Mañana Árbol" seguro', _quitarAcentos('Mañana Árbol Núñez') === 'Manana Arbol Nunez')

  // ── REPORTE ────────────────────────────────────────────
  var okN = 0, failN = 0
  for (var r = 0; r < res.length; r++) {
    if (res[r].ok) okN++; else failN++
  }
  var rows = ''
  for (var r2 = 0; r2 < res.length; r2++) {
    rows += '<tr><td style="color:' + (res[r2].ok ? '#15803D' : '#B91C1C') + '">' +
      (res[r2].ok ? '&#10003;' : '&#10007;') + '</td><td>' + _esc(res[r2].n) + '</td>' +
      (res[r2].ok ? '<td></td>' : '<td style="color:#B91C1C">' + _esc(res[r2].d) + '</td>') + '</tr>'
  }
  var html = '<html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:0;padding:14px;background:#F1F5F9;color:#1E293B;font-size:12.5px}' +
    'h2{margin:0 0 2px;font-size:17px;color:#1E293B}' +
    '.sub{font-size:11px;color:#64748B;margin-bottom:10px}' +
    '.sum{display:flex;gap:8px;margin-bottom:10px}' +
    '.sum .ch{flex:1;border-radius:10px;padding:8px;text-align:center;font-weight:700;color:#fff;font-size:13px}' +
    'table{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden}' +
    'td{padding:4px 8px;border-bottom:1px solid #F1F5F9}' +
    'td:first-child{width:24px;text-align:center}' +
    '</style></head><body>' +
    '<h2>Diagnóstico del sistema</h2>' +
    '<div class="sub">' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + ' · solo revisión, no modifica nada</div>' +
    '<div class="sum">' +
    '<div class="ch" style="background:' + (okN === res.length ? '#15803D' : '#C2410C') + '">' + okN + ' OK</div>' +
    '<div class="ch" style="background:' + (failN ? '#B91C1C' : '#64748B') + '">' + failN + ' pendientes</div>' +
    '<div class="ch" style="background:#1E293B">' + res.length + ' total</div></div>' +
    '<table>' + rows + '</table>' +
    '</body></html>'
  ui.showSidebar(HtmlService.createHtmlOutput(html).setTitle('Diagnóstico PADDS').setWidth(400))
  ss.toast(failN === 0
    ? 'Diagnóstico: ' + okN + '/' + res.length + ' verificaciones OK'
    : 'Diagnóstico: ' + failN + ' puntos a revisar de ' + res.length, 'PADDS', 6)
}

