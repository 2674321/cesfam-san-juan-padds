// ─── GUÍA DE PRUEBAS MANUAL — PADDS ──────────────────────────────────────────
// Cada paso verifica un cambio específico del plan de modificaciones.
// Marca con ✓ cuando el paso funciona correctamente.

// ═══════════════════════════════════════════════════════════════════════════════
// PRERREQUISITO: La hoja Pacientes debe tener datos reales antes de comenzar.
// ═══════════════════════════════════════════════════════════════════════════════

// ─── PASO 1: VERIFICAR ESTADO INICIAL ───────────────────────────────────────
// ANTES de cualquier cambio, verificar que la hoja esté íntegra.
//
// [ ] 1a. Ir a la hoja PACIENTES
// [ ] 1b. Fila 1: "N°" en A1, "SECTOR" en B1
// [ ] 1c. Fila 2: 🔍 en A2, "SECTOR" en B2
// [ ] 1d. Fila 3: nombres de columna correctos hasta col 111
// [ ] 1e. Filas 4+: datos de pacientes SIN mezclar (ej: SECTOR col B, ESTADO col F)
// [ ] 1f. Columna 111 (EDITOR) existe y tiene datos
// [ ] 1g. Hacer backup de la hoja Pacientes (duplicar)

// ─── PASO 2: VERIFICAR CONSTANTES ────────────────────────────────────────────
// Antes de ejecutar cambios, confirmar que 00_Constantes.gs refleja el layout
// correcto (debe estar actualizado al estado actual de la hoja).
//
// [ ] 2a. COL.EDITOR = 111, COL.OBSERVACIONES = 110, COL.PRIORIDAD = 109
// [ ] 2b. PAC_SECCIONES: CUIDADOR termina en 28
// [ ] 2c. _COLUMNAS: col 28 = "CCV VIGENTE CUIDADOR" (dropdown SI/NO)
// [ ] 2d. _COLUMNAS: col 56 = "MORBILIDAD" (texto libre)
// [ ] 2e. _COLUMNAS: col 79 = "CONTROLES MISCELÁNEOS" (texto libre acumulativo)
// [ ] 2f. _COLUMNAS: col 86 = "CONSULTA PSICOLOGA" (dropdown)
// [ ] 2g. _COLUMNAS: col 91 = "ZONA EVACUACION" (texto libre, sin dropdown)
// [ ] 2h. PAC_VALIDACIONES[91] eliminado o comentado
// [ ] 2i. _COLUMNAS: cols 101-108 = capacitaciones
// [ ] 2j. _COLUMNAS: col 109 = PRIORIDAD GENERAL, 110 = OBSERVACIONES, 111 = EDITOR
// [ ] 2k. FORM_SECCIONES: solo 4 secciones (sin VISITA)
// [ ] 2l. FORM_A_PAC: mapeo correcto incluye COL.EDITOR y misceláneos a OBSERVACIONES
// [ ] 2m. CONTROL_COLS mapeos actualizados a columnas correctas
// [ ] 2n. _SERVICIO_COL_MAP mapeos actualizados

// ─── PASO 3: EJECUTAR RECONSTRUIR PACIENTES ──────────────────────────────────
// Menú → Herramientas → Reconstruir Pacientes
//
// [ ] 3a. No aparecen errores en la consola (Ctrl+Enter)
// [ ] 3b. Datos en filas 4+ NO cambiaron de lugar
// [ ] 3c. SECTOR sigue en col B
// [ ] 3d. ESTADO sigue en col F
// [ ] 3e. NOMBRE sigue en col C
// [ ] 3f. No aparecieron columnas fantasma
// [ ] 3g. Fila 2 preservada (🔍 en A2, SECTOR en B2)
// [ ] 3h. Validaciones funcionan (dropdowns, tooltips, formatos numéricos)

// ─── PASO 4: EJECUTAR FORMATEAR PACIENTES (VISUAL) ──────────────────────────
// Menú → Herramientas → Formatear Pacientes (visual)
//
// [ ] 4a. Se muestra confirmación antes de aplicar
// [ ] 4b. Colores de secciones (filas 1-3) aplicados correctamente
// [ ] 4c. Formato condicional presente
// [ ] 4d. Coloreado EMPA visible
// [ ] 4e. Datos de pacientes (fila 4+) NO se mezclaron
// [ ] 4f. SECTOR col B, ESTADO col F intactos

// ─── PASO 5: PROBAR FORMULARIO DE ATENCIÓN ───────────────────────────────────
// (Requiere que el nuevo formulario esté integrado)
//
// [ ] 5a. Enviar un formulario de atención de prueba
// [ ] 5b. Aparece en "Respuestas Formulario" como Pendiente
// [ ] 5c. Aprobar el formulario
// [ ] 5d. Datos se copian correctamente a Pacientes:
//       - NOMBRE → col C
//       - RUT → col H
//       - NOMBRE CUIDADOR → col P
//       - PROFESIONAL → col 111 (EDITOR)
//       - MISCELONEOS → col 110 (OBSERVACIONES)
// [ ] 5e. Fecha del servicio se escribe en columna correspondiente
// [ ] 5f. Estado del control se actualiza a AL DIA

// ─── PASO 6: PROBAR onEdit — EDITOR ──────────────────────────────────────────
//
// [ ] 6a. Editar OBSERVACIONES (col 110) de un paciente
// [ ] 6b. EDITOR (col 111) se actualiza con el correo del usuario
// [ ] 6c. Editar otra celda cualquiera en la misma fila
// [ ] 6d. EDITOR se actualiza correctamente

// ─── PASO 7: PROBAR LIMPIEZA DE SEMANAS PASADAS ──────────────────────────────
//
// [ ] 7a. Ir a Agenda Profesionales
// [ ] 7b. Ejecutar limpiarSemanasPasadas
// [ ] 7c. Semanas anteriores a la fecha actual se limpian
// [ ] 7d. Semanas futuras permanecen intactas

// ═══════════════════════════════════════════════════════════════════════════════
// PRUEBAS AUTOMATIZADAS (ejecutar desde menú 🧪)
// ═══════════════════════════════════════════════════════════════════════════════

function _assert(cond, msg) {
  if (!cond) throw new Error('FALLÓ: ' + msg)
  console.log('  OK: ' + msg)
}

function _assertEq(a, b, msg) {
  if (a !== b) throw new Error('FALLÓ: ' + msg + ' — esperado "' + b + '", obtenido "' + a + '"')
  console.log('  OK: ' + msg)
}

function _logTitulo(s) {
  console.log('\n=== ' + s + ' ===')
}

function _getPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName('Pacientes')
  if (!sh) throw new Error('No se encuentra hoja Pacientes')
  return sh
}

// ─── TEST 1: COLUMNAS CRÍTICAS EN POSICIÓN CORRECTA ──────────────────────

function testColumnasCriticas() {
  _logTitulo('Constantes: columnas críticas en posición correcta')

  _assertEq(COL.SECTOR, 2, 'SECTOR = col 2')
  _assertEq(COL.VITAL, 6, 'ESTADO (VITAL) = col 6')
  _assertEq(COL.NOMBRE, 3, 'NOMBRE = col 3')
  _assertEq(COL.OBSERVACIONES, 110, 'OBSERVACIONES = col 110')
  _assertEq(COL.EDITOR, 111, 'EDITOR = col 111')
  _assertEq(COL.PRIORIDAD, 109, 'PRIORIDAD = col 109')
  _assertEq(COL.CAPACITACIONES_INI, 101, 'CAPACITACIONES_INI = col 101')
  _assertEq(COL.CAPACITACIONES_FIN, 108, 'CAPACITACIONES_FIN = col 108')
  _assertEq(COL.EMPA_CUIDADOR, 20, 'EMPA_CUIDADOR = col 20')
  _assertEq(COL.EMPA_USUARIO, 81, 'EMPA_USUARIO = col 81')
  _assertEq(COL.CCV, 28, 'CCV VIGENTE CUIDADOR (COL.CCV) = col 28')

  console.log('  Todas las constantes de columna son correctas ✓')
}

// ─── TEST 2: PAC_VALIDACIONES ACTUALIZADAS ─────────────────────────────────

function testValidacionesActualizadas() {
  _logTitulo('Constantes: validaciones actualizadas')

  var vCCV = PAC_VALIDACIONES[COL.CCV]
  _assert(!vCCV, 'col ' + COL.CCV + ' NO debe tener validaciones (CCV VIGENTE CUIDADOR es fecha)')

  var vZona = PAC_VALIDACIONES[COL.ZONA_EVACUACION]
  _assert(!vZona, 'col ' + COL.ZONA_EVACUACION + ' NO debe tener validaciones (zona evacuacion es texto libre)')

  var vPsiq = PAC_VALIDACIONES[COL.CONSULTA_PSICOLOGA]
  _assert(!!vPsiq, 'col ' + COL.CONSULTA_PSICOLOGA + ' debe tener validaciones')
  _assert(vPsiq.indexOf('Pendiente') >= 0, 'col ' + COL.CONSULTA_PSICOLOGA + ' incluye Pendiente')

  console.log('  → Validaciones actualizadas correctamente ✓')
}

// ─── TEST 3: _COLUMNAS DESCRIPCIONES CORRECTAS ─────────────────────────────

function testColumnasDescripciones() {
  _logTitulo('Constantes: descripciones de _COLUMNAS')

  _assertEq(_COLUMNAS[COL.CCV].name, 'CCV VIGENTE CUIDADOR', 'col ' + COL.CCV + ' name')
  _assertEq(_COLUMNAS[COL.MORBILIDAD].name, 'MORBILIDAD', 'col ' + COL.MORBILIDAD + ' name')
  _assertEq(_COLUMNAS[COL.CONTROLES_MISCELANEOS].name, 'CONTROLES MISCELÁNEOS', 'col ' + COL.CONTROLES_MISCELANEOS + ' name')
  _assertEq(_COLUMNAS[COL.CONSULTA_PSICOLOGA].name, 'CONSULTA PSICOLOGA', 'col ' + COL.CONSULTA_PSICOLOGA + ' name')
  _assertEq(_COLUMNAS[COL.ZONA_EVACUACION].name, 'ZONA EVACUACION', 'col ' + COL.ZONA_EVACUACION + ' name')
  _assertEq(_COLUMNAS[COL.OBSERVACIONES].name, 'OBSERVACIONES', 'col ' + COL.OBSERVACIONES + ' name')
  _assertEq(_COLUMNAS[COL.EDITOR].name, 'EDITOR', 'col ' + COL.EDITOR + ' name')

  console.log('  → Descripciones correctas ✓')
}

// ─── TEST 4: FORM_SECCIONES Y FORM_A_PAC ────────────────────────────────────

function testFormConstantes() {
  _logTitulo('Constantes: Formulario')

  _assertEq(FORM_SECCIONES.length, 4, 'FORM_SECCIONES tiene 4 secciones')

  var secciones = FORM_SECCIONES.map(function(s) { return s.nombre })
  _assert(secciones.indexOf('GESTIÓN') >= 0, 'contiene GESTIÓN')
  _assert(secciones.indexOf('REGISTRO') >= 0, 'contiene REGISTRO')
  _assert(secciones.indexOf('ATENCIÓN') >= 0, 'contiene ATENCIÓN')
  _assert(secciones.indexOf('CIERRE') >= 0, 'contiene CIERRE')

  _assertEq(FORM_A_PAC.length, 9, 'FORM_A_PAC tiene 9 mapeos')
  _assertEq(FORM_A_PAC[0][0], 6, 'FORM_A_PAC[0] col origen = 6 (FECHA ATENCIÓN)')
  _assertEq(FORM_A_PAC[1][0], 7, 'FORM_A_PAC[1] col origen = 7 (RUT USUARIO)')
  _assertEq(FORM_A_PAC[1][1], 8, 'FORM_A_PAC[1] col destino = 8 (RUN)')

  console.log('  → Formulario constantes correctas ✓')
}

// ─── TEST 5: HOJA PACIENTES NO TIENE DATOS MEZCLADOS ───────────────────────

function testPacientesDatosIntactos() {
  _logTitulo('Pacientes: datos intactos (solo lectura)')

  var sh = _getPacientes()
  var lr = sh.getLastRow()

  if (lr < 4) {
    console.log('  ⚠️  No hay datos de pacientes suficientes (lr=' + lr + '). Saltando.')
    return
  }

  // Leer primeras 3 filas de datos (4, 5, 6)
  var data = sh.getRange(4, 1, Math.min(3, lr - 3), 6).getValues()

  for (var r = 0; r < data.length; r++) {
    var sector = String(data[r][1] || '').trim()
    var estado = String(data[r][5] || '').trim()
    var nombre = String(data[r][2] || '').trim()

    // SECTOR debe ser VERDE/AMARILLO/NARANJO/Pendiente o vacío
    _assert(
      ['', 'VERDE', 'AMARILLO', 'NARANJO', 'PENDIENTE'].indexOf(sector.toUpperCase()) >= 0,
      'Fila ' + (r + 4) + ': SECTOR no debe contener nombre (' + sector + ')'
    )
    // ESTADO debe ser VIGENTE/FALLECIDO/etc, no un nombre
    _assert(
      ['', 'VIGENTE', 'FALLECIDO', 'EGRESO', 'SUSPENDIDO', 'PENDIENTE'].indexOf(estado.toUpperCase()) >= 0,
      'Fila ' + (r + 4) + ': ESTADO no debe contener nombre (' + estado + ')'
    )
  }

  console.log('  → Datos de pacientes intactos (sin mezcla) ✓')
}

// ─── TEST 6: CONTROL_COLS ACTUALIZADOS ─────────────────────────────────────

function testControlColsActualizados() {
  _logTitulo('Constantes: CONTROL_COLS mapeos correctos')

  var expected = {
    'EXAMENES CUIDADOR':    [25, 26],
    'EMPA/EMPAM CUIDADOR':  [21, 22],
    'PIC-1C':               [32, 33],
    'PIC-2C':               [34, 35],
    'EXAMENES USUARIO':     [55, 56],
    'CONTROL MEDICO':       [58, 59],
    'CCV MEDICO':           [61, 62],
    'CSCV ENFERMERIA':      [63, 64],
    'PODOLOGO':             [66, 67],
    'NUTRICIONISTA':        [68, 69],
    'FONOAUDIOLOGA':        [71, 72],
    'CONTROL KINESICO':     [73, 74],
    'ODONTOLOGIA':          [75, 76],
    'EMPA/EMPAM USUARIO':   [82, 83],
  }

  for (var i = 0; i < CONTROL_COLS.length; i++) {
    var name = CONTROL_COLS[i][0]
    var dc = CONTROL_COLS[i][1]
    var sc = CONTROL_COLS[i][2]
    var exp = expected[name]
    _assert(!!exp, name + ' debe existir en CONTROL_COLS')
    _assertEq(dc, exp[0], name + ' col fecha = ' + exp[0])
    _assertEq(sc, exp[1], name + ' col estado = ' + exp[1])
  }

  console.log('  → Todos los CONTROL_COLS mapeados correctamente ✓')
}

// ─── TEST 7: _agregarColumnasFaltantes ─────────────────────────────────────

function testAgregarColumnasFaltantesExiste() {
  _logTitulo('_agregarColumnasFaltantes: función existe')

  _assert(typeof _agregarColumnasFaltantes === 'function', '_agregarColumnasFaltantes debe ser una función')

  var src = _agregarColumnasFaltantes.toString()
  _assert(src.indexOf('insertColumns') > 0, 'debe llamar insertColumns')
  _assert(src.indexOf('getLastColumn') > 0, 'debe obtener última columna actual')
  _assert(src.indexOf('_COLUMNAS') > 0, 'debe usar _COLUMNAS para nombres')

  console.log('  → _agregarColumnasFaltantes implementada correctamente ✓')
}

// ─── TEST 7: repintarPacientes llama a _agregarColumnasFaltantes ──────────

function testRepintarLlamaAgregarColumnas() {
  _logTitulo('repintarPacientes: llama a _agregarColumnasFaltantes')

  var src = repintarPacientes.toString()
  _assert(src.indexOf('_agregarColumnasFaltantes') > 0, 'repintarPacientes debe llamar _agregarColumnasFaltantes')

  console.log('  → repintarPacientes integra _agregarColumnasFaltantes ✓')
}

// ─── EJECUTAR TODAS ─────────────────────────────────────────────────────────

function ejecutarPruebas() {
  var errors = [], passed = 0, failed = 0

  var tests = [
    testColumnasCriticas,
    testValidacionesActualizadas,
    testColumnasDescripciones,
    testFormConstantes,
    testPacientesDatosIntactos,
    testControlColsActualizados,
    testAgregarColumnasFaltantesExiste,
    testRepintarLlamaAgregarColumnas,
  ]

  for (var i = 0; i < tests.length; i++) {
    try {
      tests[i]()
      passed++
    } catch(e) {
      errors.push(tests[i].name + ': ' + (e.message || e).toString().substring(0, 120))
      failed++
    }
  }

  if (failed === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('🧪 Todas pasaron (' + passed + '/' + (passed + failed) + ')', '', 5)
  } else {
    var msg = '⚠️ ' + failed + ' fallaron:\n' + errors.join('\n')
    SpreadsheetApp.getActiveSpreadsheet().toast(msg.substring(0, 180), 'Pruebas', 15)
  }
}
