// ─── CONSTANTES GLOBALES ─────────────────────────────────────────────────────

// PACIENTES: secciones, anchos, validaciones
var PAC_SECCIONES = [
  { nombre: 'IDENTIFICACIÓN',             ini: 1,  fin: 15, bg: '#2E7D5B', fg: '#ffffff' },
  { nombre: 'CUIDADOR PRINCIPAL',         ini: 16, fin: 28, bg: '#B5651D', fg: '#ffffff' },
  { nombre: 'CLASIFICACIÓN / DEPENDENCIA',ini: 29, fin: 36, bg: '#5C4EE5', fg: '#ffffff' },
  { nombre: 'PATOLOGÍAS CRÓNICAS',        ini: 37, fin: 53, bg: '#8E44AD', fg: '#ffffff' },
  { nombre: 'CONTROLES Y SEGUIMIENTO',    ini: 54, fin: 80, bg: '#1F6FB2', fg: '#ffffff' },
  { nombre: 'SOCIAL / PSICOLÓGICO',       ini: 81, fin: 96, bg: '#117864', fg: '#ffffff' },
  { nombre: 'INMUNIZACIÓN',               ini: 97, fin: 100,bg: '#4A86E8', fg: '#ffffff', bg2: '#117A65' },
  { nombre: 'CAPACITACIONES CUIDADOR',    ini: 101,fin: 108,bg: '#6E2C00', fg: '#ffffff' },
  { nombre: 'GESTIÓN',                    ini: 109,fin: 111,bg: '#922B21', fg: '#ffffff' },
]

var PAC_ANCHOS = [
  40, 80, 180, 140, 140, 90, 45, 110, 100, 40,
  220, 100, 110, 100, 110, 200, 45, 110, 40,
  120, 120, 120, 180, 90, 110, 120, 120, 110,
  100, 70, 70, 110, 90, 110, 90, 50, 45, 45, 45, 45, 45,
  55, 45, 45, 55, 60, 50, 45, 60, 60, 55,
  70, 300, 60, 110, 120, 100, 110, 120, 100, 110, 120,
  120, 90, 110, 120, 110, 120, 120, 110, 120, 120,
  120, 120, 110, 120, 100, 50, 110, 120, 110, 250,
  120, 110, 160, 50, 50, 120, 120, 50, 120, 80, 140,
  50, 90, 90, 90, 90, 110, 110, 90, 120, 120,
  100, 100, 120, 130, 250, 250,
]

var PAC_SI_NO = ['SI', 'NO', 'N/A', 'Pend.']
var PAC_SI_NO_PROB = ['SI', 'NO', 'N/A', 'Prob. Domicilio', 'Rechaza']

var PAC_VALIDACIONES = {
  2: ['VERDE', 'AMARILLO', 'NARANJO', 'PENDIENTE'],
  6: ['VIGENTE', 'FALLECIDO', 'EGRESO', 'SUSPENDIDO', 'ALTA', 'TRASLADO', 'PENDIENTE'],
  7: ['F', 'M', 'Pend.'],
  17: ['F', 'M', 'Pend.'],
  20: ['EMPA', 'EMPAM', 'N/A', 'PENDIENTE'],
  22: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  24: PAC_SI_NO_PROB,
  26: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  27: ['REVISADO', 'PENDIENTE', 'N/A'],
  29: ['SEVERO', 'MODERADO', 'LEVE', 'INDEPENDIENTE', 'N/A'],
  30: PAC_SI_NO, 31: PAC_SI_NO,
  33: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  35: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  36: ['SI', 'NO', 'N/A', 'PENDIENTE'],
  37: PAC_SI_NO, 38: PAC_SI_NO, 39: PAC_SI_NO, 40: PAC_SI_NO,
  41: PAC_SI_NO, 42: PAC_SI_NO, 43: PAC_SI_NO, 44: PAC_SI_NO,
  45: PAC_SI_NO, 46: PAC_SI_NO, 47: PAC_SI_NO, 48: PAC_SI_NO,
  49: PAC_SI_NO, 50: PAC_SI_NO, 51: PAC_SI_NO, 52: PAC_SI_NO,
  54: PAC_SI_NO_PROB,
  56: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  58: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  59: PAC_SI_NO,
  61: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  63: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  64: PAC_SI_NO,
  66: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  68: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  69: ['NORMAL', 'SOBREPESO', 'OBESIDAD', 'BAJO PESO', 'N/A'],
  71: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  73: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  75: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  77: PAC_SI_NO,
  81: ['EMPA', 'EMPAM', 'N/A', 'PENDIENTE'],
  83: ['AL DIA', 'POR VENCER', 'VENCIDO', 'PENDIENTE', 'N/A'],
  84: ['BENEFICIARIO', 'INGRESADO', 'NO INGRESA', 'NO APLICA', 'PENDIENTE', 'EN ESPERA', 'RECHAZA', 'N/A'],
  86: ['N/A', 'PENDIENTE', 'RECHAZA', 'SIN SOBRECARGA', 'SOBRECARGA LEVE',
    'SOBRECARGA INTENSA', 'CUIDADORA REMUNERADA', 'AUSENCIA'],
  87: ['SI', 'NO', 'Pendiente', 'N/A'],
  88: ['SI', 'NO', 'N/A', 'Pend.'],
  89: PAC_SI_NO, 90: PAC_SI_NO,
  91: ['G', 'M', 'XG', 'N/A'],
  92: ['Entregado', 'No E. / Pendiente', 'N/A'],
  94: ['SI', 'NO', 'N/A', 'PENDIENTE', 'BASTON', 'ANDADOR',
    'SILLA RUEDAS', 'MULETAS'],
  95: PAC_SI_NO, 96: PAC_SI_NO,
  97: PAC_SI_NO, 98: PAC_SI_NO, 99: PAC_SI_NO, 100: PAC_SI_NO,
  109: ['URGENTE', 'POR REVISAR', 'AL DIA', 'N/A'],
}

var COL = {
  ID: 1, SECTOR: 2, NOMBRE: 3, APELLIDO: 4, APELLIDO2: 5, RUN: 8, VITAL: 6, SEXO: 7,
  EDAD_USUARIO: 10, TELEFONO: 12, CUIDADOR: 16,
  SEXO_CUIDADOR: 17, RUN_CUIDADOR: 18, EDAD_CUIDADOR: 19, EMPA_CUIDADOR: 20, EMPA_CUIDADOR_FECHA: 21,
  EMPA_USUARIO: 81, EMPA_USUARIO_FECHA: 82,
  CAPACITACIONES_INI: 101, CAPACITACIONES_FIN: 108,
  PRIORIDAD: 109, OBSERVACIONES: 110, EDITOR: 111,
  CCV: 28, CONTROLES_MISCELANEOS: 79, F_ALTA_LPP: 80, ZARIT: 85,
}

// ─── COLORES COMPARTIDOS ───────────────────────────────────────────────────

var _SECTOR_COLORS = {
  VERDE:    ['#c8e6c9', '#2e7d32'],
  AMARILLO: ['#fff9c4', '#f9a825'],
  NARANJO:  ['#ffe0b2', '#e65100'],
  PENDIENTE:['#f5f5f5', '#999999'],
}

var _ESTADO_COLORS = {
  VIGENTE:   ['#c8e6c9', '#2e7d32'],
  FALLECIDO: ['#f5f5f5', '#aaaaaa'],
  SUSPENDIDO:['#fff9c4', '#888888'],
  EGRESO:    ['#f5f5f5', '#888888'],
  ALTA:      ['#bbdefb', '#1565c0'],
  PENDIENTE: ['#fff9c4', '#f9a825'],
}

var TEXT_UPPER = [2, 3, 4, 5, 11, 16]

var CONTROL_COLS = [
  ['EXAMENES CUIDADOR',     25, 26],
  ['EMPA/EMPAM CUIDADOR',   21, 22],
  ['PIC-1C',                32, 33],
  ['PIC-2C',                34, 35],
  ['EXAMENES USUARIO',      55, 56],
  ['CONTROL MEDICO',        58, 59],
  ['CCV MEDICO',            61, 62],
  ['CSCV ENFERMERIA',       63, 64],
  ['PODOLOGO',              66, 67],
  ['NUTRICIONISTA',         68, 69],
  ['FONOAUDIOLOGA',         71, 72],
  ['CONTROL KINESICO',      73, 74],
  ['ODONTOLOGIA',           75, 76],
  ['EMPA/EMPAM USUARIO',    82, 83],
]

var COL_TO_CONTROL = {}
var COL_TO_STATUS = {}
for (var _ci = 0; _ci < CONTROL_COLS.length; _ci++) {
  COL_TO_CONTROL[CONTROL_COLS[_ci][1]] = CONTROL_COLS[_ci]
  if (CONTROL_COLS[_ci][2]) COL_TO_STATUS[CONTROL_COLS[_ci][2]] = CONTROL_COLS[_ci]
}

var _RELEVANT_COLS = {}
for (var _ri = 0; _ri < TEXT_UPPER.length; _ri++) _RELEVANT_COLS[TEXT_UPPER[_ri]] = true
_RELEVANT_COLS[COL.RUN] = true
_RELEVANT_COLS[COL.RUN_CUIDADOR] = true
_RELEVANT_COLS[COL.OBSERVACIONES] = true
_RELEVANT_COLS[COL.EDITOR] = true
_RELEVANT_COLS[COL.EDAD_USUARIO] = true
_RELEVANT_COLS[COL.EDAD_CUIDADOR] = true
_RELEVANT_COLS[COL.VITAL] = true
for (var _ri = 0; _ri < CONTROL_COLS.length; _ri++) {
  _RELEVANT_COLS[CONTROL_COLS[_ri][1]] = true
  if (CONTROL_COLS[_ri][2]) _RELEVANT_COLS[CONTROL_COLS[_ri][2]] = true
}
for (var _ri = 101; _ri <= 108; _ri++) _RELEVANT_COLS[_ri] = true
_RELEVANT_COLS[COL.CCV] = true
_RELEVANT_COLS[COL.ZARIT] = true
_RELEVANT_COLS[87] = true
_RELEVANT_COLS[88] = true

// ─── GUÍA DE COLUMNAS (para tooltips + instrucciones) ─────────────────────────
// Array es 1-based: _COLUMNAS[n] = datos de columna n.

var _COLUMNAS = (function() {
  var C = new Array(112), _cnt = 0
  function s(n, name, desc, vals, auto) { C[n] = { n: n, name: name, desc: desc, vals: vals || '', auto: auto || '' }; _cnt = Math.max(_cnt, n) }

  // ── IDENTIFICACIÓN (1-15) ──
  s(1,  'N°',                  'Número correlativo de fila', '', 'Automático')
  s(2,  'SECTOR',              'Sector geográfico asignado al paciente', 'VERDE · AMARILLO · NARANJO · Pendiente', 'Dropdown')
  s(3,  'NOMBRE',              'Nombre del paciente (completo)', '', 'Mayúscula automática')
  s(4,  'AP. PATERNO',         'Primer apellido del paciente', '', 'Mayúscula automática')
  s(5,  'AP. MATERNO',         'Segundo apellido del paciente', '', 'Mayúscula automática')
  s(6,  'ESTADO',              'Estado de vigencia del registro', 'VIGENTE · FALLECIDO · EGRESO · SUSPENDIDO · TRASLADO · PENDIENTE', 'Dropdown')
  s(7,  'SEXO',                'Sexo del paciente', 'F · M · Pend.', 'Dropdown')
  s(8,  'RUN',                 'RUT del paciente (con guión, sin puntos)', '', 'Formato automático')
  s(9,  'F. NACIMIENTO',       'Fecha de nacimiento del paciente', '', 'dd/mm/aaaa')
  s(10, 'EDAD',                'Edad en años del paciente', '', 'Automático desde F. NACIMIENTO')
  s(11, 'DIRECCION',           'Dirección del domicilio del paciente', '', 'Mayúscula automática')
  s(12, 'TELEFONO',            'Teléfono(s) de contacto del paciente', '', 'Formato automático')
  s(13, 'F. INGRESO PADI',     'Fecha de ingreso al programa PADDS', '', 'dd/mm/aaaa')
  s(14, 'F. EGRESO',           'Fecha de egreso del programa (si aplica)', '', 'dd/mm/aaaa')
  s(15, 'F. FALLECIMIENTO',    'Fecha de fallecimiento del paciente (si aplica)', '', 'dd/mm/aaaa')

  // ── CUIDADOR PRINCIPAL (16-28) ──
  s(16, 'NOMBRE CUIDADOR',     'Nombre completo del cuidador principal', '', 'Mayúscula automática')
  s(17, 'SEXO C.',             'Sexo del cuidador', 'F · M · Pend.', 'Dropdown')
  s(18, 'RUN C.',              'RUT del cuidador (con guión, sin puntos)', '', 'Formato automático')
  s(19, 'EDAD C.',             'Edad del cuidador en años', '', 'Automático o manual')
  s(20, 'F. EMPA/EMPAM CUIDADOR', 'Tipo EMPA/EMPAM asignado al cuidador según edad', 'EMPA · EMPAM · N/A · PENDIENTE', 'Automático según edad')
  s(21, 'F. EMPA/EMPAM',       'Fecha del último EMPA/EMPAM del cuidador', '', 'Control: col 21 → 22')
  s(22, 'ESTADO EMPA/EMPAM CUIDADOR', 'Vigencia del EMPA/EMPAM del cuidador', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(23, 'PATOLOGIAS CUIDADOR', 'Patologías crónicas del cuidador (anotación)', '', 'Texto libre')
  s(24, 'PSCV CUIDADOR',       'PSCV del cuidador', 'SI · NO · N/A · Prob. Domicilio · Rechaza', 'Dropdown')
  s(25, 'F. EXAMENES CUIDADOR','Fecha de exámenes del cuidador', '', 'Control: col 25 → 26')
  s(26, 'ESTADO EXAMENES CUIDADOR', 'Vigencia de exámenes del cuidador', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(27, 'EXS REVISADOS CUIDADOR', 'Exámenes revisados del cuidador', 'REVISADO · PENDIENTE · N/A', 'Dropdown')
  s(28, 'CCV VIGENTE CUIDADOR', 'Fecha de la última CCV (cardiopatía descompensada) vigente del cuidador', 'Fecha o N/A', 'Color por vigencia (verde: al día · rojo: vencido)')

  // ── CLASIFICACIÓN / DEPENDENCIA (29-36) ──
  s(29, 'BARTHEL',             'Índice Barthel — dependencia física', 'SEVERO · MODERADO · LEVE · INDEPENDIENTE · N/A', 'Dropdown')
  s(30, 'ONCOLOGICO',          'Paciente oncológico', 'SI · NO · N/A · Pend.', 'Dropdown')
  s(31, 'CARDEX',              'Cardiopatía descompensada / riesgo cardiovascular', 'SI · NO · N/A · Pend.', 'Dropdown')
  s(32, 'FECHA PIC-1C',        'Fecha del Plan de Intervención Individual 1 componente', '', 'Control: col 32 → 33')
  s(33, 'ESTADO PIC-1C',       'Vigencia del PIC-1C', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(34, 'FECHA PIC-2C',        'Fecha del Plan de Intervención Individual 2 componentes', '', 'Control: col 34 → 35')
  s(35, 'ESTADO PIC-2C',       'Vigencia del PIC-2C', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(36, 'ELECTRODEPENDENCIA',  'Paciente electrodependiente (requiere suministro eléctrico constante para equipos de soporte vital)', 'SI · NO · N/A · PENDIENTE', 'Dropdown')

  // ── PATOLOGÍAS CRÓNICAS (37-53) ──
  var PATS = ['HTA','DLP','DM','ACV','IAM','GLAUCOMA','ASMA','EPOC','EPILEPSIA','PARKINSON','ARTROSIS','HPT','DEPRESION','ALZHEIMER','DEMENCIA','ESQUIZOFRENIA','OTRAS PATOLOGIAS']
  for (var i = 0; i < PATS.length; i++) {
    var d = PATS[i] === 'OTRAS PATOLOGIAS' ? 'Otras patologías no listadas' : 'Diagnóstico de ' + PATS[i]
    var v = PATS[i] === 'OTRAS PATOLOGIAS' ? 'Texto libre' : 'SI · NO · N/A · Pend.'
    var a = PATS[i] === 'OTRAS PATOLOGIAS' ? 'Texto libre' : 'Dropdown'
    s(37 + i, PATS[i], d, v, a)
  }

  // ── CONTROLES Y SEGUIMIENTO (54-80) ──
  s(54, 'PSCV',                'PSCV del usuario', 'SI · NO · N/A · Prob. Domicilio · Rechaza', 'Dropdown')
  s(55, 'F. EXAMENES USUARIO', 'Fecha de exámenes del usuario', '', 'Control: col 55 → 56')
  s(56, 'ESTADO EXAMENES USUARIO', 'Vigencia de exámenes del usuario', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(57, 'MORBILIDAD',          'Comorbilidades y diagnósticos relevantes del usuario (texto libre)', '', 'Texto libre')
  s(58, 'F. CONTROL MEDICO',   'Fecha del último control médico', '', 'Control: col 58 → 59')
  s(59, 'ESTADO CONTROL MEDICO', 'Vigencia del control médico', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(60, 'RECETAS CONTROLADAS', 'Recetas controladas (narcóticos/psicotrópicos)', 'SI · NO · N/A · Pend.', 'Dropdown')
  s(61, 'F. CCV MEDICO',       'Fecha control cardiovascular (médico)', '', 'Control: col 61 → 62')
  s(62, 'ESTADO CCV MEDICO',   'Vigencia CCV médico', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(63, 'F. CSCV ENFERMERIA',  'Fecha control cardiovascular (enfermería)', '', 'Control: col 63 → 64')
  s(64, 'ESTADO CSCV ENFERMERIA', 'Vigencia CSCV enfermería', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(65, 'RIESGO PIE DM',       'Evaluación de riesgo de pie diabético', 'SI · NO · N/A · Pend.', 'Dropdown')
  s(66, 'F. PODOLOGO',         'Fecha de atención de podología', '', 'Control: col 66 → 67')
  s(67, 'ESTADO PODOLOGO',     'Vigencia de podología', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(68, 'F. NUTRICIONISTA',    'Fecha del control nutricional', '', 'Control: col 68 → 69')
  s(69, 'ESTADO NUTRICIONISTA','Vigencia del control nutricional', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(70, 'ESTADO NUTRICIONAL',  'Clasificación del estado nutricional', 'NORMAL · SOBREPESO · OBESIDAD · BAJO PESO · N/A', 'Dropdown')
  s(71, 'F. FONOAUDIOLOGA',    'Fecha de atención de fonoaudiología', '', 'Control: col 71 → 72')
  s(72, 'ESTADO FONOAUDIOLOGA','Vigencia de fonoaudiología', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(73, 'F. CONTROL KINESICO', 'Fecha del control kinésico', '', 'Control: col 73 → 74')
  s(74, 'ESTADO KINESICO',     'Vigencia del control kinésico', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(75, 'FECHA ODONTOLOGIA',   'Fecha de atención odontológica', '', 'Control: col 75 → 76')
  s(76, 'ESTADO ODONTOLOGIA',  'Vigencia de odontología', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(77, 'TTO INVASIVOS',       'Tratamientos invasivos realizados (ej: CISTOSTOMÍA, SNAP, OSTOMÍA, sonda, etc.)', '', 'Texto libre')
  s(78, 'LPP',                 'Lesiones por Presión / cuidados de piel', '', 'Dropdown')
  s(79, 'CONTROLES MISCELÁNEOS', 'Registro acumulativo de controles misceláneos. Cada vez que se recibe un dato se agrega: fecha + prestación. Ej: 15/06/2026 - Control glucosa', '', 'Texto libre acumulativo')
  s(80, 'F. ALTA LPP',         'Fecha de alta de lesión por presión', '', 'dd/mm/aaaa')

  // ── SOCIAL / PSICOLÓGICO (81-96) ──
  s(81, 'EMPA/EMPAM USUARIO',  'Tipo EMPA/EMPAM asignado al usuario según edad', 'EMPA · EMPAM · N/A · PENDIENTE', 'Automático según edad')
  s(82, 'F. EMPA/EMPAM USUARIO', 'Fecha del último EMPA/EMPAM del usuario', '', 'Control: col 82 → 83')
  s(83, 'ESTADO EMPA/EMPAM',   'Vigencia del EMPA/EMPAM del usuario', 'AL DIA · POR VENCER · VENCIDO · PENDIENTE · N/A', 'Automático')
  s(84, 'ESTIPENDIO',          'Tipo de estipendio o beneficio social', 'BENEFICIARIO · INGRESADO · PENDIENTE · EN ESPERA · RECHAZA · N/A', 'Dropdown')
  s(85, 'ZARIT',               'Fecha de aplicación escala Zarit', 'Fecha o N/A', 'Color por vigencia')
  // ESTADO ZARIT (col 86) eliminado — innecesario, ZARIT se colorea por fecha directamente
  s(86, 'RESULTADO ZARIT',     'Resultado escala Zarit (sobrecarga cuidador)', 'N/A · PENDIENTE · RECHAZA · SIN SOBRECARGA · SOBRECARGA LEVE · SOBRECARGA INTENSA · CUID. REMUNERADA · AUSENCIA', 'Dropdown')
  s(87, 'CONSULTA PSICOLOGA',  'Consulta psicológica realizada', 'SI · NO · Pendiente · N/A', 'Dropdown')
  s(88, 'CONSULTA TRABAJADORA SOCIAL', 'Consulta con trabajadora social — fecha de atención', 'SI · NO · N/A · Pend.', 'Dropdown')
  s(89, 'SIGGES',              'Registro en SIGGES', 'SI · NO · N/A · Pend.', 'Dropdown')
  s(90, 'PAÑALES',             'Entrega de pañales', 'SI · NO · N/A · Pend.', 'Dropdown')
  s(91, 'TALLA',               'Talla de pañales entregados', 'G · M · XG · N/A', 'Dropdown')
  s(92, 'ENTREGA PAÑALES',     'Fecha o estado de entrega de pañales', 'Entregado · No E. / Pendiente · N/A', 'Dropdown')
  s(93, 'ZONA EVACUACION',     'Zona de evacuación asignada', '', 'Texto libre')
  s(94, 'AYUDAS TECNICAS',     'Ayudas técnicas requeridas/entregadas', 'SI · NO · N/A · PENDIENTE · BASTÓN · ANDADOR · SILLA RUEDAS · MULETAS', 'Dropdown')
  s(95, 'IVADEC',              'Índice de Valoración de Dependencia (IVADEC)', 'SI · NO · N/A · Pend.', 'Dropdown')
  s(96, 'FICHA FAMILIAR',      'Ficha familiar realizada', 'SI · NO · N/A · Pend.', 'Dropdown')

  // ── INMUNIZACIÓN (97-100) ──
  s(97, 'INFLUENZA USUARIO',   'Vacuna influenza del usuario', 'SI · NO · N/A · Pend.', 'Dropdown')
  s(98, 'NEUMO23 USUARIO',     'Vacuna antineumocócica (Neumo23) del usuario', 'SI · NO · N/A · Pend.', 'Dropdown')
  s(99, 'INFLUENZA CUIDADOR',  'Vacuna influenza del cuidador', 'SI · NO · N/A · Pend.', 'Dropdown')
  s(100, 'NEUMO23 CUIDADOR',   'Vacuna antineumocócica (Neumo23) del cuidador', 'SI · NO · N/A · Pend.', 'Dropdown')

  // ── CAPACITACIONES CUIDADOR (101-108) ──
  var CAPS = ['CAP. KINESIOLOGIA','CAP. ENFERMERIA','CAP. TENS','CAP. TENS ESTIVAL','CAP. TRABAJO SOCIAL','CAP. NUTRICION','CAP. PSICOLOGIA','CAP. FONOAUDIOLOGIA']
  for (var i = 0; i < CAPS.length; i++) s(101 + i, CAPS[i], 'Capacitación: ' + CAPS[i] + ' — fecha de realización', 'Fecha o N/A', 'Color por vigencia')

  // ── GESTIÓN (109-111) ──
  s(109, 'PRIORIDAD GENERAL',  'Prioridad calculada según el peor estado de todos los controles', 'URGENTE · POR REVISAR · AL DIA · N/A', 'Automático')
  s(110, 'OBSERVACIONES',      'Próximo control o visita + observaciones de la última atención. Formato libre.', '', 'Texto libre')
  s(111, 'EDITOR',             'Correo del último usuario que editó la fila', '', 'Automático')

  C._count = _cnt
  return C
})()

// Auto-generar COL.* faltantes desde _COLUMNAS (nombre → mayúsculas sin acentos)
for (var _ci = 0; _ci < _COLUMNAS.length; _ci++) {
  var _c = _COLUMNAS[_ci]
  if (!_c) continue
  var _key = _c.name.replace(/[^A-Z0-9_]/gi, '_').replace(/_+/g, '_').replace(/^_|_$/g, '').toUpperCase()
  if (COL[_key] === undefined) COL[_key] = _ci
}

// FORMULARIO (staging/recepción)

var FORM_SECCIONES = [
  { nombre: 'GESTIÓN',     ini: 1,  fin: 5,  bg: '#2c3e50', fg: '#ffffff' },
  { nombre: 'REGISTRO',    ini: 6,  fin: 9,  bg: '#1565C0', fg: '#ffffff' },
  { nombre: 'ATENCIÓN',    ini: 10, fin: 16, bg: '#2E7D5B', fg: '#ffffff' },
  { nombre: 'CIERRE',      ini: 17, fin: 17, bg: '#6A1B9A', fg: '#ffffff' },
]

var FORM_ESTADOS = ['Pendiente', 'Aprobado', 'Rechazado']

var FORM_A_PAC = [
  [6, null],   // FECHA ATENCIÓN    → service date (set via _SERVICIO_COL_MAP)
  [7, 8],      // RUT USUARIO       → COL.RUN
  [8, 3],      // NOMBRE            → COL.NOMBRE
  [9, 4],      // APELLIDO          → COL.APELLIDO (paterno)
  [10, 111],   // PROFESIONAL       → COL.EDITOR
  [12, 70],    // ESTADO NUTRICIONAL → col 70 (ESTADO NUTRICIONAL)
  [14, 29],    // RESULTADO BARTHEL  → col 29 (BARTHEL)
  [15, 86],    // RESULTADO ZARIT    → col 86 (RESULTADO ZARIT)
  [17, 110],   // OBSERVACIONES     → COL.OBSERVACIONES
]

var _FORM_FIELD_MAP = {
  'FECHA': 6,
  'RUT USUARIO': 7,
  'NOMBRE': 8,
  'APELLIDO': 9,
  'PROFESIONAL': 10,
  'PRESTACION': 11,
  'ESTADO NUTRICIONAL (SOLO NUTRIOLOGA)': 12,
  'VISITA PERDIDA': 13,
  'RESULTADO BARTHEL': 14,
  'RESULTADO ZARIT': 15,
  'PROXIMO CONTROL': 16,
  'OBSERVACIONES': 17,
}

var _SERVICIO_COL_MAP = {
  'EXAMEN DE SANGRE USUARIO': 55,
  'TOMA DE EXAMENES': 55,
  'EXAMEN DE SANGRE CUIDADOR': 25,
  'MORBILIDAD': 57,
  'CONTROL MEDICO': 58,
  'CONTROL DE SALUD (INCLUYE CSV)': 58,
  'PRESION ARTERIAL': 63,
  'CCV MEDICO': 61,
  'CCV ENFERMERIA': 63,
  'CCV NUTRICIONAL': 68,
  'CSCV ENFERMERIA': 63,
  'PODOLOGO': 66,
  'ATENCION PODOLOGICA': 66,
  'NUTRICIONISTA': 68,
  'CONSULTA NUTRICIONAL': 68,
  'FONOAUDIOLOGA': 71,
  'ATENCION FONOAUDIOLOGA': 71,
  'CONTROL KINESICO': 73,
  'CONSULTA KINE': 73,
  'CONSULTA KINESIOLOGO': 73,
  'ODONTOLOGIA': 75,
  'EMPA': 82,
  'EMPAM': 82,
  'PIC-1C': 32,
  'VISITA PIC 1': 32,
  'PIC-2C': 34,
  'VISITA PIC 2': 34,
  'EMPA/EMPAM CUIDADOR': 21,
  'CCV CUIDADOR': 28,
  'ZARIT': 85,
  'APLICACION ZARIT': 85,
  'CAMBIO O INSTALACION DE SONDA': 77,
  'CONSULTA TRABAJADORA SOCIAL': 88,
  'CONSULTA PSICOLOGICA': 87,
  'CAPACITACION KINESIOLOGA': 101,
  'CAPACITACION ENFERMERIA': 102,
  'CAPACITACION TENS': 103,
  'CAPACITACION TRABAJADORA SOCIAL': 105,
  'CAPACITACION PSICOLOGA': 107,
  'CAPACITACION FONOAUDIOLOGA': 108,
}

var _SERVICIO_SI_MAP = {
  'INMUNIZACION INFLUENZA USUARIO': 97,
  'INMUNIZACION NEUMO 23 USUARIO': 98,
  'INMUNIZACION INFLUENZA CUIDADOR': 99,
  'INMUNIZACION NEUMO 23 CUIDADOR': 100,
}

var _FORM_VAL_MAP = {
  70: {  // ESTADO NUTRICIONAL → col 70
    'SOBRE PESO': 'SOBREPESO',
  },
}

// AGENDA
var HOJA = 'Agenda Profesionales'
var PC = 6
var GC = 1
var GI = [1, 1+PC+GC, 1+2*(PC+GC)]
var GF = 3

var S = {
  header: { bg: '#1a3c5e', fg: '#fff', sz: 12 },
  sub:    { bg: ['#b0bec5','#90caf9','#bbdefb','#bbdefb','#bbdefb','#ffe082'] },
  slots:  {
    abreviadas: { bg: '#fff8e1', fg: '#6d4c00' },
    vdi:        { bg: '#e8f5e9', fg: '#1b5e20' },
    colacion:   { bg: '#fce4ec', fg: '#880e4f' },
    registro:   { bg: '#f3f3f3', fg: '#555' },
  },
}
var ANCHOS = [150, 175, 240, 160, 280, 310]
var SUBS = ['HORARIO','ATENCIÓN','NOMBRE','RUT','DIRECCIÓN','OBSERVACIONES']
var DIAS = ['LUNES','MARTES','MIÉRCOLES','JUEVES','VIERNES']

var SLOTS_DEFAULT = [
  ['8:00','abreviadas'], ['8:15','abreviadas'], ['8:30','abreviadas'], ['8:45','abreviadas'],
  ['9:00','VDI'], ['10:00','VDI'], ['11:00','VDI'], ['12:00','VDI'],
  ['13:00','COLACION'],
  ['14:00','VDI'], ['15:00','VDI'],
  ['16:00','REGISTRO'],
]

var SLOTS = [
  SLOTS_DEFAULT,  // LUNES
  SLOTS_DEFAULT,  // MARTES
  SLOTS_DEFAULT,  // MIÉRCOLES
  SLOTS_DEFAULT,  // JUEVES
  [               // VIERNES — solo hasta las 12:00
    ['8:00','abreviadas'], ['8:15','abreviadas'], ['8:30','abreviadas'], ['8:45','abreviadas'],
    ['9:00','VDI'], ['10:00','VDI'], ['11:00','VDI'], ['12:00','VDI'],
  ],
]

function _calcularRS() {
  var total = 1 // week header
  for (var d = 0; d < 5; d++) {
    total += 3 + (SLOTS[d] || SLOTS_DEFAULT).length  // header + sub-header + slots + separator
  }
  return total
}
