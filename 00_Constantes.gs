// ════════════════════════════════════════════════════════
// ÍNDICE ▏00_Constantes.gs │ constantes globales PADDS

// ─────────────────────────────────────────────────────────
// ─── CONSTANTES GLOBALES ─────────────────────────────────────────────────────

// color indica la vigencia (verde al día / ámbar por vencer / rojo vencido /

var HOJA_PAC = 'Pacientes'
var HOJA_FORM = 'Recepción Formulario Profesional'
var PAC_SECCIONES = [
  { nombre: 'IDENTIFICACIÓN',             ini: 1,  fin: 15, bg: '#1565C0', bg2: '#1E88E5', fg: '#ffffff' },
  { nombre: 'CUIDADOR PRINCIPAL',         ini: 16, fin: 25, bg: '#E65100', bg2: '#FB8C00', fg: '#ffffff' },
  { nombre: 'CLASIFICACIÓN / DEPENDENCIA',ini: 26, fin: 31, bg: '#6A1B9A', bg2: '#8E24AA', fg: '#ffffff' },
  { nombre: 'PATOLOGÍAS CRÓNICAS',        ini: 32, fin: 48, bg: '#C62828', bg2: '#E53935', fg: '#ffffff' },
  { nombre: 'CONTROLES Y SEGUIMIENTO',    ini: 49, fin: 64, bg: '#00695C', bg2: '#00897B', fg: '#ffffff' },
  { nombre: 'SOCIAL / PSICOLÓGICO',       ini: 65, fin: 76, bg: '#AD1457', bg2: '#D81B60', fg: '#ffffff' },
  { nombre: 'PAÑALES',                    ini: 77, fin: 79, bg: '#00838F', bg2: '#00ACC1', fg: '#ffffff' },
  { nombre: 'INMUNIZACIÓN',               ini: 80, fin: 83, bg: '#283593', bg2: '#3949AB', fg: '#ffffff' },
  { nombre: 'CAPACITACIONES CUIDADOR',    ini: 84, fin: 91, bg: '#5D4037', bg2: '#795548', fg: '#ffffff' },
  { nombre: 'SONDA FOLEY',                ini: 92, fin: 95, bg: '#2E7D32', bg2: '#43A047', fg: '#ffffff' },
  { nombre: 'CURACIONES AVANZADAS',       ini: 96, fin: 103,bg: '#D84315', bg2: '#FF7043', fg: '#ffffff' },
  { nombre: 'CONTROL DE SIGNOS VITALES (CSV)', ini: 104,fin: 108,bg: '#B71C1C', bg2: '#D32F2F', fg: '#ffffff' },
  { nombre: 'GESTIÓN',                    ini: 109,fin: 111,bg: '#455A64', bg2: '#78909C', fg: '#ffffff' },
]

var PAC_ANCHOS = [
  40, 80, 170, 130, 130, 80, 45, 110, 100, 40,
  240, 170, 110, 100, 110, 190, 45, 110, 40,
  270, 80, 110, 110, 110, 110,
  100, 70, 70, 100, 100, 105,
  45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45,
  300, 70, 260, 110, 110, 110, 110, 100, 100, 110,
  100, 110, 110, 300, 130,
  100, 110, 110, 100,
  90, 160, 110, 130,
  100, 100, 100, 100,
  160, 70, 110, 100,
  110, 110, 100, 100, 110, 110, 100, 110,
  110, 100, 110, 110, 100, 110, 100, 100,
  100, 60, 110, 60, 110, 100, 110, 110,
  100, 90, 180, 90, 90, 90, 120, 300, 220,
]

var PAC_SI_NO = ['SI', 'NO', 'N/A']
var PAC_SI_NO_PROB = ['SI', 'NO', 'N/A', 'Prob. Domicilio', 'Rechaza']

var _MEDICAMENTOS_LISTA = [
  'CLONAZEPAM', 'CITALOPRAM', 'ALPRAZOLAM', 'DIAZEPAM', 'LORAZEPAM',
  'ESCITALOPRAM', 'SERTRALINA', 'FLUOXETINA', 'VENLAFAXINA', 'QUETIAPINA',
  'RISPERIDONA', 'OLANZAPINA', 'ZOLPIDEM', 'GABAPENTINA', 'PREGABALINA',
  'TRAMADOL', 'MORFINA', 'METILFENIDATO', 'CARBAMAZEPINA', 'VALPROICO',
  'N/A',
]

var _CHECKBOX_COLS = [31]
for (var _cb = 32; _cb <= 47; _cb++) _CHECKBOX_COLS.push(_cb)
_CHECKBOX_COLS.push(77, 92, 96)

var PAC_VALIDACIONES = {
  2: ['VERDE', 'AMARILLO', 'NARANJO', 'PENDIENTE'],
  6: ['VIGENTE', 'FALLECIDO', 'EGRESO', 'EGRESO POR ALTA', 'SUSPENDIDO', 'ALTA', 'TRASLADO', 'PENDIENTE'],
  7: ['F', 'M', 'Pend.'],
  17: ['F', 'M', 'Pend.'],
  21: PAC_SI_NO_PROB,
  22: ['EMPA', 'EMPAM', 'N/A', 'PENDIENTE'],
  26: ['SEVERO', 'MODERADO', 'LEVE', 'INDEPENDIENTE', 'N/A'],
  27: PAC_SI_NO, 28: PAC_SI_NO,
  49: PAC_SI_NO_PROB,
  55: PAC_SI_NO,
  58: ['NORMAL', 'SOBREPESO', 'OBESIDAD', 'BAJO PESO', 'N/A'],
  65: ['EMPA', 'EMPAM', 'N/A', 'PENDIENTE'],
  67: ['BENEFICIARIO', 'INGRESADO', 'NO INGRESA', 'NO APLICA', 'PENDIENTE', 'EN ESPERA', 'RECHAZA', 'N/A'],
  69: ['N/A', 'PENDIENTE', 'RECHAZA', 'SIN SOBRECARGA', 'SOBRECARGA LEVE',
    'SOBRECARGA INTENSA', 'CUIDADORA REMUNERADA', 'AUSENCIA'],
  70: PAC_SI_NO, 71: PAC_SI_NO, 72: PAC_SI_NO, 73: PAC_SI_NO, 74: PAC_SI_NO,
  76: ['SI', 'NO', 'N/A', 'PENDIENTE', 'BASTON', 'ANDADOR',
    'SILLA RUEDAS', 'MULETAS'],
  78: ['G', 'M', 'XG', 'XXG', 'N/A'],
  93: ['VESICAL', 'SUPRAPUBICA', 'N/A'],
  94: ['FR12', 'FR14', 'FR16', 'FR18', 'FR20', 'N/A'],
  98: ['I', 'II', 'III', 'IV', 'N/A'],
  100: ['AVANZADA', 'SIMPLE', 'N/A'],
  63: _MEDICAMENTOS_LISTA,
  109: ['URGENTE', 'POR REVISAR', 'AL DIA', 'N/A'],
}

var _FECHAS_VA = [9, 13, 14, 15, 23, 24, 25, 29, 30, 51, 52, 53, 54, 56, 57, 59, 60,
  61, 64, 66, 68, 79, 84, 85, 86, 87, 88, 89, 90, 91,
  95, 101, 102, 103]

var _VACUNA_COLS = [80, 81, 82, 83]
var _VACUNA_VALS = ['SI', 'NO', 'N/A', 'R', 'P']

var _CONTROL_FECHAS = [
  ['EXAMENES CUIDADOR',       24, 'EXAMENES CUIDADOR'],
  ['EMPA/EMPAM CUIDADOR',     23, 'EMPA/EMPAM CUIDADOR'],
  ['PIC-1C',                  29, 'PIC-1C'],
  ['PIC-2C',                  30, 'PIC-2C'],
  ['EXAMENES USUARIO',        51, 'EXAMENES USUARIO'],
  ['CONTROL MEDICO',          52, 'CONTROL MEDICO'],
  ['CCV MEDICO',              53, 'CCV MEDICO'],
  ['CSCV ENFERMERIA',         54, 'CSCV ENFERMERIA'],
  ['PODOLOGO',                56, 'PODOLOGO'],
  ['NUTRICIONISTA',           57, 'NUTRICIONISTA'],
  ['FONOAUDIOLOGA',           59, 'FONOAUDIOLOGA'],
  ['CONTROL KINESICO',        60, 'CONTROL KINESICO'],
  ['ODONTOLOGIA',             61, 'ODONTOLOGIA'],
  ['EMPA/EMPAM USUARIO',      66, 'EMPA/EMPAM USUARIO'],
  ['RECETAS CONTROLADAS',     64, 'RECETAS CONTROLADAS'],
  ['CCV VIGENTE CUIDADOR',    25, 'CCV VIGENTE CUIDADOR'],
  ['ZARIT',                   68, 'ZARIT'],
  ['SONDA FOLEY',             95, 'SONDA FOLEY'],
  ['CURACIONES',             101, 'CURACIONES'],
  ['PROXIMA CURACION',       102, 'PROXIMA CURACION'],
]

var COL = {
  ID: 1, SECTOR: 2, NOMBRE: 3, APELLIDO: 4, APELLIDO2: 5, VITAL: 6, SEXO: 7,
  RUN: 8, F_NACIMIENTO: 9, EDAD_USUARIO: 10, DIRECCION: 11, TELEFONO: 12,
  F_INGRESO_PADI: 13, F_EGRESO: 14, F_FALLECIMIENTO: 15,
  CUIDADOR: 16, SEXO_CUIDADOR: 17, RUN_CUIDADOR: 18, EDAD_CUIDADOR: 19,
  PATOLOGIAS_CUIDADOR: 20, PSCV_CUIDADOR: 21, EMPA_CUIDADOR: 22,
  F_EMPA_CUIDADOR: 23, F_EXAMENES_CUIDADOR: 24, CCV: 25,
  BARTHEL: 26, ONCOLOGICO: 27, CARDEX: 28, F_PIC1C: 29, F_PIC2C: 30,
  ELECTRODEPENDENCIA: 31,
  PATS_INI: 32, PATS_FIN: 47, OTRAS_PATOLOGIAS: 48,
  PSCV: 49, MORBILIDAD: 50, F_EXAMENES_USUARIO: 51, F_CONTROL_MEDICO: 52,
  F_CCV_MEDICO: 53, F_CSCV_ENFERMERIA: 54, RIESGO_PIE_DM: 55, F_PODOLOGO: 56,
  F_NUTRICIONISTA: 57, ESTADO_NUTRICIONAL: 58, F_FONOAUDIOLOGA: 59,
  F_CONTROL_KINESICO: 60, F_ODONTOLOGIA: 61, CONTROLES_MISCELANEOS: 62,
  RECETAS_CONTROLADAS: 63, F_RECETA: 64,
  EMPA_USUARIO: 65, F_EMPA_USUARIO: 66, ESTIPENDIO: 67, ZARIT: 68,
  RESULTADO_ZARIT: 69, CONSULTA_PSICOLOGA: 70, CONSULTA_TRABAJADORA_SOCIAL: 71,
  SIGGES: 72, IVADEC: 73, FICHA_FAMILIAR: 74, ZONA_EVACUACION: 75,
  AYUDAS_TECNICAS: 76,
  PAÑALES: 77, TALLA_PAÑALES: 78, F_ENTREGA_PAÑALES: 79,
  F_INFLUENZA_U: 80, F_NEUMO23_U: 81, F_INFLUENZA_C: 82, F_NEUMO23_C: 83,
  CAP_INI: 84, CAP_FIN: 91,
  SONDA_FOLEY: 92, TIPO_SONDA: 93, CALIBRE_SONDA: 94,
  F_CAMBIO_SONDA: 95,
  LPP: 96, UBICACION_LPP: 97, ESTADIO_LPP: 98, TTO_INVASIVOS: 99,
  TIPO_CURACION: 100, F_ULTIMA_CURACION: 101, F_PROXIMA_CURACION: 102,
  F_ALTA_LPP: 103,
  P_A: 104, HEMOGLOBINA_GLICOCILADA: 105, LDL_70: 106, RAC: 107, VFG: 108,
  PRIORIDAD: 109, OBSERVACIONES: 110, EDITOR: 111,
}

var _FECHA_BY_COL = {}
for (var _fi = 0; _fi < _CONTROL_FECHAS.length; _fi++) {
  _FECHA_BY_COL[_CONTROL_FECHAS[_fi][1]] = _CONTROL_FECHAS[_fi]
}

var TEXT_UPPER = [2, 3, 4, 5, 11, 16]

var _RELEVANT_COLS = {}
for (var _ri = 0; _ri < TEXT_UPPER.length; _ri++) _RELEVANT_COLS[TEXT_UPPER[_ri]] = true
_RELEVANT_COLS[COL.RUN] = true
_RELEVANT_COLS[COL.RUN_CUIDADOR] = true
_RELEVANT_COLS[COL.TELEFONO] = true
_RELEVANT_COLS[COL.OBSERVACIONES] = true
_RELEVANT_COLS[COL.EDITOR] = true
_RELEVANT_COLS[COL.EDAD_USUARIO] = true
_RELEVANT_COLS[COL.EDAD_CUIDADOR] = true
_RELEVANT_COLS[COL.SEXO] = true
_RELEVANT_COLS[COL.SEXO_CUIDADOR] = true
_RELEVANT_COLS[9] = true
_RELEVANT_COLS[COL.VITAL] = true
for (var _ri = 0; _ri < _CONTROL_FECHAS.length; _ri++) {
  _RELEVANT_COLS[_CONTROL_FECHAS[_ri][1]] = true
}
for (var _ri = COL.CAP_INI; _ri <= COL.CAP_FIN; _ri++) _RELEVANT_COLS[_ri] = true
_RELEVANT_COLS[COL.EMPA_CUIDADOR] = true
_RELEVANT_COLS[COL.EMPA_USUARIO] = true

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
  'EGRESO POR ALTA': ['#eceff1', '#6b7280'],
  ALTA:      ['#bbdefb', '#1565c0'],
  PENDIENTE: ['#fff9c4', '#f9a825'],
  TRASLADO:  ['#b2dfdb', '#00695c'],
}

// en _actualizarEstadosFila y en las reglas condicionales de 06_Formato.

var _VITAL_ROW_COLORS = {
  FALLECIDO:  { fg: '#aaaaaa', strike: true,  italic: false },
  SUSPENDIDO: { fg: '#888888', strike: false, italic: true },
  EGRESO:     { fg: '#888888', strike: false, italic: false },
  'EGRESO POR ALTA': { fg: '#9e9e9e', strike: true, italic: false },
  ALTA:       { fg: '#1565c0', strike: false, italic: false },
  TRASLADO:   { fg: '#00695c', strike: false, italic: false },
}

var _ESTADO_FECHA_COLORS = {
  'AL DIA':     ['#c8e6c9', '#2e7d32'],
  'POR VENCER': ['#ffe0b2', '#e65100'],
  'VENCIDO':    ['#ffcdd2', '#c62828'],
  'PENDIENTE':  ['#fff9c4', '#f9a825'],
  'N/A':        ['#f5f5f5', '#999999'],
}

// ─── GUÍA DE COLUMNAS (111) ────────────────────────────────────────────────

var _COLUMNAS = (function() {
  var C = new Array(111), _cnt = 0
  function s(n, name, desc, vals, auto) { C[n] = { n: n, name: name, desc: desc, vals: vals || '', auto: auto || '' }; _cnt = Math.max(_cnt, n) }

  s(1,  'N°',                  'Número correlativo de fila', '', 'Automático')
  s(2,  'SECTOR',              'Sector geográfico asignado al paciente', 'VERDE · AMARILLO · NARANJO · Pendiente', 'Dropdown')
  s(3,  'NOMBRE',              'Nombre del paciente (completo)', '', 'Mayúscula automática')
  s(4,  'AP. PATERNO',         'Primer apellido del paciente', '', 'Mayúscula automática')
  s(5,  'AP. MATERNO',         'Segundo apellido del paciente', '', 'Mayúscula automática')
  s(6,  'ESTADO',              'Estado de vigencia del registro', 'VIGENTE · FALLECIDO · EGRESO · EGRESO POR ALTA · SUSPENDIDO · ALTA · TRASLADO · PENDIENTE', 'Dropdown')
  s(7,  'SEXO',                'Sexo del paciente', 'F · M · Pend.', 'Dropdown')
  s(8,  'RUN',                 'RUT del paciente (con guión, sin puntos)', '', 'Formato automático')
  s(9,  'F. NACIMIENTO',       'Fecha de nacimiento del paciente', '', 'dd/mm/aaaa')
  s(10, 'EDAD',                'Edad en años del paciente', '', 'Automático desde F. NACIMIENTO')
  s(11, 'DIRECCION',           'Dirección del domicilio del paciente', '', 'Mayúscula automática')
  s(12, 'TELEFONO',            'Teléfono de contacto del paciente (se formatea solo)', '', 'Formato automático')
  s(13, 'F. INGRESO PADI',     'Fecha de ingreso al programa PADDS', '', 'dd/mm/aaaa')
  s(14, 'F. EGRESO',           'Fecha de egreso del programa (si aplica)', '', 'dd/mm/aaaa')
  s(15, 'F. FALLECIMIENTO',    'Fecha de fallecimiento del paciente (si aplica)', '', 'dd/mm/aaaa')

  s(16, 'NOMBRE CUIDADOR',     'Nombre completo del cuidador principal', '', 'Mayúscula automática')
  s(17, 'SEXO C.',             'Sexo del cuidador', 'F · M · Pend.', 'Dropdown')
  s(18, 'RUN C.',              'RUT del cuidador (con guión, sin puntos)', '', 'Formato automático')
  s(19, 'EDAD C.',             'Edad del cuidador en años', '', 'Automático o manual')
  s(20, 'PATOLOGIAS CUIDADOR', 'Patologías crónicas del cuidador (anotación)', '', 'Texto libre')
  s(21, 'PSCV CUIDADOR',       'PSCV del cuidador', 'SI · NO · N/A · Prob. Domicilio · Rechaza', 'Dropdown')
  s(22, 'EMPA/EMPAM C.',       'Tipo EMPA/EMPAM del cuidador según edad', 'EMPA · EMPAM · N/A · PENDIENTE', 'Automático según edad')
  s(23, 'F. EMPA/EMPAM C.',    'Fecha del último EMPA/EMPAM del cuidador', 'Fecha · N/A · (vacío = pendiente)', 'Color por vigencia')
  s(24, 'F. EXAMENES C.',      'Fecha de exámenes del cuidador', 'Fecha · N/A · (vacío = pendiente)', 'Color por vigencia')
  s(25, 'CCV VIGENTE C.',      'Fecha de la última CCV del cuidador', 'Fecha · N/A', 'Color por vigencia')

  s(26, 'BARTHEL',             'Índice Barthel — dependencia física', 'SEVERO · MODERADO · LEVE · INDEPENDIENTE · N/A', 'Dropdown')
  s(27, 'ONCOLOGICO',          'Paciente oncológico', 'SI · NO · N/A', 'Dropdown')
  s(28, 'CARDEX',              'Cardiopatía descompensada / riesgo cardiovascular', 'SI · NO · N/A', 'Dropdown')
  s(29, 'F. PIC-1C',           'Fecha del Plan de Intervención Individual 1 componente', 'Fecha · N/A', 'Color por vigencia')
  s(30, 'F. PIC-2C',           'Fecha del Plan de Intervención Individual 2 componentes', 'Fecha · N/A', 'Color por vigencia')
  s(31, 'ELECTRODEPENDENCIA',  'Paciente electrodependiente', 'Casilla: marcada = sí · vacía = no', 'Checkbox')

  var PATS = ['HTA','DLP','DM','ACV','IAM','GLAUCOMA','ASMA','EPOC','EPILEPSIA','PARKINSON','ARTROSIS','HPT','DEPRESION','ALZHEIMER','DEMENCIA','ESQUIZOFRENIA']
  for (var i = 0; i < PATS.length; i++) s(32 + i, PATS[i], 'Diagnóstico de ' + PATS[i], 'Casilla: marcada = padece · vacía = no', 'Checkbox')
  s(48, 'OTRAS PATOLOGIAS',    'Otras patologías no listadas (especificar)', '', 'Texto libre')

  s(49, 'PSCV',                'PSCV del usuario', 'SI · NO · N/A · Prob. Domicilio · Rechaza', 'Dropdown')
  s(50, 'MORBILIDAD',          'Comorbilidades y diagnósticos relevantes (texto libre)', '', 'Texto libre')
  s(51, 'F. EXAMENES U.',      'Fecha de exámenes del usuario', 'Fecha · N/A', 'Color por vigencia')
  s(52, 'F. CONTROL MEDICO',   'Fecha del último control médico', 'Fecha · N/A', 'Color por vigencia')
  s(53, 'F. CCV MEDICO',       'Fecha control cardiovascular (médico)', 'Fecha · N/A', 'Color por vigencia')
  s(54, 'F. CSCV ENFERMERIA',  'Fecha control cardiovascular (enfermería)', 'Fecha · N/A', 'Color por vigencia')
  s(55, 'RIESGO PIE DM',       'Evaluación del riesgo de pie diabético (paciente con DM)', 'SI · NO · N/A', 'Dropdown')
  s(56, 'F. PODOLOGO',         'Fecha de atención de podología', 'Fecha · N/A', 'Color por vigencia')
  s(57, 'F. NUTRICIONISTA',    'Fecha del control nutricional', 'Fecha · N/A', 'Color por vigencia')
  s(58, 'ESTADO NUTRICIONAL',  'Clasificación del estado nutricional', 'NORMAL · SOBREPESO · OBESIDAD · BAJO PESO · N/A', 'Dropdown')
  s(59, 'F. FONOAUDIOLOGA',    'Fecha de atención de fonoaudiología', 'Fecha · N/A', 'Color por vigencia')
  s(60, 'F. CONTROL KINESICO', 'Fecha del control kinésico', 'Fecha · N/A', 'Color por vigencia')
  s(61, 'F. ODONTOLOGIA',      'Fecha de atención odontológica', 'Fecha · N/A', 'Color por vigencia')
  s(62, 'CONTROLES MISCELANEOS', 'Anotación de controles no listados: fecha + prestación (ej: 15/06/2026 - Control glucosa)', '', 'Texto libre acumulativo')
  s(63, 'RECETAS CONTROLADAS', 'NOMBRE del medicamento controlado (ej: CLONAZEPAM)', 'Lista editable (se puede escribir otro)', 'Dropdown')
  s(64, 'F. RECETA',           'Fecha de la última receta controlada', 'Fecha · N/A', 'Color por vigencia')

  s(65, 'EMPA/EMPAM U.',       'Tipo EMPA/EMPAM del usuario según edad', 'EMPA · EMPAM · N/A · PENDIENTE', 'Automático según edad')
  s(66, 'F. EMPA/EMPAM U.',    'Fecha del último EMPA/EMPAM del usuario', 'Fecha · N/A', 'Color por vigencia')
  s(67, 'ESTIPENDIO',          'Tipo de estipendio o beneficio social', 'BENEFICIARIO · INGRESADO · NO INGRESA · NO APLICA · PENDIENTE · EN ESPERA · RECHAZA · N/A', 'Dropdown')
  s(68, 'F. ZARIT',            'Fecha de aplicación escala Zarit', 'Fecha · N/A', 'Color por vigencia')
  s(69, 'RESULTADO ZARIT',     'Resultado escala Zarit (sobrecarga cuidador)', 'N/A · PENDIENTE · RECHAZA · SIN SOBRECARGA · SOBRECARGA LEVE · SOBRECARGA INTENSA · CUID. REMUNERADA · AUSENCIA', 'Dropdown')
  s(70, 'CONSULTA PSICOLOGA',  'Consulta psicológica realizada', 'SI · NO · N/A', 'Dropdown')
  s(71, 'CONSULTA TRABAJADORA SOCIAL', 'Consulta con trabajadora social', 'SI · NO · N/A', 'Dropdown')
  s(72, 'SIGGES',              'Registro en SIGGES', 'SI · NO · N/A', 'Dropdown')
  s(73, 'IVADEC',              'Índice de Valoración de Dependencia (IVADEC)', 'SI · NO · N/A', 'Dropdown')
  s(74, 'FICHA FAMILIAR',      'Ficha familiar realizada', 'SI · NO · N/A', 'Dropdown')
  s(75, 'ZONA EVACUACION',     'Zona de evacuación asignada', '', 'Texto libre')
  s(76, 'AYUDAS TECNICAS',     'Ayudas técnicas requeridas/entregadas', 'SI · NO · N/A · PENDIENTE · BASTÓN · ANDADOR · SILLA RUEDAS · MULETAS', 'Dropdown')

  s(77, 'PAÑALES',             'Paciente beneficiario de pañales', 'Casilla: marcada = sí', 'Checkbox')
  s(78, 'TALLA PAÑALES',       'Talla de pañales entregados', 'G · M · XG · N/A', 'Dropdown')
  s(79, 'F. ENTREGA PAÑALES',  'Fecha de la última entrega de pañales', 'Fecha · N/A', 'Color por vigencia')

  s(80, 'F. INFLUENZA U.',     'Estado de vacunación influenza del usuario', 'SI · NO · N/A · R (refuerzo) · P (programada)')
  s(81, 'F. NEUMO23 U.',       'Estado de vacunación antineumocócica (Neumo23) del usuario', 'SI · NO · N/A · R (refuerzo) · P (programada)')
  s(82, 'F. INFLUENZA C.',     'Estado de vacunación influenza del cuidador', 'SI · NO · N/A · R (refuerzo) · P (programada)')
  s(83, 'F. NEUMO23 C.',       'Estado de vacunación antineumocócica (Neumo23) del cuidador', 'SI · NO · N/A · R (refuerzo) · P (programada)')

  var CAPS = ['CAP. KINESIOLOGIA','CAP. ENFERMERIA','CAP. TENS','CAP. TENS ESTIVAL','CAP. TRABAJO SOCIAL','CAP. NUTRICION','CAP. PSICOLOGIA','CAP. FONOAUDIOLOGIA']
  for (var i = 0; i < CAPS.length; i++) s(84 + i, CAPS[i], 'Capacitación: ' + CAPS[i] + ' — fecha de realización', 'Fecha · N/A', 'Color por vigencia')

  s(92, 'SONDA FOLEY',         'El usuario porta sonda vesical', 'Casilla: marcada = sí', 'Checkbox')
  s(93, 'TIPO SONDA',          'Tipo de sonda', 'VESICAL · SUPRAPUBICA · N/A', 'Dropdown')
  s(94, 'CALIBRE SONDA',       'Calibre de la sonda (medida FR)', 'FR12 · FR14 · FR16 · FR18 · FR20 · N/A', 'Dropdown')
  s(95, 'F. ULTIMO CAMBIO',    'Fecha del último cambio de sonda', 'Fecha · N/A', 'Color por vigencia')

  s(96, 'LPP',                 'El usuario presenta Lesión por Presión', 'Casilla: marcada = sí', 'Checkbox')
  s(97, 'UBICACION LPP',       'Ubicación(es) de la(s) LPP (ej: sacro, talón izq.)', '', 'Texto libre')
  s(98, 'ESTADIO LPP',         'Estadio de la LPP (clasificación)', 'I · II · III · IV · N/A', 'Dropdown')
  s(99, 'TTO INVASIVOS',       'Tratamientos invasivos: CISTOSTOMÍA, SNAP, OSTOMÍA, sonda, etc.', '', 'Texto libre')
  s(100, 'TIPO CURACION',      'Tipo de curación aplicada', 'AVANZADA · SIMPLE · N/A', 'Dropdown')
  s(101, 'F. ULTIMA CURACION', 'Fecha de la última curación', 'Fecha · N/A', 'Color por vigencia')
  s(102, 'F. PROXIMA CURACION','Fecha de la próxima curación (vence al pasar la fecha)', 'Fecha · N/A', 'Color por vigencia')
  s(103, 'F. ALTA LPP',        'Fecha de alta de la LPP', 'Fecha · N/A', 'dd/mm/aaaa')

  s(104, 'P/A',                    'Presión arterial (ej: 120/80)', '', 'Texto libre')
  s(105, 'HEMOGLOBINA GLICOCILADA', 'Hemoglobina glicosilada (HbA1c) — último resultado', '', 'Texto libre')
  s(106, 'LDL - 70',               'Colesterol LDL (objetivo menor a 70)', '', 'Texto libre')
  s(107, 'RAC',                    'Relación albúmina/creatinina', '', 'Texto libre')
  s(108, 'VFG',                    'Velocidad de filtración glomerular', '', 'Texto libre')

  s(109, 'PRIORIDAD GENERAL',  'Prioridad calculada según el peor control (fechas vencidas)', 'URGENTE · POR REVISAR · AL DIA · N/A', 'Automático')
  s(110, 'OBSERVACIONES',      'Próximo control o visita + observaciones de la última atención', '', 'Texto libre')
  s(111, 'EDITOR',             'Correo del último usuario que editó la fila', '', 'Automático')

  C._count = _cnt
  return C
})()

for (var _ci = 0; _ci < _COLUMNAS.length; _ci++) {
  var _c = _COLUMNAS[_ci]
  if (!_c) continue
  var _key = _c.name.replace(/[^A-Z0-9_]/gi, '_').replace(/_+/g, '_').replace(/^_|_$/g, '').toUpperCase()
  if (COL[_key] === undefined) COL[_key] = _ci
}

var FORM_SECCIONES = [
  { nombre: 'GESTIÓN',     ini: 1,  fin: 5,  bg: '#2c3e50', fg: '#ffffff' },
  { nombre: 'REGISTRO',    ini: 6,  fin: 9,  bg: '#1565C0', fg: '#ffffff' },
  { nombre: 'ATENCIÓN',    ini: 10, fin: 16, bg: '#2E7D5B', fg: '#ffffff' },
  { nombre: 'CIERRE',      ini: 17, fin: 17, bg: '#6A1B9A', fg: '#ffffff' },
]

var FORM_ESTADOS = ['Pendiente', 'Aprobado', 'Rechazado']

var FORM_A_PAC = [
  [6, null],
  [7, 8],
  [8, 3],
  [9, 4],
  [10, COL.EDITOR],
  [12, 58],
  [14, 26],
  [15, 69],
  [17, COL.OBSERVACIONES],
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
  'EXAMEN DE SANGRE USUARIO': 51,
  'TOMA DE EXAMENES': 51,
  'EXAMEN DE SANGRE CUIDADOR': 24,
  'MORBILIDAD': 50,
  'CONTROL MEDICO': 52,
  'CONTROL DE SALUD (INCLUYE CSV)': 52,
  'PRESION ARTERIAL': 54,
  'CCV MEDICO': 53,
  'CCV ENFERMERIA': 54,
  'CCV NUTRICIONAL': 57,
  'CSCV ENFERMERIA': 54,
  'PODOLOGO': 56,
  'ATENCION PODOLOGICA': 56,
  'NUTRICIONISTA': 57,
  'CONSULTA NUTRICIONAL': 57,
  'FONOAUDIOLOGA': 59,
  'ATENCION FONOAUDIOLOGA': 59,
  'CONTROL KINESICO': 60,
  'CONSULTA KINE': 60,
  'CONSULTA KINESIOLOGO': 60,
  'ODONTOLOGIA': 61,
  'EMPA': 66,
  'EMPAM': 66,
  'PIC-1C': 29,
  'VISITA PIC 1': 29,
  'PIC-2C': 30,
  'VISITA PIC 2': 30,
  'EMPA/EMPAM CUIDADOR': 23,
  'CCV CUIDADOR': 25,
  'ZARIT': 68,
  'APLICACION ZARIT': 68,
  'CAMBIO O INSTALACION DE SONDA': 95,
  'CURACION': 101,
  'CURACION AVANZADA': 101,
  'CONSULTA TRABAJADORA SOCIAL': 71,
  'CONSULTA PSICOLOGICA': 70,
  'CAPACITACION KINESIOLOGA': 84,
  'CAPACITACION ENFERMERIA': 85,
  'CAPACITACION TENS': 86,
  'CAPACITACION TRABAJADORA SOCIAL': 88,
  'CAPACITACION PSICOLOGA': 90,
  'CAPACITACION FONOAUDIOLOGA': 91,
}

var _SERVICIO_SI_MAP = {
  'INMUNIZACION INFLUENZA USUARIO': 80,
  'INMUNIZACION NEUMO 23 USUARIO': 81,
  'INMUNIZACION INFLUENZA CUIDADOR': 82,
  'INMUNIZACION NEUMO 23 CUIDADOR': 83,
}

var _PRESTACIONES_LIST = []
var _plSet = {}
function _regPrestacion(n) { if (!_plSet[n]) { _plSet[n] = 1; _PRESTACIONES_LIST.push(n) } }
for (var _pk1 in _SERVICIO_COL_MAP) _regPrestacion(_pk1)
for (var _pk2 in _SERVICIO_SI_MAP) _regPrestacion(_pk2)
_regPrestacion('CONTROL CUIDADOR')
_regPrestacion('CONTROL DE SALUD (INCLUYE CSV)')

var _FORM_VAL_MAP = {
  58: {
    'SOBRE PESO': 'SOBREPESO',
  },
}

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
  SLOTS_DEFAULT, SLOTS_DEFAULT, SLOTS_DEFAULT, SLOTS_DEFAULT,
  [
    ['8:00','abreviadas'], ['8:15','abreviadas'], ['8:30','abreviadas'], ['8:45','abreviadas'],
    ['9:00','VDI'], ['10:00','VDI'], ['11:00','VDI'], ['12:00','VDI'],
  ],
]

function _calcularRS() {
  var total = 1
  for (var d = 0; d < 5; d++) {
    total += 3 + (SLOTS[d] || SLOTS_DEFAULT).length
  }
  return total
}

