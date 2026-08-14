// ════════════════════════════════════════════════════════
// ÍNDICE ▏00_Constantes.gs │ constantes globales PADDS

// ─────────────────────────────────────────────────────────
// ─── CONSTANTES GLOBALES ─────────────────────────────────────────────────────

// color indica la vigencia (verde al día / ámbar por vencer / rojo vencido /

var HOJA_PAC = 'Pacientes'
var HOJA_FORM = 'Recepción Formulario Profesional'
var PAC_SECCIONES = [
  { nombre: 'IDENTIFICACIÓN',             ini: 1,  fin: 15, bg: '#3B5B7C', bg2: '#4D6D8E', fg: '#ffffff' },
  { nombre: 'CUIDADOR PRINCIPAL',         ini: 16, fin: 25, bg: '#94603C', bg2: '#A5704C', fg: '#ffffff' },
  { nombre: 'CLASIFICACIÓN / DEPENDENCIA',ini: 26, fin: 31, bg: '#62508A', bg2: '#74629C', fg: '#ffffff' },
  { nombre: 'PATOLOGÍAS CRÓNICAS',        ini: 32, fin: 48, bg: '#933F3F', bg2: '#A55151', fg: '#ffffff' },
  { nombre: 'CONTROLES Y SEGUIMIENTO',    ini: 49, fin: 65, bg: '#2E6E61', bg2: '#408476', fg: '#ffffff' },
  { nombre: 'SOCIAL / PSICOLÓGICO',       ini: 66, fin: 77, bg: '#8F4662', bg2: '#A15872', fg: '#ffffff' },
  { nombre: 'PAÑALES',                    ini: 78, fin: 80, bg: '#37798A', bg2: '#478A9B', fg: '#ffffff' },
  { nombre: 'INMUNIZACIÓN',               ini: 81, fin: 84, bg: '#465B85', bg2: '#586D97', fg: '#ffffff' },
  { nombre: 'CAPACITACIONES CUIDADOR',    ini: 85, fin: 92, bg: '#5F5144', bg2: '#6F6154', fg: '#ffffff' },
  { nombre: 'SONDA FOLEY',                ini: 93, fin: 96, bg: '#48704D', bg2: '#58805D', fg: '#ffffff' },
  { nombre: 'CURACIONES AVANZADAS',       ini: 97, fin: 104,bg: '#A25A36', bg2: '#B26B47', fg: '#ffffff' },
  { nombre: 'CONTROL DE SIGNOS VITALES (CSV)', ini: 105,fin: 109,bg: '#923939', bg2: '#A44B4B', fg: '#ffffff' },
  { nombre: 'GESTIÓN',                    ini: 110,fin: 112,bg: '#43596B', bg2: '#546A7C', fg: '#ffffff' },
]

var PAC_ANCHOS = [
  40, 80, 170, 130, 130, 80, 45, 110, 100, 40,
  240, 170, 110, 100, 110, 190, 45, 110, 40,
  270, 80, 110, 110, 110, 110,
  100, 70, 70, 100, 100, 105,
  45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45,
  300, 70, 260, 110, 110, 110, 110, 100, 100, 110, 110,
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
var _VACUNA_VALS = ['SI', 'NO', 'N/A', 'R', 'P']

var _MEDICAMENTOS_LISTA = [
  'CLONAZEPAM', 'CITALOPRAM', 'ALPRAZOLAM', 'DIAZEPAM', 'LORAZEPAM',
  'ESCITALOPRAM', 'SERTRALINA', 'FLUOXETINA', 'VENLAFAXINA', 'QUETIAPINA',
  'RISPERIDONA', 'OLANZAPINA', 'ZOLPIDEM', 'GABAPENTINA', 'PREGABALINA',
  'TRAMADOL', 'MORFINA', 'METILFENIDATO', 'CARBAMAZEPINA', 'VALPROICO',
  'N/A',
]

var _CHECKBOX_COLS = [31]
for (var _cb = 32; _cb <= 47; _cb++) _CHECKBOX_COLS.push(_cb)
_CHECKBOX_COLS.push(56, 78, 93, 97)

var PAC_VALIDACIONES = {
  2: ['VERDE', 'AMARILLO', 'NARANJO', 'PENDIENTE'],
  6: ['VIGENTE', 'FALLECIDO', 'EGRESO', 'EGRESO POR ALTA', 'SUSPENDIDO', 'ALTA', 'TRASLADO', 'PENDIENTE'],
  7: ['F', 'M', 'PENDIENTE'],
  17: ['F', 'M', 'PENDIENTE'],
  21: PAC_SI_NO_PROB,
  22: ['EMPA', 'EMPAM', 'N/A', 'PENDIENTE'],
  26: ['SEVERO', 'SEVERO - CPU', 'SEVERO + CA', 'MODERADO', 'MOD + CERTIF',
    'LEVE', 'LEVE + CA', 'INDEPENDIENTE', 'N/A'],
  27: PAC_SI_NO, 28: PAC_SI_NO,
  49: PAC_SI_NO_PROB,
  55: ['NEUROPATICO', 'ISQUEMICO', 'NEUROISQUEMICO', 'DM 1', 'DM 2', 'N/A'],
  59: ['NORMAL', 'SOBREPESO', 'OBESIDAD', 'OBESIDAD MORBIDA', 'BAJO PESO', 'N/A'],
  66: ['EMPA', 'EMPAM', 'N/A', 'PENDIENTE'],
  68: ['BENEFICIARIO', 'INGRESADO', 'NO INGRESA', 'NO APLICA', 'PENDIENTE', 'EN ESPERA', 'RECHAZA', 'N/A'],
  70: ['N/A', 'PENDIENTE', 'RECHAZA', 'SIN SOBRECARGA', 'SOBRECARGA LEVE',
    'SOBRECARGA INTENSA', 'CUIDADORA REMUNERADA', 'AUSENCIA'],
  73: PAC_SI_NO, 74: ['P', 'N/A'], 75: PAC_SI_NO,
  77: ['SILLA RUEDAS', 'ANDADOR', 'BASTON', 'MULETAS',
    'CAMA', 'COLCHON ANTIESCARAS', 'COJIN + COLCHON', 'N/A'],
  79: ['S', 'G', 'M', 'XG', 'XXG', 'N/A'],
  81: _VACUNA_VALS, 82: _VACUNA_VALS, 83: _VACUNA_VALS, 84: _VACUNA_VALS,
  94: ['VESICAL', 'SUPRAPUBICA', 'TEMPORAL', 'TRANSITORIA', 'N/A'],
  95: ['FR12', 'FR14', 'FR16', 'FR18', 'FR20', 'FR22', 'FR24', 'N/A'],
  99: ['I', 'II', 'III', 'IV', 'N/A'],
  101: ['AVANZADA', 'SIMPLE', 'N/A'],
  64: _MEDICAMENTOS_LISTA,
  110: ['URGENTE', 'POR REVISAR', 'AL DIA', 'N/A'],
}

// Columnas de texto libre: el formateo SIEMPRE les quita cualquier validación
// (no deben tener dropdown, fecha ni casilla). NO incluir columnas con
// PAC_VALIDACIONES, _FECHAS_VA ni _CHECKBOX_COLS.
var PAC_LIBRES = [20, 48, 50, 63, 76, 98, 100, 105, 106, 107, 108, 109, 111]

var _FECHAS_VA = [9, 13, 14, 15, 23, 24, 25, 29, 30, 51, 52, 53, 54, 57, 58, 60, 61, 62,
  65, 67, 69, 71, 72, 80, 85, 86, 87, 88, 89, 90, 91, 92,
  96, 102, 103, 104]

var _VACUNA_COLS = [81, 82, 83, 84]

var _CONTROL_FECHAS = [
  ['EXAMENES CUIDADOR',       24, 'EXAMENES CUIDADOR'],
  ['EMPA/EMPAM CUIDADOR',     23, 'EMPA/EMPAM CUIDADOR'],
  ['PIC-1C',                  29, 'PIC-1C'],
  ['PIC-2C',                  30, 'PIC-2C'],
  ['EXAMENES USUARIO',        51, 'EXAMENES USUARIO'],
  ['CONTROL MEDICO',          52, 'CONTROL MEDICO'],
  ['CCV MEDICO',              53, 'CCV MEDICO'],
  ['CSCV ENFERMERIA',         54, 'CSCV ENFERMERIA'],
  ['PODOLOGO',                57, 'PODOLOGO'],
  ['NUTRICIONISTA',           58, 'NUTRICIONISTA'],
  ['FONOAUDIOLOGA',           60, 'FONOAUDIOLOGA'],
  ['CONTROL KINESICO',        61, 'CONTROL KINESICO'],
  ['ODONTOLOGIA',             62, 'ODONTOLOGIA'],
  ['EMPA/EMPAM USUARIO',      67, 'EMPA/EMPAM USUARIO'],
  ['RECETAS CONTROLADAS',     65, 'RECETAS CONTROLADAS'],
  ['CCV VIGENTE CUIDADOR',    25, 'CCV VIGENTE CUIDADOR'],
  ['ZARIT',                   69, 'ZARIT'],
  ['SONDA FOLEY',             96, 'SONDA FOLEY'],
  ['CURACIONES',             102, 'CURACIONES'],
  ['PROXIMA CURACION',       103, 'PROXIMA CURACION'],
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
  F_CCV_MEDICO: 53, F_CSCV_ENFERMERIA: 54, RIESGO_PIE_DM: 55,
  INSULINO_DEPENDIENTE: 56,
  F_PODOLOGO: 57,
  F_NUTRICIONISTA: 58, ESTADO_NUTRICIONAL: 59, F_FONOAUDIOLOGA: 60,
  F_CONTROL_KINESICO: 61, F_ODONTOLOGIA: 62, CONTROLES_MISCELANEOS: 63,
  RECETAS_CONTROLADAS: 64, F_RECETA: 65,
  EMPA_USUARIO: 66, F_EMPA_USUARIO: 67, ESTIPENDIO: 68, ZARIT: 69,
  RESULTADO_ZARIT: 70, CONSULTA_PSICOLOGA: 71, CONSULTA_TRABAJADORA_SOCIAL: 72,
  SIGGES: 73, IVADEC: 74, FICHA_FAMILIAR: 75, ZONA_EVACUACION: 76,
  AYUDAS_TECNICAS: 77,
  PAÑALES: 78, TALLA_PAÑALES: 79, F_ENTREGA_PAÑALES: 80,
  F_INFLUENZA_U: 81, F_NEUMO23_U: 82, F_INFLUENZA_C: 83, F_NEUMO23_C: 84,
  CAP_INI: 85, CAP_FIN: 92,
  SONDA_FOLEY: 93, TIPO_SONDA: 94, CALIBRE_SONDA: 95,
  F_CAMBIO_SONDA: 96,
  LPP: 97, UBICACION_LPP: 98, ESTADIO_LPP: 99, TTO_INVASIVOS: 100,
  TIPO_CURACION: 101, F_ULTIMA_CURACION: 102, F_PROXIMA_CURACION: 103,
  F_ALTA_LPP: 104,
  P_A: 105, HEMOGLOBINA_GLICOCILADA: 106, LDL_70: 107, RAC: 108, VFG: 109,
  PRIORIDAD: 110, OBSERVACIONES: 111, EDITOR: 112,
}

// Columnas con dropdown de MULTISELECCIÓN: cada opción elegida se agrega o
// quita de la lista separada por comas (ver _toggleMultiselect en 04_Eventos).
var PAC_MULTISELECT = [COL.AYUDAS_TECNICAS, COL.RIESGO_PIE_DM, COL.RECETAS_CONTROLADAS]

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
  // Tintas claras + texto profundo: mejor contraste que el par ámbar/amarillo anterior.
  VERDE:    ['#DCFCE7', '#15803D'],
  AMARILLO: ['#FEF3C7', '#B45309'],
  NARANJO:  ['#FFEDD5', '#C2410C'],
  // PENDIENTE = sector aún NO asignado: gris pizarra (distinto del AMARILLO).
  PENDIENTE:['#F1F5F9', '#64748B'],
}

var _ESTADO_COLORS = {
  VIGENTE:   ['#DCFCE7', '#15803D'],
  FALLECIDO: ['#F1F5F9', '#64748B'],
  SUSPENDIDO:['#EEF0F4', '#475569'],
  EGRESO:    ['#E0F2FE', '#0369A1'],
  'EGRESO POR ALTA': ['#EEF1F5', '#64748B'],
  ALTA:      ['#F0FDFA', '#0F766E'],
  PENDIENTE: ['#FEF3C7', '#B45309'],
  TRASLADO:  ['#CCFBF1', '#115E59'],
}

// en _actualizarEstadosFila y en las reglas condicionales de 06_Formato.

var _VITAL_ROW_COLORS = {
  FALLECIDO:  { fg: '#64748B', strike: true,  italic: false },
  SUSPENDIDO: { fg: '#475569', strike: false, italic: true },
  EGRESO:     { fg: '#0369A1', strike: false, italic: false },
  'EGRESO POR ALTA': { fg: '#64748B', strike: true, italic: false },
  ALTA:       { fg: '#0F766E', strike: false, italic: false },
  TRASLADO:   { fg: '#115E59', strike: false, italic: false },
}

var _ESTADO_FECHA_COLORS = {
  'AL DIA':     ['#DCFCE7', '#15803D'],
  'POR VENCER': ['#FFEDD5', '#C2410C'],
  'VENCIDO':    ['#FEE2E2', '#B91C1C'],
  'PENDIENTE':  ['#FEF3C7', '#B45309'],
  'N/A':        ['#F1F5F9', '#64748B'],
}

// ─── COLORES DE OPCIONES DE DROPDOWN (una celda coloreada por opción en la
// hoja oculta _Opciones: el dropdown "desde un rango" muestra cada opción con
// el color de su celda fuente y el chip seleccionado toma ese color) ───────

var _OPC_BG = {
  'SI': '#15803D', 'NO': '#B91C1C', 'N/A': '#64748B',
  'VERDE': '#15803D', 'AMARILLO': '#B45309', 'NARANJO': '#C2410C', 'PENDIENTE': '#B45309',
  'VIGENTE': '#15803D', 'FALLECIDO': '#64748B', 'SUSPENDIDO': '#475569',
  'EGRESO': '#0369A1', 'EGRESO POR ALTA': '#64748B', 'ALTA': '#0F766E', 'TRASLADO': '#115E59',
  'F': '#BE185D', 'M': '#0369A1', 'PEND.': '#B45309',
  'SEVERO': '#B91C1C', 'MODERADO': '#C2410C', 'LEVE': '#B45309', 'INDEPENDIENTE': '#15803D',
  'SEVERO - CPU': '#991B1B', 'SEVERO + CA': '#B91C1C',
  'MOD + CERTIF': '#C2410C', 'LEVE + CA': '#B45309',
  'CAMA': '#0E7490', 'COLCHON ANTIESCARAS': '#0E7490', 'COJIN + COLCHON': '#0E7490',
  'NEUROPATICO': '#7E22CE', 'ISQUEMICO': '#B91C1C', 'NEUROISQUEMICO': '#BE185D',
  'DM 1': '#0369A1', 'DM 2': '#0E7490',
  'NORMAL': '#15803D', 'SOBREPESO': '#C2410C', 'OBESIDAD': '#B91C1C', 'OBESIDAD MORBIDA': '#881337', 'BAJO PESO': '#B45309',
  'SIN SOBRECARGA': '#15803D', 'SOBRECARGA LEVE': '#C2410C', 'SOBRECARGA INTENSA': '#B91C1C',
  'CUIDADORA REMUNERADA': '#57534E', 'AUSENCIA': '#64748B',
  'EMPA': '#0369A1', 'EMPAM': '#7E22CE',
  'BENEFICIARIO': '#15803D', 'INGRESADO': '#0F766E', 'NO INGRESA': '#B91C1C',
  'NO APLICA': '#64748B', 'EN ESPERA': '#B45309', 'RECHAZA': '#B91C1C',
  'BASTON': '#0E7490', 'ANDADOR': '#0E7490', 'SILLA RUEDAS': '#0E7490', 'MULETAS': '#0E7490',
  'G': '#0369A1', 'XG': '#7E22CE', 'XXG': '#881337', 'S': '#0E7490',
  'VESICAL': '#15803D', 'SUPRAPUBICA': '#7E22CE',
  'TEMPORAL': '#0E7490', 'TRANSITORIA': '#C2410C',
  'FR12': '#0E7490', 'FR14': '#0E7490', 'FR16': '#0E7490', 'FR18': '#0E7490', 'FR20': '#0E7490',
  'FR22': '#0E7490', 'FR24': '#0E7490',
  'I': '#15803D', 'II': '#C2410C', 'III': '#B91C1C', 'IV': '#881337',
  'AVANZADA': '#7E22CE', 'SIMPLE': '#0E7490',
  'R': '#15803D', 'P': '#B45309',
  'URGENTE': '#B91C1C', 'POR REVISAR': '#C2410C', 'AL DIA': '#15803D',
  'PROB. DOMICILIO': '#D97706',
}

// Opciones sin color semántico definido: se les asigna por turno estos colores.
var _PALETA_OPCIONES = ['#475569', '#0F766E', '#7E22CE', '#0E7490', '#BE185D',
  '#57534E', '#0369A1', '#15803D', '#C2410C', '#B91C1C']

// Sobrescritura del color de UNA opción dentro de UNA columna concreta.
// "PENDIENTE" en SECTOR (col 2) es "sector sin asignar" → gris azulado, distinto
// del valor "AMARILLO". En el resto de columnas PENDIENTE mantiene su ámbar.
var _OPC_BG_COL = {
  2: { 'PENDIENTE': '#64748B', 'PEND.': '#64748B' },
}

// ─── GUÍA DE COLUMNAS (112) ────────────────────────────────────────────────

var _COLUMNAS = (function() {
  var C = new Array(112), _cnt = 0
  function s(n, name, desc, vals, auto) { C[n] = { n: n, name: name, desc: desc, vals: vals || '', auto: auto || '' }; _cnt = Math.max(_cnt, n) }

  s(1,  'N°',                  'Número correlativo de fila', '', 'Automático')
  s(2,  'SECTOR',              'Sector geográfico asignado al paciente', 'VERDE · AMARILLO · NARANJO · Pendiente', 'Dropdown')
  s(3,  'NOMBRE',              'Nombre del paciente (completo)', '', 'Mayúscula automática')
  s(4,  'AP. PATERNO',         'Primer apellido del paciente', '', 'Mayúscula automática')
  s(5,  'AP. MATERNO',         'Segundo apellido del paciente', '', 'Mayúscula automática')
  s(6,  'ESTADO',              'Estado de vigencia del registro', 'VIGENTE · FALLECIDO · EGRESO · EGRESO POR ALTA · SUSPENDIDO · ALTA · TRASLADO · PENDIENTE', 'Dropdown')
  s(7,  'SEXO',                'Sexo del paciente', 'F · M · Pend.', 'Dropdown')
  s(8,  'RUN',                 'RUT del paciente (con guión, sin puntos) — el dígito verificador se valida solo (celda roja si es inválido)', '', 'Validación automática')
  s(9,  'F. NACIMIENTO',       'Fecha de nacimiento del paciente', '', 'dd/mm/aaaa')
  s(10, 'EDAD',                'Edad en años del paciente', '', 'Automático desde F. NACIMIENTO')
  s(11, 'DIRECCION',           'Dirección del domicilio del paciente', '', 'Mayúscula automática')
  s(12, 'TELEFONO',            'Teléfono de contacto del paciente (se formatea solo)', '', 'Formato automático')
  s(13, 'F. INGRESO PADI',     'Fecha de ingreso al programa PADDS', '', 'dd/mm/aaaa')
  s(14, 'F. EGRESO',           'Fecha de egreso del programa (si aplica)', '', 'dd/mm/aaaa')
  s(15, 'F. FALLECIMIENTO',    'Fecha de fallecimiento del paciente (si aplica)', '', 'dd/mm/aaaa')

  s(16, 'NOMBRE CUIDADOR',     'Nombre completo del cuidador principal', '', 'Mayúscula automática')
  s(17, 'SEXO C.',             'Sexo del cuidador', 'F · M · Pend.', 'Dropdown')
  s(18, 'RUN C.',              'RUT del cuidador (con guión, sin puntos) — el dígito verificador se valida solo (celda roja si es inválido)', '', 'Validación automática')
  s(19, 'EDAD C.',             'Edad del cuidador en años', '', 'Automático o manual')
  s(20, 'PATOLOGIAS CUIDADOR', 'Patologías crónicas del cuidador (anotación)', '', 'Texto libre')
  s(21, 'PSCV CUIDADOR',       'PSCV del cuidador', 'SI · NO · N/A · Prob. Domicilio · Rechaza', 'Dropdown')
  s(22, 'EMPA/EMPAM C.',       'Tipo EMPA/EMPAM del cuidador según edad', 'EMPA · EMPAM · N/A · PENDIENTE', 'Automático según edad')
  s(23, 'F. EMPA/EMPAM C.',    'Fecha del último EMPA/EMPAM del cuidador', 'Fecha · N/A · (vacío = pendiente)', 'Color por vigencia')
  s(24, 'F. EXAMENES C.',      'Fecha de exámenes del cuidador', 'Fecha · N/A · (vacío = pendiente)', 'Color por vigencia')
  s(25, 'CCV VIGENTE C.',      'Fecha de la última CCV del cuidador', 'Fecha · N/A', 'Color por vigencia')

  s(26, 'BARTHEL',             'Índice Barthel — dependencia física', 'SEVERO · SEVERO - CPU · SEVERO + CA · MODERADO · MOD + CERTIF · LEVE · LEVE + CA · INDEPENDIENTE · N/A', 'Dropdown')
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
  s(55, 'RIESGO PIE DM',       'Tipo de pie diabético y diabetes mellitus (se pueden elegir varias, separadas por coma)', 'NEUROPATICO · ISQUEMICO · NEUROISQUEMICO · DM 1 · DM 2 · N/A', 'Multiselección')
  s(56, 'INSULINO DEPENDIENTE', 'Paciente insulino dependiente (DM que requiere insulina)', 'Casilla: marcada = sí · vacía = no', 'Checkbox')
  s(57, 'F. PODOLOGO',         'Fecha de atención de podología', 'Fecha · N/A', 'Color por vigencia')
  s(58, 'F. NUTRICIONISTA',    'Fecha del control nutricional', 'Fecha · N/A', 'Color por vigencia')
  s(59, 'ESTADO NUTRICIONAL',  'Clasificación del estado nutricional', 'NORMAL · SOBREPESO · OBESIDAD · BAJO PESO · N/A', 'Dropdown')
  s(60, 'F. FONOAUDIOLOGA',    'Fecha de atención de fonoaudiología', 'Fecha · N/A', 'Color por vigencia')
  s(61, 'F. CONTROL KINESICO', 'Fecha del control kinésico', 'Fecha · N/A', 'Color por vigencia')
  s(62, 'F. ODONTOLOGIA',      'Fecha de atención odontológica', 'Fecha · N/A', 'Color por vigencia')
  s(63, 'CONTROLES MISCELANEOS', 'Anotación de controles no listados: fecha + prestación (ej: 15/06/2026 - Control glucosa)', '', 'Texto libre acumulativo')
  s(64, 'RECETAS CONTROLADAS', 'NOMBRE del medicamento controlado (ej: CLONAZEPAM)', 'Lista editable (se puede escribir otro)', 'Dropdown')
  s(65, 'F. RECETA',           'Fecha de la última receta controlada', 'Fecha · N/A', 'Color por vigencia')

  s(66, 'EMPA/EMPAM U.',       'Tipo EMPA/EMPAM del usuario según edad', 'EMPA · EMPAM · N/A · PENDIENTE', 'Automático según edad')
  s(67, 'F. EMPA/EMPAM U.',    'Fecha del último EMPA/EMPAM del usuario', 'Fecha · N/A', 'Color por vigencia')
  s(68, 'ESTIPENDIO',          'Tipo de estipendio o beneficio social', 'BENEFICIARIO · INGRESADO · NO INGRESA · NO APLICA · PENDIENTE · EN ESPERA · RECHAZA · N/A', 'Dropdown')
  s(69, 'F. ZARIT',            'Fecha de aplicación escala Zarit', 'Fecha · N/A', 'Color por vigencia')
  s(70, 'RESULTADO ZARIT',     'Resultado escala Zarit (sobrecarga cuidador)', 'N/A · PENDIENTE · RECHAZA · SIN SOBRECARGA · SOBRECARGA LEVE · SOBRECARGA INTENSA · CUID. REMUNERADA · AUSENCIA', 'Dropdown')
  s(71, 'CONSULTA PSICOLOGA',  'Fecha de la consulta psicológica', 'Fecha · N/A', 'dd/mm/aaaa')
  s(72, 'CONSULTA TRABAJADORA SOCIAL', 'Fecha de la consulta con trabajadora social', 'Fecha · N/A', 'dd/mm/aaaa')
  s(73, 'SIGGES',              'Registro en SIGGES', 'SI · NO · N/A', 'Dropdown')
  s(74, 'IVADEC',              'Índice de Valoración de Dependencia (IVADEC)', 'P (pendiente/programada) · N/A', 'Dropdown')
  s(75, 'FICHA FAMILIAR',      'Ficha familiar realizada', 'SI · NO · N/A', 'Dropdown')
  s(76, 'ZONA EVACUACION',     'Zona de evacuación asignada', '', 'Texto libre')
  s(77, 'AYUDAS TECNICAS',     'Ayudas técnicas requeridas/entregadas (se pueden elegir varias, separadas por coma)', 'SILLA RUEDAS · ANDADOR · BASTON · MULETAS · CAMA · COLCHON ANTIESCARAS · COJIN + COLCHON · N/A', 'Multiselección')

  s(78, 'PAÑALES',             'Paciente beneficiario de pañales', 'Casilla: marcada = sí', 'Checkbox')
  s(79, 'TALLA PAÑALES',       'Talla de pañales entregados', 'G · M · XG · N/A', 'Dropdown')
  s(80, 'F. ENTREGA PAÑALES',  'Fecha de la última entrega de pañales', 'Fecha · N/A', 'Color por vigencia')

  s(81, 'INMUNIZ. INFLUENZA U.', 'Vacuna influenza del usuario — estado', 'SI · NO · N/A · R (refuerzo) · P (programada)', 'Dropdown')
  s(82, 'INMUNIZ. NEUMO23 U.',   'Vacuna antineumocócica (Neumo23) del usuario — estado', 'SI · NO · N/A · R (refuerzo) · P (programada)', 'Dropdown')
  s(83, 'INMUNIZ. INFLUENZA C.', 'Vacuna influenza del cuidador — estado', 'SI · NO · N/A · R (refuerzo) · P (programada)', 'Dropdown')
  s(84, 'INMUNIZ. NEUMO23 C.',   'Vacuna antineumocócica (Neumo23) del cuidador — estado', 'SI · NO · N/A · R (refuerzo) · P (programada)', 'Dropdown')

  var CAPS = ['CAP. KINESIOLOGIA','CAP. ENFERMERIA','CAP. TENS','CAP. TENS ESTIVAL','CAP. TRABAJO SOCIAL','CAP. NUTRICION','CAP. PSICOLOGIA','CAP. FONOAUDIOLOGIA']
  for (var i = 0; i < CAPS.length; i++) s(85 + i, CAPS[i], 'Capacitación: ' + CAPS[i] + ' — fecha de realización', 'Fecha · N/A', 'Color por vigencia')

  s(93, 'SONDA FOLEY',         'El usuario porta sonda vesical', 'Casilla: marcada = sí', 'Checkbox')
  s(94, 'TIPO SONDA',          'Tipo de sonda', 'VESICAL · SUPRAPUBICA · TEMPORAL · TRANSITORIA · N/A', 'Dropdown')
  s(95, 'CALIBRE SONDA',       'Calibre de la sonda (medida FR)', 'FR12 · FR14 · FR16 · FR18 · FR20 · FR22 · FR24 · N/A', 'Dropdown')
  s(96, 'F. ULTIMO CAMBIO',    'Fecha del último cambio de sonda', 'Fecha · N/A', 'Color por vigencia')

  s(97, 'LPP',                 'El usuario presenta Lesión por Presión', 'Casilla: marcada = sí', 'Checkbox')
  s(98, 'UBICACION LPP',       'Ubicación(es) de la(s) LPP (ej: sacro, talón izq.)', '', 'Texto libre')
  s(99, 'ESTADIO LPP',         'Estadio de la LPP (clasificación)', 'I · II · III · IV · N/A', 'Dropdown')
  s(100, 'TTO INVASIVOS',      'Tratamientos invasivos: CISTOSTOMÍA, SNAP, OSTOMÍA, sonda, etc.', '', 'Texto libre')
  s(101, 'TIPO CURACION',      'Tipo de curación aplicada', 'AVANZADA · SIMPLE · N/A', 'Dropdown')
  s(102, 'F. ULTIMA CURACION', 'Fecha de la última curación', 'Fecha · N/A', 'Color por vigencia')
  s(103, 'F. PROXIMA CURACION','Fecha de la próxima curación (vence al pasar la fecha)', 'Fecha · N/A', 'Color por vigencia')
  s(104, 'F. ALTA LPP',        'Fecha de alta de la LPP', 'Fecha · N/A', 'dd/mm/aaaa')

  s(105, 'P/A',                    'Presión arterial (ej: 120/80)', '', 'Texto libre')
  s(106, 'HEMOGLOBINA GLICOCILADA', 'Hemoglobina glicosilada (HbA1c) — último resultado', '', 'Texto libre')
  s(107, 'LDL - 70',               'Colesterol LDL (objetivo menor a 70)', '', 'Texto libre')
  s(108, 'RAC',                    'Relación albúmina/creatinina', '', 'Texto libre')
  s(109, 'VFG',                    'Velocidad de filtración glomerular', '', 'Texto libre')

  s(110, 'PRIORIDAD GENERAL',  'Prioridad calculada según el peor control (fechas vencidas)', 'URGENTE · POR REVISAR · AL DIA · N/A', 'Automático')
  s(111, 'OBSERVACIONES',      'Próximo control o visita + observaciones de la última atención', '', 'Texto libre')
  s(112, 'EDITOR',             'Correo del último usuario que editó la fila', '', 'Automático')

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
  { nombre: 'GESTIÓN',     ini: 1,  fin: 5,  bg: '#475569', fg: '#ffffff' },
  { nombre: 'REGISTRO',    ini: 6,  fin: 9,  bg: '#0F766E', fg: '#ffffff' },
  { nombre: 'ATENCIÓN',    ini: 10, fin: 16, bg: '#15803D', fg: '#ffffff' },
  { nombre: 'CIERRE',      ini: 17, fin: 19, bg: '#7E22CE', fg: '#ffffff' },
]

var FORM_ESTADOS = ['Pendiente', 'Gestionado', 'Rechazado']

var FORM_A_PAC = [
  [6, null],
  [7, 8],
  [8, 3],
  [9, 4],
  [10, COL.EDITOR],
  [12, 59],
  [14, 26],
  [15, 70],
  [17, COL.OBSERVACIONES],
  [18, COL.ONCOLOGICO],
  [19, COL.ESTIPENDIO],
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
  'USUARIO ONCOLOGICO': 18,
  'POSTULACION ESTIPENDIO': 19,
}

var _SERVICIO_COL_MAP = {
  'EXAMEN DE SANGRE USUARIO': 51,
  'TOMA DE EXAMENES': 51,
  'EXAMEN DE SANGRE CUIDADOR': 24,
  'MORBILIDAD': 50,
  'VISITA DE INGRESO': 13,
  'CONTROL MEDICO': 52,
  'CONTROL DE SALUD (INCLUYE CSV)': 52,
  'PRESION ARTERIAL': 54,
  'CCV MEDICO': 53,
  'CCV ENFERMERIA': 54,
  'CCV NUTRICIONAL': 58,
  'CSCV ENFERMERIA': 54,
  'PODOLOGO': 57,
  'ATENCION PODOLOGICA': 57,
  'NUTRICIONISTA': 58,
  'CONSULTA NUTRICIONAL': 58,
  'FONOAUDIOLOGA': 60,
  'ATENCION FONOAUDIOLOGA': 60,
  'CONTROL KINESICO': 61,
  'CONSULTA KINE': 61,
  'CONSULTA KINESIOLOGO': 61,
  'ODONTOLOGIA': 62,
  'EMPA': 67,
  'EMPAM': 67,
  'PIC-1C': 29,
  'VISITA PIC 1': 29,
  'PIC-2C': 30,
  'VISITA PIC 2': 30,
  'EMPA/EMPAM CUIDADOR': 23,
  'CCV CUIDADOR': 25,
  'ZARIT': 69,
  'APLICACION ZARIT': 69,
  'CAMBIO O INSTALACION DE SONDA': 96,
  'CAMBIO DE SONDA': 96,
  'CAMBIO SONDA': 96,
  'CAMBIO DE SONDA FOLEY': 96,
  'CAMBIO SONDA FOLEY': 96,
  'INSTALACION DE SONDA': 96,
  'INSTALACION SONDA': 96,
  'INSTALACION DE SONDA FOLEY': 96,
  'INSTALACION SONDA FOLEY': 96,
  'SONDA FOLEY': 96,
  'CURACION': 102,
  'CURACION SIMPLE': 102,
  'CURACION AVANZADA': 102,
  'CONSULTA TRABAJADORA SOCIAL': 72,
  'CONSULTA PSICOLOGICA': 71,
  'CAPACITACION KINESIOLOGA': 85,
  'CAPACITACION ENFERMERIA': 86,
  'CAPACITACION TENS': 87,
  'CAPACITACION TRABAJADORA SOCIAL': 89,
  'CAPACITACION PSICOLOGA': 91,
  'CAPACITACION FONOAUDIOLOGA': 92,
}

var _SERVICIO_SI_MAP = {
  'INMUNIZACION INFLUENZA USUARIO': 81,
  'INMUNIZACION NEUMO 23 USUARIO': 82,
  'INMUNIZACION INFLUENZA CUIDADOR': 83,
  'INMUNIZACION NEUMO 23 CUIDADOR': 84,
}

var _PRESTACIONES_LIST = []
var _plSet = {}
function _regPrestacion(n) { if (!_plSet[n]) { _plSet[n] = 1; _PRESTACIONES_LIST.push(n) } }
for (var _pk1 in _SERVICIO_COL_MAP) _regPrestacion(_pk1)
for (var _pk2 in _SERVICIO_SI_MAP) _regPrestacion(_pk2)
_regPrestacion('CONTROL CUIDADOR')
_regPrestacion('CONTROL DE SALUD (INCLUYE CSV)')

var _FORM_VAL_MAP = {
  59: {
    'SOBRE PESO': 'SOBREPESO',
  },
  26: {
    'DEPENDENCIA LEVE': 'LEVE',
    'DEPENDENCIA MODERADA': 'MODERADO',
    'DEPENDENCIA SEVERA': 'SEVERO',
  },
  27: {
    'No': 'NO',
    'no': 'NO',
  },
}

// Normaliza un valor del formulario antes de guardarlo en Pacientes.
// Usa el mapa exacto _FORM_VAL_MAP y, para el Barthel, coincidencias parciales
// del texto "Otros" del formulario (combinaciones CA avanzado / certificados).
function _formValorNormalizado(ppCol, v) {
  var s = String(v == null ? '' : v).trim()
  if (!s) return s
  var map = _FORM_VAL_MAP[ppCol]
  if (map && map[s]) return map[s]
  if (ppCol === 26) {
    var u = s.toUpperCase()
    if (u.indexOf('CA AVANZADO') >= 0) return 'SEVERO + CA'
    if (u.indexOf('CERTIF') >= 0) return 'MOD + CERTIF'
    if (u.indexOf('CA ') >= 0 || u.indexOf(' CA') >= 0) return 'SEVERO + CA'
  }
  return s
}

var HOJA = 'Agenda Profesionales'
var PC = 6
var GC = 1
var GI = [1, 1+PC+GC, 1+2*(PC+GC)]
var GF = 3

// Hojas de sistema: en ellas NUNCA se aplica el formato de agenda (semana).
// Si la hoja activa es una de estas, las funciones de agenda usan HOJA
// ('Agenda Profesionales'); en CUALQUIER otra hoja, el formato se aplica ahí.
var HOJAS_SISTEMA = [
  HOJA_PAC, 'Dashboard', HOJA_FORM, 'Ingresos', 'Parámetros',
  'Referencia Columnas', 'Config', '_PlantillaSemana', '_Resalte', '_Opciones',
]

var S = {
  header: { bg: '#1E293B', fg: '#fff', sz: 12 },
  sub:    { bg: ['#334155','#0F766E','#115E59','#1E293B','#3A5B6E','#B45309'] },
  slots:  {
    abreviadas: { bg: '#FEF3C7', fg: '#92400E' },
    vdi:        { bg: '#DCFCE7', fg: '#166534' },
    colacion:   { bg: '#FCE7F3', fg: '#9D174D' },
    registro:   { bg: '#F1F5F9', fg: '#475569' },
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

// ─── LENGUAJE VISUAL COMÚN (todas las hojas del sistema) ───────────────────
// Tokens compartidos para sus cabeceras principales, tablas de datos y paneles,
// para que el sistema se vea coherente entre hojas manteniendo lo funcional.
var _UI = {
  font:     'Arial',
  border:   '#E2E8F0', // borde de celdas de datos (pizarra claro)
  borderMed:'#CBD5E1', // borde medio (detalles/cabeceras menores)
  rowH:     26,        // alto estándar de filas de datos
  hdrBg:    '#1E293B', // pizarra oscura (cabeceras de hoja / paneles, neutro profesional)
  hdrSub:   '#CBD5E1', // texto secundario sobre pizarra
  accent:   '#0F766E', // teal (acento de marca, no azul)
  accent2:  '#115E59', // teal profundo (barras de sección)
  accentL:  '#99F6E4', // teal claro (bordes de tarjetas)
  tint:     '#F0FDFA', // tinte teal suave (fondos de apoyo)
  ink:      '#1E293B', // texto principal
  inkSub:   '#64748B', // texto secundario
  zebraBg:  ['#FFFFFF', '#F8FAFC'],
  frozenBg: ['#F8FAFC', '#EEF2F7'],
  chipCard: '#F8FAFC', // fondo de etiquetas de tarjetas
  tabBW:    '#99F6E4', // color de pestaña del sistema (teal claro)
  okBg: '#DCFCE7', okFg: '#15803D',
  warnBg: '#FEF3C7', warnFg: '#B45309',
  errBg: '#FEE2E2', errFg: '#B91C1C',
  naBg: '#F1F5F9', naFg: '#64748B',
}
