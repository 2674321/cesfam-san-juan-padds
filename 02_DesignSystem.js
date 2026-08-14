// ════════════════════════════════════════════════════════
// ÍNDICE ▏02_DesignSystem.gs │ PADDS V4 — capa de componentes
//
// Design System sobre _UI (00_Constantes): tokens de componentes
// reutilizables para todas las ventanas HTML de la capa V4.
// Nada de esto toca las hojas: es solo presentación.

// ─── EXTENSIÓN DEL CSS BASE (_uiCss de 01_Utilidades) ──────────────────────
// Devuelve DECLARACIONES CSS (sin <style>) con componentes: tarjetas,
// botones, chips, badges, tabs, acordeones, tablas, estados y empty states.
// Uso: '<style>' + _uiCss() + _uiExtCss() + '...' + '</style>'
function _uiExtCss() {
  return '.btn{display:inline-block;padding:9px 14px;background:#1E293B;color:#fff;border:none;border-radius:6px;font-size:13px;font-family:inherit;cursor:pointer;text-align:center;text-decoration:none}' +
    '.btn:hover{background:#0F172A}.btn:disabled{opacity:.55;cursor:default}' +
    '.btn-sec{background:#fff;color:#1E293B;border:1px solid #CBD5E1}.btn-sec:hover{background:#F1F5F9}' +
    '.btn-sm{padding:6px 10px;font-size:12px;border-radius:5px}' +
    '.btn-teal{background:#0F766E}.btn-teal:hover{background:#115E59}' +
    '.btn-danger{background:#B91C1C}.btn-danger:hover{background:#991B1B}' +
    '.btn-block{display:block;width:100%;margin:6px 0}' +
    '.chip{display:inline-block;padding:2px 8px;border-radius:6px;font-size:11px;font-weight:500;margin:2px 4px 2px 0;background:#F1F5F9;color:#475569}' +
    '.chip-ok{background:#DCFCE7;color:#15803D}.chip-warn{background:#FEF3C7;color:#B45309}' +
    '.chip-err{background:#FEE2E2;color:#B91C1C}.chip-info{background:#E0F2FE;color:#0369A1}' +
    '.chip-neu{background:#F1F5F9;color:#64748B}' +
    '.badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600}' +
    '.b-ok{background:#DCFCE7;color:#15803D}.b-warn{background:#FEF3C7;color:#B45309}' +
    '.b-err{background:#FEE2E2;color:#B91C1C}.b-info{background:#E0F2FE;color:#0369A1}.b-neu{background:#F1F5F9;color:#64748B}' +
    '.card{background:#fff;border:1px solid #E2E8F0;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.06);padding:12px 14px;margin-bottom:12px}' +
    '.card-t{font-size:12px;font-weight:600;color:#475569;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}' +
    '.kv{display:flex;justify-content:space-between;gap:10px;padding:5px 0;border-bottom:1px solid #F1F5F9;font-size:13px}' +
    '.kv:last-child{border-bottom:none}.kv .k{color:#64748B}.kv .v{font-weight:500;text-align:right}' +
    '.grid2{display:grid;grid-template-columns:1fr 1fr;gap:8px}.grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px}' +
    '.num{font-size:22px;font-weight:600;line-height:1.1}' +
    '.tabs{display:flex;gap:4px;margin-bottom:12px;border-bottom:1px solid #E2E8F0}' +
    '.tab{padding:7px 12px;font-size:13px;font-weight:500;color:#64748B;cursor:pointer;border-bottom:2px solid transparent;background:none;border-top:none;border-left:none;border-right:none}' +
    '.tab:hover{color:#0F766E}.tab.active{color:#0F766E;border-bottom-color:#0F766E}' +
    '.acc{border:1px solid #E2E8F0;border-radius:8px;margin-bottom:8px;background:#fff;overflow:hidden}' +
    '.acc-h{padding:9px 12px;font-size:13px;font-weight:600;cursor:pointer;display:flex;justify-content:space-between;align-items:center;user-select:none;color:#1E293B}' +
    '.acc-h:after{content:"\\25BC";font-size:10px;color:#94A3B8}' +
    '.acc.collapsed .acc-h:after{content:"\\25B6"}' +
    '.acc-c{padding:8px 12px 10px;border-top:1px solid #F1F5F9}' +
    '.acc.collapsed .acc-c{display:none}' +
    '.kvrow{display:grid;grid-template-columns:repeat(2,1fr);gap:2px 14px}' +
    '.empty{padding:34px 16px;text-align:center;color:#94A3B8}' +
    '.empty .ic{font-size:30px;margin-bottom:8px}' +
    '.spinner{display:inline-block;width:14px;height:14px;border:2px solid #E2E8F0;border-top-color:#0F766E;border-radius:50%;animation:sp .8s linear infinite;vertical-align:middle}' +
    '@keyframes sp{to{transform:rotate(360deg)}}' +
    'table.tb{width:100%;border-collapse:collapse;font-size:12px}' +
    'table.tb th{text-align:left;padding:5px 8px;color:#64748B;font-weight:600;border-bottom:1px solid #E2E8F0}' +
    'table.tb td{padding:6px 8px;border-bottom:1px solid #F1F5F9;vertical-align:top}' +
    'table.tb tr:hover td{background:#F8FAFC}' +
    '.link{color:#0F766E;cursor:pointer;text-decoration:none}.link:hover{text-decoration:underline}' +
    '.st-ok{color:#15803D}.st-warn{color:#C2410C}.st-err{color:#B91C1C}.st-info{color:#0369A1}.st-neu{color:#64748B}' +
    '.kpi{background:#fff;border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;text-align:center}' +
    '.kpi .n{font-size:22px;font-weight:600;line-height:1.1}.kpi .l{font-size:11px;color:#64748B;margin-top:2px}' +
    '.scroll{overflow-y:auto;max-height:58vh}'
}

// ─── TOKENS DE ESTADO PARA LAS CAPAS DE TAREAS/ALERTAS ─────────────────────
// Semántica común: SUCCESS / WARNING / DANGER / INFO / NEUTRAL.
var _UI_STATE = {
  SUCCESS: { key: 'ok',    cls: 'chip-ok',   badge: 'b-ok',   fg: '#15803D', bg: '#DCFCE7' },
  WARNING: { key: 'warn',  cls: 'chip-warn', badge: 'b-warn', fg: '#C2410C', bg: '#FFEDD5' },
  DANGER:  { key: 'err',   cls: 'chip-err',  badge: 'b-err',  fg: '#B91C1C', bg: '#FEE2E2' },
  INFO:    { key: 'info',  cls: 'chip-info', badge: 'b-info', fg: '#0369A1', bg: '#E0F2FE' },
  NEUTRAL: { key: 'neu',   cls: 'chip-neu',  badge: 'b-neu',  fg: '#64748B', bg: '#F1F5F9' },
}

// Severidad → estado visual (para tareas y alertas).
function _sevEstado(sev) {
  if (sev === 'CRITICA') return _UI_STATE.DANGER
  if (sev === 'IMPORTANTE') return _UI_STATE.WARNING
  if (sev === 'PENDIENTE') return _UI_STATE.INFO
  return _UI_STATE.NEUTRAL
}

// ─── HELPERS DE MARCADO REUTILIZABLES ──────────────────────────────────────

function _uiKpi(n, label, cls) {
  return '<div class="kpi' + (cls ? ' ' + cls : '') + '"><div class="n">' + n + '</div><div class="l">' + label + '</div></div>'
}

function _uiChip(texto, sev) {
  var s = sev ? _sevEstado(sev) : _UI_STATE.NEUTRAL
  return '<span class="chip ' + s.cls + '">' + texto + '</span>'
}

function _uiEmpty(icono, titulo, detalle) {
  return '<div class="empty"><div class="ic">' + (icono || '🔍') + '</div>' +
    '<div style="font-weight:600;color:#475569">' + (titulo || 'Sin información') + '</div>' +
    (detalle ? '<div style="font-size:12px;margin-top:4px">' + detalle + '</div>' : '') + '</div>'
}

function _uiAcc(titulo, contenido, abierto) {
  return '<div class="acc' + (abierto ? '' : ' collapsed') + '">' +
    '<div class="acc-h" onclick="this.parentNode.classList.toggle(\'collapsed\')">' + titulo + '</div>' +
    '<div class="acc-c">' + contenido + '</div></div>'
}
