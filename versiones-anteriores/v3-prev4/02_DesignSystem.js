// ════════════════════════════════════════════════════════
// ÍNDICE ▏02_DesignSystem.gs │ PADDS V4 + mejoras V5 — capa de componentes
//
// Design System sobre _UI (00_Constantes): tokens de componentes
// reutilizables para todas las ventanas HTML de la capa V4.
// Nada de esto toca las hojas: es solo presentación.
//
// V5 (aditivo, no rompe el API V4):
//  - _uiAcc con id/clave, teclado y ARIA (persistencia vía sessionStorage)
//  - botones con feedback (loading → éxito/error → restaurar) y toast
//  - resultado PDF visible + animaciones (fadeIn, flash)
//  - _uiScriptComun(): controlador JS compartido para todas las ventanas

// ─── EXTENSIÓN DEL CSS BASE (_uiCss de 01_Utilidades) ──────────────────────
// Devuelve DECLARACIONES CSS (sin <style>) con componentes: tarjetas,
// botones, chips, badges, tabs, acordeones, tablas, estados y empty states.
// Uso: '<style>' + _uiCss() + _uiExtCss() + '...' + '</style>'
function _uiExtCss() {
  return '.btn-sm{padding:6px 10px;font-size:12px;border-radius:5px}' +
    '.btn-teal{background:#0F766E}.btn-teal:hover{background:#115E59}' +
    '.btn-danger{background:#B91C1C}.btn-danger:hover{background:#991B1B}' +
    '.btn-block{display:block;width:100%;margin:6px 0}' +
    '.chip{display:inline-block;padding:2px 8px;border-radius:6px;font-size:11px;font-weight:500;margin:2px 4px 2px 0;background:#F1F5F9;color:#475569}' +
    '.chip-ok{background:#DCFCE7;color:#15803D}.chip-warn{background:#FFEDD5;color:#C2410C}' +
    '.chip-err{background:#FEE2E2;color:#B91C1C}.chip-info{background:#E0F2FE;color:#0369A1}' +
    '.chip-neu{background:#F1F5F9;color:#64748B}' +
    '.badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600}' +
    '.b-ok{background:#DCFCE7;color:#15803D}.b-warn{background:#FFEDD5;color:#C2410C}' +
    '.b-err{background:#FEE2E2;color:#B91C1C}.b-info{background:#E0F2FE;color:#0369A1}.b-neu{background:#F1F5F9;color:#64748B}' +
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
    '.acc-h:after{content:"\\25BC";font-size:10px;color:#64748B}' +
    '.acc.collapsed .acc-h:after{content:"\\25B6"}' +
    '.acc-c{padding:8px 12px 10px;border-top:1px solid #F1F5F9}' +
    '.acc.collapsed .acc-c{display:none}' +
    '.kvrow{display:grid;grid-template-columns:repeat(2,1fr);gap:2px 14px}' +
    '.empty{padding:34px 16px;text-align:center;color:#64748B}' +
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
    '.scroll{overflow-y:auto;max-height:58vh}' +
    // V5: botones con feedback, toast, resultado PDF, animaciones, acordeón
    '.btn{transition:background .15s ease,transform .05s ease}.btn:active{transform:scale(.97)}' +
    '.btn:focus-visible{outline:2px solid #0F766E;outline-offset:1px}' +
    '.btn-ok{background:#15803D!important;border-color:#15803D!important}.btn-err{background:#B91C1C!important;border-color:#B91C1C!important}' +
    '.btn .spin{display:inline-block;width:12px;height:12px;border:2px solid rgba(255,255,255,.35);border-top-color:#fff;border-radius:50%;animation:sp .7s linear infinite;vertical-align:-2px;margin-right:5px}' +
    '.btn-sec .spin{border-color:#CBD5E1;border-top-color:#1E293B}' +
    '.acc-h:hover{background:#F8FAFC}.acc-h:focus-visible{outline:2px solid #0F766E;outline-offset:-2px}' +
    '.it,.fadein{animation:fadeIn .18s ease}' +
    '.flash{animation:flashB 1.2s ease}' +
    '.toast{position:fixed;left:12px;right:12px;bottom:14px;padding:10px 12px;border-radius:10px;font-size:13px;font-weight:600;color:#fff;box-shadow:0 4px 14px rgba(0,0,0,.25);z-index:99;animation:toastIn .22s ease;opacity:1;transition:opacity .3s}' +
    '.toast.out{opacity:0}' +
    '.toast.ok{background:#15803D}.toast.err{background:#B91C1C}.toast.warn{background:#C2410C}.toast.neu{background:#334155}' +
    '.pdf-result{margin:6px 0;padding:8px 10px;background:#F0FDFA;border:1px solid #99F6E4;border-radius:8px;font-size:12.5px;display:flex;justify-content:space-between;align-items:center;gap:8px;animation:fadeIn .2s ease}' +
    '.pdf-result .pr-t{color:#15803D;font-weight:600}.pdf-result .pr-l{color:#0F766E;font-weight:700;text-decoration:none}.pdf-result .pr-e{color:#64748B;font-size:11.5px}' +
    '@keyframes fadeIn{from{opacity:0;transform:translateY(3px)}to{opacity:1;transform:none}}' +
    '@keyframes toastIn{from{transform:translateY(10px);opacity:0}to{transform:none;opacity:1}}' +
    '@keyframes flashB{0%,75%{background:#CCFBF1;border-color:#0F766E}100%{background:#fff}}' +
    // V5.2: botones grandes, badges de cuenta y barra "paciente actual"
    '.btn-lg{height:auto;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px;padding:12px 8px;line-height:1.25;text-align:center;font-size:13px;font-weight:600}' +
    '.btn-lg .btn-sub{font-size:10.5px;font-weight:500;opacity:.75;display:block}' +
    '.bdg{display:inline-block;min-width:18px;padding:1px 6px;border-radius:999px;font-size:10.5px;font-weight:700;vertical-align:1px}' +
    '.bdg-err{background:#FEE2E2;color:#B91C1C}.bdg-warn{background:#FFEDD5;color:#C2410C}.bdg-ok{background:#DCFCE7;color:#15803D}' +
    // V5.3: transición de pantalla, acordeón suave, selección, dropdown "Más",
    // barra "Paciente actual" discreta e inferior (contexto, no navegación)
    'body{animation:appear .22s ease}' +
    '@keyframes appear{from{opacity:0}to{opacity:1}}' +
    '.acc-c{animation:accIn .18s ease}' +
    '@keyframes accIn{from{opacity:0;transform:translateY(-3px)}to{opacity:1;transform:none}}' +
    '.it{transition:border-color .12s ease,background .12s ease}' +
    '.it:hover{border-color:#99F6E4;background:#F8FAFC}' +
    '.it:active{transform:scale(.995)}' +
    '.it-sel{border-color:#0F766E!important;background:#F0FDFA!important;box-shadow:inset 0 0 0 1px #99F6E4}' +
    '.drop{position:relative}' +
    '.drop-panel{display:none;position:absolute;bottom:calc(100% + 6px);left:0;right:0;background:#fff;border:1px solid #E2E8F0;border-radius:10px;box-shadow:0 8px 24px rgba(0,0,0,.12);padding:6px;z-index:60;animation:accIn .16s ease}' +
    '.drop.open .drop-panel{display:block}' +
    '.drop-panel .btn{width:100%;margin:2px 0}' +
    'body.pacbar-on{padding-bottom:44px}' +
    '.pac-bar{position:fixed;left:0;right:0;bottom:0;z-index:70;display:flex;align-items:center;gap:8px;padding:7px 12px;background:rgba(255,255,255,.94);backdrop-filter:blur(6px);border-top:1px solid #E2E8F0;font-size:11.5px;color:#64748B;box-shadow:0 -2px 10px rgba(0,0,0,.04)}' +
    '.pac-bar .pp-l{font-weight:700;color:#0F766E;white-space:nowrap;font-size:10.5px;text-transform:uppercase;letter-spacing:.4px}' +
    '.pac-bar .pp-n{font-weight:600;color:#334155;flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}' +
    '.pac-bar .pp-r{font-family:Consolas,monospace;color:#64748B}' +
    '.pac-x{border:none;background:none;color:#64748B;font-size:13px;cursor:pointer;padding:2px 5px;border-radius:4px;line-height:1}' +
    '.pac-x:hover{background:#E2E8F0;color:#475569}'
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

var _ACC_SEQ = 0

function _uiKpi(n, label, cls) {
  return '<div class="kpi' + (cls ? ' ' + cls : '') + '"><div class="n">' + n + '</div><div class="l">' + label + '</div></div>'
}

function _uiEmpty(icono, titulo, detalle) {
  return '<div class="empty"><div class="ic">' + (icono || '🔍') + '</div>' +
    '<div style="font-weight:600;color:#475569">' + (titulo || 'Sin información') + '</div>' +
    (detalle ? '<div style="font-size:12px;margin-top:4px">' + detalle + '</div>' : '') + '</div>'
}

// Acordeón con id estable (clave) para persistencia, teclado y ARIA.
// El toggle inline restaura el comportamiento V4 (independiente de scripts);
// _uiAccInit (en _uiScriptComun) solo agrega teclado/persistencia y no
// interfiere con los headers que ya traen onclick. Compatible con 3 args.
function _uiAcc(titulo, contenido, abierto, clave, attrs) {
  _ACC_SEQ++
  var id = 'acc-' + (clave != null ? String(clave).replace(/[^A-Za-z0-9_]/g, '_') : String(_ACC_SEQ))
  var ab = abierto !== false
  var togg = "onclick=\"var p=this.parentNode;var c=p.classList.toggle('collapsed');this.setAttribute('aria-expanded',c?'false':'true');\""
  return '<div class="acc' + (ab ? '' : ' collapsed') + '" id="' + id + '"' + (attrs || '') + '>' +
    '<div class="acc-h" role="button" tabindex="0" aria-expanded="' + (ab ? 'true' : 'false') + '" ' + togg + '>' + titulo + '</div>' +
    '<div class="acc-c">' + contenido + '</div></div>'
}

// ─── CONTROLADOR JS COMPARTIDO (V5) ───────────────────────────────────────
// Se inyecta UNA vez por ventana HTML (en _uiV4Shell o al final del body).
// Proporciona: acordeón (clic + teclado + persistencia), estados de botón
// (loading → éxito/error → restaurar), toast y llamadas con feedback:
// _uiRun / _uiSel / _uiFicha / _uiPdf. No toca las hojas (solo presentación).
function _uiScriptComun() {
  return '<script>' +
    'function _esc(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}' +
    // acordeón: toggle por delegación + teclado + persistencia por id
    'function _uiAccInit(){' +
    'var accsInit=document.querySelectorAll(".acc");if(!accsInit.length)return;' +
    'document.addEventListener("click",function(e){' +
    'var t=e.target;while(t&&t!==document&&!(t.classList&&t.classList.contains("acc-h")))t=t.parentNode;' +
    'if(!t||!t.classList||!t.classList.contains("acc-h")||(t.hasAttribute&&t.hasAttribute("onclick")))return;' +
    'var a=t.parentNode;if(!a||!a.classList.contains("acc"))return;' +
    'var col=!a.classList.contains("collapsed");' +
    'if(col){a.classList.add("collapsed");a.setAttribute("aria-expanded","false")}else{a.classList.remove("collapsed");a.setAttribute("aria-expanded","true")}' +
    'if(a.id){try{sessionStorage.setItem("padds_acc_"+a.id,col?"0":"1")}catch(x){}}' +
    '});' +
    'document.addEventListener("keydown",function(e){' +
    'if((e.key==="Enter"||e.key===" ")&&e.target&&e.target.classList&&e.target.classList.contains("acc-h")){e.preventDefault();e.target.click()}' +
    '});' +
    'var accs=document.querySelectorAll(".acc");' +
    'for(var i=0;i<accs.length;i++){' +
    'var a=accs[i],st=null;' +
    'if(!a.id)continue;' +
    'try{st=sessionStorage.getItem("padds_acc_"+a.id)}catch(x){}' +
    'if(st==="0"&&!a.classList.contains("collapsed")){a.classList.add("collapsed");a.setAttribute("aria-expanded","false")}' +
    'else if(st==="1"&&a.classList.contains("collapsed")){a.classList.remove("collapsed");a.setAttribute("aria-expanded","true")}' +
    '}' +
    '}' +
    // estados de botón (loading / éxito / error / restaurar)
    'function _uiLoading(b,t){if(!b)return;if(!b._orig){b._orig=b.textContent;b._cls=b.className}b.disabled=true;b.innerHTML="<span class=\\"spin\\"></span>"+t}' +
    'function _uiOk(b,t){if(!b)return;b.disabled=false;b.innerHTML=t;b.className=b._cls+" btn-ok"}' +
    'function _uiFail(b,t){if(!b)return;b.disabled=false;b.innerHTML=t;b.className=b._cls+" btn-err"}' +
    'function _uiRestore(b,ms){if(!b)return;setTimeout(function(){b.disabled=false;b.className=b._cls||b.className;b.textContent=b._orig||b.textContent},ms||1500)}' +
    // toast
    'function _uiToast(m,t,ms){' +
    'var d=document.createElement("div");d.className="toast "+(t||"ok");d.textContent=m;document.body.appendChild(d);' +
    'setTimeout(function(){d.className+=" out"},ms||2600);' +
    'setTimeout(function(){if(d.parentNode)d.parentNode.removeChild(d)},(ms||2600)+400)' +
    '}' +
    // llamada genérica con feedback (bloquea el botón hasta terminar)
    'function _uiRun(fn,b,args,opts){' +
    'opts=opts||{};_uiLoading(b,opts.loadTxt||"Procesando…");' +
    'var ok=function(r){' +
    'if(r&&r.error){_uiFail(b,opts.errTxt||"No se pudo completar");_uiToast(r.error,"err",4500);_uiRestore(b,2800);return}' +
    'if(opts.ok){try{opts.ok(r,b)}catch(e){_uiRestore(b,700)}return}' +
    '_uiOk(b,opts.okTxt||"✓ Completado");_uiRestore(b,1500)' +
    '};' +
    'var err=function(e){_uiFail(b,opts.errTxt||"No se pudo completar");_uiToast("No se pudo completar","err",4500);_uiRestore(b,3000)};' +
    'var cmd=google.script.run.withSuccessHandler(ok).withFailureHandler(err)[fn];' +
    'if(args&&args.length)cmd.apply(null,args);else cmd()' +
    '}' +
    // seleccionar: navega en la hoja + resalta la fila + feedback
    'function _uiSel(b,fila,label){' +
    'if(b)_uiLoading(b,"Seleccionando…");' +
    'google.script.run.withSuccessHandler(function(r){' +
    'if(r&&r.error){if(b){_uiFail(b,"No se encontró");_uiRestore(b,2400)}_uiToast(r.error,"err",4000);return}' +
    'if(b){_uiOk(b,"✓ Seleccionado");_uiRestore(b,1300)}' +
    'var el=b?b.closest(".it,.kv,.card"):null;' +
    'if(el){el.classList.remove("flash");void el.offsetWidth;el.classList.add("flash");' +
    'if(el.classList.contains("it")&&el.scrollIntoView)el.scrollIntoView({behavior:"smooth",block:"center"})}' +
    '_uiToast("✓ "+(label||"Paciente")+" seleccionado · fila "+fila,"ok")' +
    '}).withFailureHandler(function(e){' +
    'if(b){_uiFail(b,"No se pudo seleccionar");_uiRestore(b,2600)}' +
    '_uiToast("No se pudo completar","err",4500)' +
    '})._navegarAFila(fila)' +
    '}' +
    // abrir la ficha del paciente con feedback
    'function _uiFicha(b,fila){' +
    '_uiLoading(b,"Abriendo ficha…");' +
    'google.script.run.withSuccessHandler(function(r){' +
    'if(r&&r.error){_uiFail(b,"No se pudo abrir");_uiToast(r.error,"err",4500);_uiRestore(b,2800);return}' +
    '_uiOk(b,"✓ Ficha abierta");_uiRestore(b,1500);_uiToast("✓ Ficha del paciente abierta (fila "+fila+")","ok")' +
    '}).withFailureHandler(function(e){' +
    '_uiFail(b,"No se pudo abrir");_uiToast("No se pudo completar","err",4500);_uiRestore(b,2800)' +
    '})._abrirFichaV4(fila)' +
    '}' +
    // generar PDF con feedback y resultado visible (enlace)
    'function _uiPdf(b,fila){' +
    '_uiLoading(b,"Generando PDF…");' +
    'google.script.run.withSuccessHandler(function(r){' +
    'if(r&&r.error){_uiFail(b,"No se pudo generar");_uiToast(r.error,"err",5000);_uiRestore(b,3000);return}' +
    '_uiOk(b,"✓ PDF generado");_uiRestore(b,1800);' +
    'var u=(r&&r.url)||(r&&r.download);' +
    'if(u){try{window.open(u,"_blank")}catch(x){}}' +
    'var box=document.createElement("div");box.className="pdf-result";' +
    'box.innerHTML="<span class=\\"pr-t\\">✓ PDF generado</span>"+(u?"<a class=\\"pr-l\\" href=\\""+u.replace(/&/g,"&amp;")+"\\" target=\\"_blank\\">Abrir PDF ↗</a>":"<span class=\\"pr-e\\">Sin enlace</span>");' +
    'if(b&&b.parentNode)b.parentNode.insertBefore(box,b.nextSibling)' +
    '}).withFailureHandler(function(e){' +
    '_uiFail(b,"No se pudo generar");_uiToast("No se pudo generar el PDF","err",5000);_uiRestore(b,3000)' +
    '})._generarFichaPdfSidebar(fila)' +
    '}' +
    // V5.2: paciente actual (sessionStorage) + barra persistente en cada ventana
    'function _uiPkey(){var k=null;try{k=sessionStorage.getItem("padds_pac")}catch(x){}return k||""}' +
    'function _uiPac(b,o){' +
    'if(!o)return;' +
    'var rec={f:o.f,n:String(o.n||""),ru:String(o.ru||""),s:String(o.s||""),e:String(o.e||""),x:String(o.x||""),ed:String(o.ed||"")};' +
    'try{sessionStorage.setItem("padds_pac",JSON.stringify(rec))}catch(x){}' +
    'var bb=document.getElementById("pacbar");' +
    'if(bb){var nn=bb.querySelector(".pp-n"),rr=bb.querySelector(".pp-r");' +
    'if(nn)nn.textContent=rec.n;if(rr)rr.textContent=rec.ru||""}' +
    '}' +
    'function _uiPacF(b,mode){' +
    'var f=b.getAttribute("data-f"),n=b.getAttribute("data-n");' +
    'if(f){try{sessionStorage.setItem("padds_pac",JSON.stringify({f:Number(f),n:decodeURIComponent(n||""),ru:"",s:"",e:"",x:"",ed:""}))}catch(x){}}' +
    'if(mode==="ficha"){_uiFicha(b,Number(f))}else{_uiSel(b,Number(f),decodeURIComponent(n||""))}' +
    '}' +
    'function _uiLimpiarPac(){' +
    'try{sessionStorage.removeItem("padds_pac")}catch(x){}' +
    'var bb=document.getElementById("pacbar");if(bb&&bb.parentNode)bb.parentNode.removeChild(bb);' +
    'var bd=document.body;if(bd&&bd.className&&bd.className.indexOf("pacbar-on")>=0)bd.className=bd.className.replace("pacbar-on","");' +
    '_uiToast("Paciente actual limpiado","neu",2000)' +
    '}' +
    'function _uiPacBar(){' +
    'var c=_uiPkey();if(!c)return;var o=null;try{o=JSON.parse(c)}catch(x){}' +
    'if(!o||!o.f)return;' +
    'var h="<div class=\\"pac-bar\\" id=\\"pacbar\\">" +' +
    '"<span class=\\"pp-l\\">Paciente actual</span>" +' +
    '"<span class=\\"pp-n\\">"+_esc(o.n||"")+"</span>" +' +
    '"<span class=\\"pp-r\\">"+_esc(o.ru||"")+"</span>" +' +
    '"<button class=\\"pac-x\\" title=\\"Limpiar\\" aria-label=\\"Limpiar paciente actual\\" onclick=\\"_uiLimpiarPac()\\">✕</button></div>";' +
    'var bd=document.body;' +
    'if(bd){var cl=bd.className||"";if(cl.indexOf("pacbar-on")<0)bd.className=cl+" pacbar-on";' +
    'if(bd.insertAdjacentHTML)bd.insertAdjacentHTML("beforeend",h)}' +
    '}' +
    // dropdown "Más": alterna el panel, cierra al clic fuera o con Esc (a11y)
    'function _uiDropInit(){' +
    'if(!document.querySelector("[data-drop]"))return;' +
    'document.addEventListener("click",function(e){' +
    'var t=e.target,btn=null;' +
    'while(t&&t!==document){if(t.hasAttribute&&t.hasAttribute("data-drop")){btn=t;break}t=t.parentNode}' +
    'if(btn){' +
    'var id=btn.getAttribute("data-drop");var d=document.getElementById(id);' +
    'var others=document.querySelectorAll(".drop.open");' +
    'for(var i=0;i<others.length;i++){if(others[i]!==d)others[i].classList.remove("open")}' +
    'if(d){var op=d.classList.toggle("open");btn.setAttribute("aria-expanded",op?"true":"false")}return}' +
    'var dd=document.querySelectorAll(".drop.open");' +
    'for(var j=0;j<dd.length;j++){if(!dd[j].contains(e.target))dd[j].classList.remove("open")}' +
    '});' +
    'document.addEventListener("keydown",function(e){' +
    'if(e.key==="Escape"){var dd=document.querySelectorAll(".drop.open");for(var k=0;k<dd.length;k++)dd[k].classList.remove("open")}' +
    '});' +
    '}' +
    'function _uiInitComun(){_uiAccInit();_uiPacBar();_uiDropInit()}' +
    '_uiInitComun();' +
    '</script>'
}
