'use strict';

// Panel admin. Login contra hoja Usuarios; datos desde /.netlify/functions/admin.
// Acciones: cambiar estado/precio (/estado), guardar precios y stock (/config),
// exportar PDV por rango de fechas (/export-pdv).

const $ = id => document.getElementById(id);
const fmt = n => '$' + Number(n || 0).toLocaleString('es-AR');

let session = { mail:'', pss:'', nombre:'' };
let DATA = { pedidos:[], cosecha:{verduras:[],extras:[],bolsones:0}, catalogo:[], config:{} };
let filtro = 'todos', selRow = null;
let seleccion = new Set();   // excelRows seleccionadas para cambio masivo

// ── LOGIN ───────────────────────────────────────────────────────────────────
async function login() {
  const mail = $('inMail').value.trim(), pss = $('inPss').value.trim();
  const msg = $('msgLogin'); msg.classList.add('hidden');
  if (!mail || !pss) { msg.textContent='Ingresá mail y contraseña.'; msg.classList.remove('hidden'); return; }
  const btn = $('btnLogin'); btn.disabled=true; btn.innerHTML='<span class="spinner"></span> Verificando…';
  try {
    const d = await postAdmin({});
    if (d.__authError) { msg.textContent = d.error; msg.classList.remove('hidden'); return; }
    session = { mail, pss, nombre: d.nombre };
    aplicarData(d);
    $('screenLogin').classList.add('hidden');
    $('screenApp').classList.remove('hidden');
    $('btnSalir').classList.remove('hidden');
    $('hdrBadge').textContent = 'Hola, ' + d.nombre;
  } finally { btn.disabled=false; btn.textContent='Ingresar'; }

  async function postAdmin() {
    const r = await fetch('/.netlify/functions/admin',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({mail,pss})});
    const d = await r.json();
    if (!r.ok || !d.ok) return { __authError:true, error: d.error || 'Error al ingresar.' };
    return d;
  }
}
function logout(){ session={mail:'',pss:'',nombre:''}; $('inPss').value=''; $('screenApp').classList.add('hidden'); $('btnSalir').classList.add('hidden'); $('hdrBadge').textContent=''; $('screenLogin').classList.remove('hidden'); }

async function recargar() {
  const r = await fetch('/.netlify/functions/admin',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({mail:session.mail,pss:session.pss})});
  const d = await r.json();
  if (r.ok && d.ok) aplicarData(d);
}
function aplicarData(d){ DATA = { pedidos:d.pedidos||[], cosecha:d.cosecha||{verduras:[],extras:[],bolsones:0}, catalogo:d.catalogo||[], config:d.config||{}, fechaCosecha:d.fechaCosecha||'', cosechas:d.cosechas||[] };
  poblarCosechaSelects();
  renderPedidos(); renderDetalle(); renderCosecha(); renderConfig(); stats(); }

// Rellena los <select> de cosecha (Pedidos/Detalle/Cosecha) con las fechas guardadas.
function fmtFecha(iso){ if(!iso) return ''; const [y,m,d]=iso.split('-'); return `${d}/${m}/${y}`; }
function poblarCosechaSelects(){
  const opts = '<option value="">Todas las cosechas</option>' +
    (DATA.cosechas||[]).map(f=>`<option value="${f}">${fmtFecha(f)}</option>`).join('');
  ['pedCosecha','detCosecha','cosCosecha'].forEach(id=>{ const s=$(id); if(s){ const v=s.value; s.innerHTML=opts; s.value=v; } });
}

// ── NAV ───────────────────────────────────────────────────────────────────────
function tab(btn, name){
  document.querySelectorAll('.nav button').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
  btn.classList.add('active'); $('panel-'+name).classList.add('active');
}

// ── PEDIDOS ─────────────────────────────────────────────────────────────────────
function stats(){
  const activos = DATA.pedidos.filter(p=>p.estado!=='cancelado');
  $('stHoy').textContent = DATA.pedidos.length;
  $('stPend').textContent = DATA.pedidos.filter(p=>p.estado==='pendiente').length;
  $('stTotal').textContent = fmt(activos.reduce((s,p)=>s+(p.total||0),0));
}
function setFiltro(btn,f){ document.querySelectorAll('#panel-pedidos .filtro-btn').forEach(b=>b.classList.remove('active')); btn.classList.add('active'); filtro=f; renderPedidos(); }

const ESTADO_LBL = { pendiente:'⏳ Pendiente de Aprobación', confirmado:'✅ Confirmado', pagado:'💰 Pagado', cancelado:'❌ Cancelado' };
function estadoBadge(e){
  return `<span class="badge badge-${e}">${ESTADO_LBL[e]||e}</span>`;
}
// Predicado de filtro de la lista de pedidos: estado + búsqueda + fecha + cosecha.
function pedidoPasaFiltro(p){
  const q = ($('buscar').value||'').toLowerCase();
  const desde=$('pedDesde').value, hasta=$('pedHasta').value, cosecha=$('pedCosecha').value;
  const mf = filtro==='todos' || p.estado===filtro;
  const mb = p.nombre.toLowerCase().includes(q) || (p.tel||'').includes(q) || (p.barrio||'').toLowerCase().includes(q);
  const md = (!desde || (p.fechaISO && p.fechaISO>=desde)) && (!hasta || (p.fechaISO && p.fechaISO<=hasta));
  const mc = !cosecha || p.fechaCosechaISO===cosecha;
  return mf && mb && md && mc;
}
function renderPedidos(){
  const lista = DATA.pedidos.filter(pedidoPasaFiltro);
  const cont = $('listaPedidos');
  if (!lista.length){ cont.innerHTML='<div class="empty"><div>📭</div><p>No hay pedidos.</p></div>'; updateBulkCount(); return; }
  cont.innerHTML = lista.map(p=>{
    const cuerpo = [
      p.bolsones>0 ? `Bolsón semanal ×${p.bolsones}` : '',
      p.detalle, p.extras, p.notas ? '📝 '+p.notas : ''
    ].filter(Boolean).join('\n');
    const wa = `https://wa.me/549${(p.tel||'').replace(/\D/g,'')}?text=${encodeURIComponent('¡Hola '+p.nombre.split(' ')[0]+'! 🌿 Te escribo por tu pedido de Huerta Sustentable.')}`;
    return `<div class="pedido-card">
      <div class="pedido-header">
        <div style="display:flex;gap:.6rem;align-items:flex-start">
          <input type="checkbox" class="sel-check" ${seleccion.has(p.excelRow)?'checked':''} onclick="toggleSel(${p.excelRow},this.checked)" title="Seleccionar">
          <div>
            <div class="pedido-nombre">${p.nombre}</div>
            <div class="pedido-meta">📱 ${p.tel}${p.email?' · 📧 '+p.email:''}</div>
            <div class="pedido-meta">📍 ${p.dir||'-'}${p.barrio?' — '+p.barrio:''} · ${p.fechaISO||''}</div>
          </div>
        </div>
        <div style="text-align:right">
          <div class="pedido-precio">${p.total>0?fmt(p.total):'a confirmar'}</div>
          <button class="mini" style="margin-top:4px" onclick="abrirPrecio(${p.excelRow})">✏️ precio</button>
        </div>
      </div>
      <div class="pedido-body">🛒 ${cuerpo||'(sin detalle)'}</div>
      <div class="pedido-footer">
        ${estadoBadge(p.estado)}
        ${p.pago?`<span class="badge badge-sec">💳 ${p.pago}</span>`:''}
        ${p.entrega?`<span class="badge badge-sec">${p.entrega==='Retiro'?'🌿 Retiro':'🛵 Envío'}</span>`:''}
        <a class="btn-wa" href="${wa}" target="_blank">💬</a>
        <button class="mini" style="margin-left:auto" onclick="abrirEstado(${p.excelRow})">Cambiar estado</button>
      </div>
    </div>`;
  }).join('');
  updateBulkCount();
}

// ── SELECCIÓN MASIVA ──────────────────────────────────────────────────────────
function toggleSel(row,on){ if(on) seleccion.add(row); else seleccion.delete(row); updateBulkCount(); }
function toggleSelAll(cb){
  DATA.pedidos.filter(pedidoPasaFiltro).forEach(p=>{
    if(cb.checked) seleccion.add(p.excelRow); else seleccion.delete(p.excelRow);
  });
  renderPedidos();
}
function updateBulkCount(){ const el=$('bulkCount'); if(el) el.textContent = seleccion.size ? seleccion.size+' seleccionado(s)' : ''; }
async function aplicarEstadoMasivo(){
  const est = $('bulkEstado').value;
  if (!est){ alert('Elegí un estado.'); return; }
  if (!seleccion.size){ alert('Seleccioná al menos un pedido.'); return; }
  const rows = [...seleccion];
  const r = await fetch('/.netlify/functions/estado',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({mail:session.mail,pss:session.pss,excelRows:rows,estado:est})});
  const d = await r.json();
  if (!r.ok || !d.ok){ alert(d.error||'No se pudo guardar.'); return; }
  rows.forEach(row=>{ const p=DATA.pedidos.find(x=>x.excelRow===row); if(p)p.estado=est; });
  seleccion.clear(); $('bulkEstado').value=''; const sa=$('selAll'); if(sa)sa.checked=false;
  renderPedidos(); stats();
}

// Modales estado/precio
function abrirEstado(row){ selRow=row; $('ovEstado').classList.add('show'); }
function abrirPrecio(row){ selRow=row; const p=DATA.pedidos.find(x=>x.excelRow===row); $('inPrecio').value=p&&p.total?p.total:''; $('ovPrecio').classList.add('show'); }
function cerrar(id){ $(id).classList.remove('show'); }
async function setEstado(est){
  await accionPedido({ estado: est });
  const p=DATA.pedidos.find(x=>x.excelRow===selRow); if(p)p.estado=est;
  cerrar('ovEstado'); renderPedidos(); stats();
}
async function setPrecio(){
  const total = parseFloat($('inPrecio').value)||0;
  await accionPedido({ total });
  const p=DATA.pedidos.find(x=>x.excelRow===selRow); if(p)p.total=total;
  cerrar('ovPrecio'); renderPedidos(); stats();
}
async function accionPedido(extra){
  const r = await fetch('/.netlify/functions/estado',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({mail:session.mail,pss:session.pss,excelRow:selRow,...extra})});
  const d = await r.json();
  if (!r.ok || !d.ok) alert(d.error || 'No se pudo guardar.');
}

// ── CÁLCULO DE COSECHA (client-side, filtrable por fecha) ────────────────────────
// Parsea "Lechuga x2 planta | Mizuna x1 atado" → [{nombre,cantidad,unidad}]
function parseDetalle(text){
  if(!text) return [];
  return String(text).split(/\s*[|/]\s*/).map(s=>s.trim()).filter(Boolean).map(seg=>{
    const m=seg.match(/^(.*?)\s*[x×]\s*(\d+(?:[.,]\d+)?)\s*(.*)$/i);
    if(!m) return {nombre:seg,cantidad:1,unidad:''};
    return {nombre:m[1].trim(),cantidad:Number(String(m[2]).replace(',','.')),unidad:(m[3]||'').trim()};
  });
}
function catIndex(){ const idx={}; (DATA.catalogo||[]).forEach(v=>{ idx[v.nombre.toLowerCase()]={nombre:v.nombre,grupo:v.grupo,unidad:v.unidad}; }); return idx; }
function pedidosEnRango(desde,hasta,cosecha){
  return DATA.pedidos.filter(p=>{
    if(desde && (!p.fechaISO || p.fechaISO<desde)) return false;
    if(hasta && (!p.fechaISO || p.fechaISO>hasta)) return false;
    if(cosecha && p.fechaCosechaISO!==cosecha) return false;
    return true;
  });
}
function computeCosecha(desde,hasta,cosecha){
  const catIdx=catIndex(), verduras={}, extras={}; let bolsones=0;
  for(const p of pedidosEnRango(desde,hasta,cosecha)){
    if(p.estado==='cancelado') continue;
    bolsones += p.bolsones||0;
    for(const it of parseDetalle(p.detalle)){
      const key=it.nombre.toLowerCase();
      const meta=catIdx[key]||{nombre:it.nombre,grupo:'Otros',unidad:it.unidad};
      if(!verduras[key]) verduras[key]={nombre:meta.nombre,grupo:meta.grupo,unidad:meta.unidad,cantidad:0};
      verduras[key].cantidad+=it.cantidad;
    }
    for(const e of parseDetalle(p.extras)) extras[e.nombre]=(extras[e.nombre]||0)+e.cantidad;
  }
  return { bolsones,
    verduras:Object.values(verduras).filter(v=>v.cantidad>0).sort((a,b)=>a.grupo.localeCompare(b.grupo)||a.nombre.localeCompare(b.nombre)),
    extras:Object.entries(extras).filter(([,c])=>c>0).map(([nombre,cantidad])=>({nombre,cantidad})) };
}
function limpiarFecha(pref){
  $(pref+'Desde').value=''; $(pref+'Hasta').value='';
  const sel=$(pref+'Cosecha'); if(sel) sel.value='';
  if(pref==='ped') renderPedidos(); else if(pref==='det') renderDetalle(); else renderCosecha();
}

// ── DETALLE POR VERDURA ──────────────────────────────────────────────────────────
function renderDetalle(){
  const v = computeCosecha($('detDesde').value, $('detHasta').value, $('detCosecha').value).verduras;
  if (!v.length){ $('tablaDetalle').innerHTML='<div class="empty"><div>🌱</div><p>Sin verduras pedidas en el rango.</p></div>'; return; }
  let html='<table><thead><tr><th>Verdura</th><th>Grupo</th><th>Unidad</th><th class="num">Cantidad</th></tr></thead><tbody>';
  v.forEach(x=>{ html+=`<tr><td>${x.nombre}</td><td>${x.grupo}</td><td>${x.unidad||''}</td><td class="num"><strong>${x.cantidad}</strong></td></tr>`; });
  html+='</tbody></table>';
  $('tablaDetalle').innerHTML=html;
}

// ── COSECHA ────────────────────────────────────────────────────────────────────
function renderCosecha(){
  const c = computeCosecha($('cosDesde').value, $('cosHasta').value, $('cosCosecha').value);
  let html = `<div class="stat-pill" style="display:inline-block;margin-bottom:1rem">🧺 Bolsones semanales: <strong>${c.bolsones||0}</strong></div>`;
  const grupos = {};
  (c.verduras||[]).forEach(v=>{ (grupos[v.grupo]=grupos[v.grupo]||[]).push(v); });
  Object.keys(grupos).forEach(g=>{
    html += `<h3 class="grupo-h">${g}</h3><table><tbody>`;
    grupos[g].forEach(v=>{ html+=`<tr><td>${v.nombre}</td><td class="num"><strong>${v.cantidad}</strong> ${v.unidad||''}</td></tr>`; });
    html += '</tbody></table>';
  });
  if (c.extras && c.extras.length){
    html += `<h3 class="grupo-h">🍱 Extras</h3><table><tbody>`;
    c.extras.forEach(e=>{ html+=`<tr><td>${e.nombre}</td><td class="num"><strong>${e.cantidad}</strong></td></tr>`; });
    html += '</tbody></table>';
  }
  if (!(c.verduras||[]).length && !(c.extras||[]).length && !c.bolsones) html += '<div class="empty"><div>🧺</div><p>Sin datos de cosecha.</p></div>';
  $('resumenCosecha').innerHTML = html;
}

// ── IMPRESIÓN (Detalle / Cosecha) ─────────────────────────────────────────────────
function imprimir(tipo){
  const desde=$(tipo+'Desde').value, hasta=$(tipo+'Hasta').value, cosecha=$(tipo+'Cosecha').value;
  const c = computeCosecha(desde, hasta, cosecha);
  const filtroTxt = [ cosecha?'Cosecha del '+fmtFecha(cosecha):'', desde?'Desde '+fmtFecha(desde):'', hasta?'Hasta '+fmtFecha(hasta):'' ]
    .filter(Boolean).join(' · ') || 'Todos los pedidos';
  let titulo, cuerpo;
  if (tipo==='det'){
    titulo='Detalle por verdura — a cosechar';
    cuerpo = c.verduras.length
      ? '<table><thead><tr><th>Verdura</th><th>Grupo</th><th>Unidad</th><th class="num">Cantidad</th></tr></thead><tbody>'
        + c.verduras.map(x=>`<tr><td>${x.nombre}</td><td>${x.grupo}</td><td>${x.unidad||''}</td><td class="num">${x.cantidad}</td></tr>`).join('')
        + '</tbody></table>'
      : '<p>Sin verduras para cosechar.</p>';
  } else {
    titulo='Cosecha — resumen';
    cuerpo = `<p><strong>🧺 Bolsones semanales: ${c.bolsones||0}</strong></p>`;
    const grupos={}; c.verduras.forEach(v=>{ (grupos[v.grupo]=grupos[v.grupo]||[]).push(v); });
    Object.keys(grupos).forEach(g=>{
      cuerpo += `<h3>${g}</h3><table><tbody>`
        + grupos[g].map(v=>`<tr><td>${v.nombre}</td><td class="num">${v.cantidad} ${v.unidad||''}</td></tr>`).join('')
        + '</tbody></table>';
    });
    if (c.extras.length) cuerpo += '<h3>🍱 Extras</h3><table><tbody>'
      + c.extras.map(e=>`<tr><td>${e.nombre}</td><td class="num">${e.cantidad}</td></tr>`).join('') + '</tbody></table>';
    if (!c.verduras.length && !c.extras.length && !c.bolsones) cuerpo='<p>Sin datos de cosecha.</p>';
  }
  const html = `<!doctype html><html lang="es"><head><meta charset="utf-8"><title>${titulo}</title>
    <style>body{font-family:Arial,Helvetica,sans-serif;padding:1.5rem;color:#1a1a1a}h1{font-size:1.3rem;color:#2d5a27;margin-bottom:.2rem}h3{color:#2d5a27;margin:1rem 0 .3rem}.sub{color:#555;font-size:.9rem;margin-bottom:1rem}table{width:100%;border-collapse:collapse;margin-bottom:.5rem}th,td{border:1px solid #ccc;padding:.4rem .6rem;text-align:left;font-size:.9rem}th{background:#e8f5e2}.num{text-align:right}</style>
    </head><body><h1>🌱 Huerta Sustentable — ${titulo}</h1>
    <div class="sub">${filtroTxt}</div>${cuerpo}
    <script>window.onload=function(){window.print();}<\/script></body></html>`;
  const w = window.open('', '_blank');
  if (!w){ alert('Permití las ventanas emergentes para imprimir.'); return; }
  w.document.write(html); w.document.close();
}

// ── CONFIG (precios / stock) ─────────────────────────────────────────────────────
const CFG_FIELDS = [
  ['PRECIO_BOLSON','Bolsón semanal ($)'], ['PRECIO_ENVIO','Costo de envío ($)'],
  ['PRECIO_SOPERA','Bandeja sopera ($)'], ['PRECIO_ENSALADA','Bandeja de ensalada ($)'],
  ['PRECIO_ESCABECHE','Escabeche berenjenas ($)'],
];
const SS_FIELDS = [['SS_SOPERA','Bandeja sopera'],['SS_ENSALADA','Bandeja de ensalada'],['SS_ESCABECHE','Escabeche de berenjenas']];

function renderConfig(){
  const cfg = DATA.config||{};
  if ($('fechaCosecha')) $('fechaCosecha').value = DATA.fechaCosecha || '';
  $('cfgGrid').innerHTML = CFG_FIELDS.map(([k,l])=>`<div class="cfg-item"><label>${l}</label><input type="number" id="cfg-${k}" value="${cfg[k]??''}"></div>`).join('')
    + SS_FIELDS.map(([k,l])=>`<div class="cfg-item" style="display:flex;align-items:center;gap:.6rem;align-self:end"><button class="toggle ${cfg[k]?'off':'on'}" id="cfg-${k}" data-on="${cfg[k]?0:1}" onclick="togCfgSS('${k}')"></button><span style="font-size:.82rem">${l} (con stock)</span></div>`).join('');

  $('cfgVerduras').innerHTML = (DATA.catalogo||[]).map((v,i)=>{
    const disp = v.disponible ?? 0;
    const dispTxt = disp<=0
      ? '<span style="color:#c0392b;font-weight:600">agotado</span>'
      : `disp: <strong>${disp}</strong>`;
    return `
    <div class="vrow">
      <span class="vn">${v.nombre} <small>(${v.grupo} · ${v.unidad||''}) · cosech: ${v.cosechada??0} · ${dispTxt}</small></span>
      <input type="number" id="vp-${i}" value="${v.precio}" min="0">
      <button class="toggle ${v.stock?'on':'off'}" id="vs-${i}" data-on="${v.stock?1:0}" onclick="togVerdStock(${i})" title="Habilitar/deshabilitar en el formulario"></button>
    </div>`;
  }).join('') || '<p style="color:var(--texto-suave);font-size:.85rem">Sin verduras en el catálogo.</p>';
}
function togCfgSS(k){ const b=$('cfg-'+k); const on=b.dataset.on==='1'?0:1; b.dataset.on=on; b.classList.toggle('on',!!on); b.classList.toggle('off',!on); }
function togVerdStock(i){ const b=$('vs-'+i); const on=b.dataset.on==='1'?0:1; b.dataset.on=on; b.classList.toggle('on',!!on); b.classList.toggle('off',!on); }

async function guardarConfig(){
  const config = {};
  CFG_FIELDS.forEach(([k])=>{ const val=parseInt($('cfg-'+k).value); if(Number.isFinite(val)) config[k]=val; });
  SS_FIELDS.forEach(([k])=>{ config[k] = $('cfg-'+k).dataset.on!=='1'; }); // toggle "on"=con stock → SS=false
  const precios = (DATA.catalogo||[]).map((v,i)=>({nombre:v.nombre, precio:parseInt($('vp-'+i).value)||v.precio}));
  const stock   = (DATA.catalogo||[]).map((v,i)=>({nombre:v.nombre, enStock: $('vs-'+i).dataset.on==='1'}));

  const btn=$('btnGuardarCfg'); btn.disabled=true; btn.innerHTML='<span class="spinner"></span> Guardando…';
  try {
    const r = await fetch('/.netlify/functions/config',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({mail:session.mail,pss:session.pss,config,precios,stock})});
    const d = await r.json();
    if (!r.ok || !d.ok) { alert(d.error||'No se pudo guardar.'); return; }
    // reflejar local
    DATA.config = {...DATA.config, ...config};
    DATA.catalogo.forEach((v,i)=>{ v.precio=precios[i].precio; v.stock=stock[i].enStock; });
    const m=$('msgCfg'); m.classList.remove('hidden'); setTimeout(()=>m.classList.add('hidden'),2500);
  } finally { btn.disabled=false; btn.innerHTML='💾 Guardar cambios'; }
}

async function guardarFechaCosecha(){
  const fechaCosecha = $('fechaCosecha').value;
  if (!fechaCosecha){ alert('Elegí una fecha.'); return; }
  const btn=$('btnFecha'); btn.disabled=true; btn.innerHTML='<span class="spinner"></span> Guardando…';
  try {
    const r = await fetch('/.netlify/functions/config',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({mail:session.mail,pss:session.pss,fechaCosecha})});
    const d = await r.json();
    if (!r.ok || !d.ok){ alert(d.error||'No se pudo guardar.'); return; }
    DATA.fechaCosecha = fechaCosecha;
    const m=$('msgFecha'); m.classList.remove('hidden'); setTimeout(()=>m.classList.add('hidden'),2500);
    await recargar();  // recalcula disponibles con la nueva cosecha
  } finally { btn.disabled=false; btn.textContent='Guardar fecha'; }
}

// ── EXPORTAR PDV ─────────────────────────────────────────────────────────────────
async function exportarPDV(){
  const desde=$('pdvDesde').value, hasta=$('pdvHasta').value;
  const msg=$('msgPdv'); msg.className='msg hidden';
  if (!desde || !hasta){ msg.className='msg error'; msg.textContent='Elegí fecha desde y hasta.'; return; }
  const btn=$('btnPdv'); btn.disabled=true; btn.innerHTML='<span class="spinner"></span> Generando…';
  try {
    const r = await fetch('/.netlify/functions/export-pdv',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({mail:session.mail,pss:session.pss,desde,hasta})});
    const d = await r.json();
    if (!r.ok || !d.ok){ msg.className='msg error'; msg.textContent='⚠️ '+(d.error||'No se pudo generar.'); return; }
    descargarBase64(d.base64, d.filename, 'application/vnd.ms-excel');
    msg.className='msg ok';
    msg.innerHTML = `✓ Generado: <strong>${d.pedidos}</strong> pedidos, <strong>${d.filas}</strong> filas.` + (d.warnings?` ⚠️ ${d.warnings} ítems sin código de producto (revisá la hoja Mapeo PDV).`:'');
  } finally { btn.disabled=false; btn.textContent='Generar .xls'; }
}
function descargarBase64(b64, filename, mime){
  const bin = atob(b64); const arr = new Uint8Array(bin.length);
  for (let i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i);
  const blob = new Blob([arr], {type:mime});
  const a = document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename;
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(a.href);
}

// Enter para login
$('inPss').addEventListener('keydown', e=>{ if(e.key==='Enter') login(); });
