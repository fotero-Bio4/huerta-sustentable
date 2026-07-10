'use strict';

// Formulario de pedidos. Catálogo y precios vienen de /.netlify/functions/catalogo
// (hoja Excel en SharePoint). Al confirmar, se guarda vía /.netlify/functions/pedido
// y se ofrece confirmar por WhatsApp.

const GRUPO_EMOJI = {
  'Hortalizas de hoja':'🥬', 'Hortalizas de fruto':'🍅', 'Zapallos':'🎃',
  'Tallos y hojas':'🌿', 'Aromáticas':'🌱',
};
const EXTRAS_DEF = [
  { key:'sopera',    nombre:'Bandeja sopera',          desc:'Mix de verduras listas para sopa o caldo', cfgPrecio:'PRECIO_SOPERA',    cfgSS:'SS_SOPERA' },
  { key:'ensalada',  nombre:'Bandeja de ensalada',     desc:'Elegís el tipo abajo',                     cfgPrecio:'PRECIO_ENSALADA',  cfgSS:'SS_ENSALADA' },
  { key:'escabeche', nombre:'Escabeche de berenjenas', desc:'Berenjenas en escabeche artesanal',        cfgPrecio:'PRECIO_ESCABECHE', cfgSS:'SS_ESCABECHE' },
];
const TIPOS_ENSALADA = [
  'Achicoria + huevo', 'Zanahoria + repollo + huevo', 'Lechuga morada + zanahoria',
  'Repollo + zanahoria + cherry', 'Lechuga crespa + apio + rabanito',
];

let CONFIG = {}, CATALOG = [], BOLSON = { cosechada:0, pedida:0, disponible:0 };
let cantVerd = [];                 // cantVerd[gi][ii]
let cantBolson = 0;
let extQty = { sopera:0, ensalada:0, escabeche:0 };
let pagoSel = null, modo = 'bolsones', modoEntrega = 'envio', tieneDir = false;

const $ = id => document.getElementById(id);
const fmt = n => '$' + Number(n).toLocaleString('es-AR');

// ── Carga inicial ───────────────────────────────────────────────────────────
async function cargarCatalogo() {
  try {
    const r = await fetch('/.netlify/functions/catalogo', { cache: 'no-store' });
    const d = await r.json();
    if (!r.ok || !d.ok) throw new Error(d.error || 'No se pudo cargar el catálogo.');
    CONFIG = d.config || {};
    CATALOG = d.catalogo || [];
    BOLSON = d.bolson || { cosechada:0, pedida:0, disponible:0 };
    cantVerd = CATALOG.map(g => g.items.map(() => 0));
    renderTodo();
    $('loaderCatalogo').style.display = 'none';
    $('formBody').style.display = 'block';
  } catch (err) {
    $('loaderCatalogo').style.display = 'none';
    const e = $('errorCatalogo');
    e.style.display = 'block';
    e.textContent = '⚠️ ' + err.message + ' Probá recargar la página en unos minutos.';
  }
}

// Bloque de datos bancarios para transferencia (todo desde la hoja "Config" del Excel).
function datosBancariosHTML() {
  const linea = (icon, label, val) => val ? `<div>${icon} ${label}: <strong>${val}</strong></div>` : '';
  const conCopia = (icon, label, val, fn) => val
    ? `<div>${icon} ${label}: <strong>${val}</strong> <button type="button" class="btn-copiar-alias" onclick="${fn}(this)">📋 Copiar</button></div>`
    : '';
  return `📎 Una vez confirmado el pedido, transferís y enviás el comprobante por WhatsApp.
    ${linea('🏦', 'Banco', CONFIG.BANCO)}
    ${linea('👤', 'Titular', CONFIG.TITULAR)}
    ${conCopia('🔢', 'CBU', CONFIG.CBU, 'copiarCBU')}
    ${conCopia('📲', 'Alias', CONFIG.ALIAS, 'copiarAlias')}`;
}

// Copia texto al portapapeles (con fallback para navegadores viejos / sin permiso).
function copiarTexto(btn, valor) {
  if (!valor) return;
  const ok = () => {
    const prev = btn.innerHTML;
    btn.innerHTML = '✓ Copiado';
    btn.classList.add('copiado');
    setTimeout(() => { btn.innerHTML = prev; btn.classList.remove('copiado'); }, 1800);
  };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(valor).then(ok).catch(() => copiarFallback(valor, ok));
  } else {
    copiarFallback(valor, ok);
  }
}
function copiarFallback(valor, cb) {
  const ta = document.createElement('textarea');
  ta.value = valor; ta.style.position = 'fixed'; ta.style.opacity = '0';
  document.body.appendChild(ta); ta.focus(); ta.select();
  try { document.execCommand('copy'); cb && cb(); } catch (e) {}
  document.body.removeChild(ta);
}
function copiarAlias(btn) { copiarTexto(btn, CONFIG.ALIAS); }
function copiarCBU(btn)   { copiarTexto(btn, CONFIG.CBU); }

function renderTodo() {
  $('envio-monto').textContent = fmt(CONFIG.PRECIO_ENVIO || 0);
  $('bolson-sub').textContent = '';
  $('pago-alias-detalle').textContent = 'Alias: ' + (CONFIG.ALIAS || '—') + (CONFIG.BANCO ? ' — ' + CONFIG.BANCO : '');
  $('info-transferencia').innerHTML = datosBancariosHTML();
  renderGrupos();
  renderExtras();
  renderEnsaladaRows();
  renderBolsonStock();
  // Único medio de pago: Transferencia. Se deja preseleccionada.
  const pc = document.querySelector('.pago-card');
  if (pc) selPago(pc, 'Transferencia');
}

// Estado de stock del bolsón: si no hay disponible, deshabilita el control.
function renderBolsonStock() {
  const box = $('bolson-box'), desc = $('bolson-desc'), ctrl = $('bolson-ctrl');
  if (!box) return;
  const disp = BOLSON.disponible || 0;
  if (disp <= 0) {
    cantBolson = 0; $('bolson-qty').textContent = 0; $('bolson-sub').textContent = '';
    box.classList.add('agotado');
    desc.innerHTML = '<span class="stock-badge">Sin stock</span> Por el momento no hay bolsones disponibles';
    ctrl.style.pointerEvents = 'none'; ctrl.style.opacity = '.4';
  } else {
    box.classList.remove('agotado');
    desc.innerHTML = `Surtido de verduras de estación según disponibilidad <span class="stock-hint">Quedan ${disp} disponibles</span>`;
    ctrl.style.pointerEvents = ''; ctrl.style.opacity = '';
  }
}

function renderGrupos() {
  $('gruposContainer').innerHTML = CATALOG.map((g, gi) => `
    <div class="grupo-titulo">${GRUPO_EMOJI[g.grupo]||'🥗'} ${g.grupo}</div>
    ${g.items.map((p, ii) => {
      const hint = p.sinStock
        ? '<span class="stock-badge">Sin stock</span>'
        : `<span class="stock-hint">Quedan ${p.disponible}</span>`;
      return `
      <div class="producto-row${p.sinStock?' agotado':''}">
        <span class="producto-nombre">${p.nombre} <span style="color:#aaa;font-size:.78rem;">(${p.unidad})</span> ${hint}</span>
        <span class="precio-ref">${fmt(p.precio)}</span>
        <div class="qty-control">
          <button class="qty-btn" onclick="chVerd(${gi},${ii},-1)">−</button>
          <span class="qty-num" id="v-q-${gi}-${ii}">0</span>
          <button class="qty-btn" onclick="chVerd(${gi},${ii},1)">+</button>
        </div>
        <span class="prod-sub" id="v-s-${gi}-${ii}"></span>
      </div>`;
    }).join('')}
  `).join('') || '<p style="color:var(--texto-suave);font-size:.88rem">No hay productos disponibles en este momento.</p>';
}

function renderExtras() {
  $('extrasContainer').innerHTML = EXTRAS_DEF.map(e => {
    const sinStock = !!CONFIG[e.cfgSS];
    const precio = CONFIG[e.cfgPrecio] || 0;
    return `
      <div class="extra-card" style="${sinStock?'opacity:.55':''}">
        <div class="extra-header">
          <div class="extra-info">
            <div class="extra-nombre">${e.nombre} ${sinStock?'<span class="stock-badge">Sin stock</span>':''}</div>
            <div class="extra-desc">${sinStock?'Sin stock por el momento':e.desc+' — '+fmt(precio)+' c/u'}</div>
          </div>
          <div class="qty-control" style="${sinStock?'pointer-events:none;opacity:.4':''}">
            <button class="qty-btn" onclick="chExtra('${e.key}',-1)">−</button>
            <span class="qty-num" id="${e.key}-qty">0</span>
            <button class="qty-btn" onclick="chExtra('${e.key}',1)">+</button>
          </div>
          <span class="extra-subtotal" id="${e.key}-sub"></span>
        </div>
      </div>`;
  }).join('');
}

function renderEnsaladaRows() {
  $('ensalada-rows').innerHTML = TIPOS_ENSALADA.map(t =>
    `<button type="button" class="qty-btn" style="width:auto;border-radius:14px;padding:.25rem .7rem;font-size:.78rem;font-weight:500;margin:.15rem;display:inline-flex" onclick="agregarTipoEns('${t.replace(/'/g,"\\'")}')">+ ${t}</button>`
  ).join('');
}
function agregarTipoEns(t) {
  const ta = $('ensalada-libre');
  ta.value = ta.value ? ta.value + ' / ' + t : t;
}

// ── Interacciones ─────────────────────────────────────────────────────────────
function selEntrega(tipo) {
  modoEntrega = tipo;
  document.querySelectorAll('.etab').forEach(e=>e.classList.remove('active'));
  document.querySelectorAll('.entrega-panel').forEach(p=>p.classList.remove('show'));
  $('etab-'+tipo).classList.add('active');
  $('panel-'+tipo).classList.add('show');
  calcResumen();
}
function onDireccion(val){ tieneDir = val.trim().length>0; $('envio-aviso').style.display = tieneDir?'flex':'none'; calcResumen(); }
function chBolson(d){
  const max = BOLSON.disponible || 0;
  cantBolson = Math.min(max, Math.max(0, cantBolson+d));
  $('bolson-qty').textContent = cantBolson;
  $('bolson-sub').textContent = cantBolson>0 ? fmt(cantBolson*(CONFIG.PRECIO_BOLSON||0)) : '';
  calcResumen();
}
function chVerd(gi,ii,d){
  const p = CATALOG[gi].items[ii];
  if (p.sinStock) return;
  cantVerd[gi][ii] = Math.min(p.disponible, Math.max(0, cantVerd[gi][ii]+d));
  $(`v-q-${gi}-${ii}`).textContent = cantVerd[gi][ii];
  const sub = cantVerd[gi][ii]*p.precio;
  $(`v-s-${gi}-${ii}`).textContent = sub>0?fmt(sub):'';
  calcResumen();
}
function chExtra(key,d){
  const def = EXTRAS_DEF.find(e=>e.key===key);
  if (CONFIG[def.cfgSS]) return;
  extQty[key] = Math.max(0, extQty[key]+d);
  $(`${key}-qty`).textContent = extQty[key];
  const sub = extQty[key]*(CONFIG[def.cfgPrecio]||0);
  $(`${key}-sub`).textContent = extQty[key]>0?fmt(sub):'';
  if (key==='ensalada') $('card-ensalada-panel').style.display = extQty.ensalada>0?'block':'none';
  calcResumen();
}
function selPago(el,nombre){
  document.querySelectorAll('.pago-card').forEach(c=>c.classList.remove('selected'));
  el.classList.add('selected'); pagoSel=nombre;
  $('info-transferencia').classList.toggle('show', nombre==='Transferencia');
}
function cambiarTab(btn,tab){
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c=>c.classList.remove('active'));
  btn.classList.add('active'); $('tab-'+tab).classList.add('active'); modo=tab; calcResumen();
}

// ── Recolección de datos ───────────────────────────────────────────────────────
function getItems() {
  const items = [];
  CATALOG.forEach((g,gi)=>g.items.forEach((p,ii)=>{ if(cantVerd[gi][ii]>0) items.push({nombre:p.nombre,cantidad:cantVerd[gi][ii],unidad:p.unidad,precio:p.precio}); }));
  return items;
}
function getExtras() {
  return EXTRAS_DEF.filter(e=>extQty[e.key]>0 && !CONFIG[e.cfgSS])
    .map(e=>({nombre:e.nombre,cantidad:extQty[e.key],precio:CONFIG[e.cfgPrecio]||0}));
}

// ── Resumen ─────────────────────────────────────────────────────────────────────
function calcResumen() {
  const box=$('resumenBox'); let items=[], total=0;
  if (modo==='bolsones' && cantBolson>0){ const s=cantBolson*(CONFIG.PRECIO_BOLSON||0); items.push({n:`Bolsón semanal ×${cantBolson}`,p:s}); total+=s; }
  else if (modo==='armalo'){
    let st=0;
    getItems().forEach(it=>{ const s=it.cantidad*it.precio; items.push({n:`${it.nombre} ×${it.cantidad} ${it.unidad}`,p:s}); st+=s; total+=s; });
    $('total-armado-val').textContent=fmt(st);
  }
  getExtras().forEach(e=>{ const s=e.cantidad*e.precio; items.push({n:`${e.nombre} ×${e.cantidad}`,p:s}); total+=s; });
  const conEnvio = modoEntrega==='envio' && tieneDir;
  if (conEnvio){ items.push({n:'🛵 Envío a domicilio',p:CONFIG.PRECIO_ENVIO||0,envio:true}); total+=CONFIG.PRECIO_ENVIO||0; }
  if (!items.length){ box.style.display='none'; return; }
  box.style.display='block';
  $('resumenItems').innerHTML = items.map(it=>`<div class="resumen-row${it.envio?' envio':''}"><span>${it.n}</span><span>${fmt(it.p)}</span></div>`).join('');
  $('resumenTotal').textContent=fmt(total);
}

// ── Confirmar ─────────────────────────────────────────────────────────────────
async function confirmarPedido() {
  const errEl=$('errorEnvio'); errEl.style.display='none';
  const nombre=$('nombre').value.trim(), telefono=$('telefono').value.trim();
  const email=$('email').value.trim(), barrio=$('barrio').value;
  const direccion = modoEntrega==='envio' ? $('direccion').value.trim() : 'Retiro en huerta';
  const fail = m => { errEl.textContent='⚠️ '+m; errEl.style.display='block'; errEl.scrollIntoView({behavior:'smooth',block:'center'}); };

  if (!nombre||!telefono) return fail('Completá tu nombre y teléfono.');
  if (!barrio) return fail('Seleccioná tu barrio.');
  if (!pagoSel) return fail('Elegí una forma de pago.');
  if (modoEntrega==='envio' && !$('direccion').value.trim()) return fail('Ingresá tu dirección de entrega.');

  const bolsones = modo==='bolsones' ? cantBolson : 0;
  const items = modo==='armalo' ? getItems() : [];
  const extras = getExtras();
  if (bolsones===0 && !items.length && !extras.length) return fail('Agregá al menos un producto, bolsón o extra.');

  // Notas (incluye preferencias de ensalada y aclaraciones)
  const notasModo = modo==='bolsones' ? $('notas-bolson').value.trim() : $('notas-armado').value.trim();
  const notasExtras = $('notas-extras').value.trim();
  const ensLibre = $('ensalada-libre').value.trim();
  const notas = [notasModo, ensLibre?('Ensalada: '+ensLibre):'', notasExtras].filter(Boolean).join(' — ');

  const payload = { nombre, telefono, email, barrio, direccion, entrega: modoEntrega==='envio'?'Envío':'Retiro',
    bolsones, precioBolson: CONFIG.PRECIO_BOLSON||0, items, extras, notas, pago: pagoSel };

  const btn=$('btnConfirmar'); btn.disabled=true; btn.innerHTML='<span class="spinner"></span> Enviando…';
  try {
    const r = await fetch('/.netlify/functions/pedido',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    const d = await r.json();
    if (!r.ok || !d.ok) throw new Error(d.error || 'No se pudo enviar el pedido.');
    armarMensajeWhatsApp(payload);
    $('mainForm').style.display='none';
    $('successScreen').classList.add('show');
    let txt = `Tu pedido fue enviado correctamente${email?` y te llega una copia al confirmar`:''}. Nos comunicamos por WhatsApp para confirmar el pedido y el precio final.`;
    if (pagoSel==='Transferencia') {
      txt += ` Cuando esté confirmado, transferís al alias <strong>${CONFIG.ALIAS||'—'}</strong> y mandás el comprobante.`;
      if (CONFIG.ALIAS) txt += `<br><button type="button" class="btn-copiar-alias" onclick="copiarAlias(this)">📋 Copiar alias</button>`;
    }
    $('successMsg').innerHTML = txt;
    window.scrollTo({top:0,behavior:'smooth'});
  } catch (err) {
    fail(err.message);
  } finally { btn.disabled=false; btn.innerHTML='Enviar pedido ✓'; }
}

function armarMensajeWhatsApp(p) {
  let total = 0, detalle;
  if (p.bolsones>0){ detalle=`Bolsón semanal ×${p.bolsones}`; total += p.bolsones*p.precioBolson; }
  else { detalle = p.items.map(it=>{ total += it.cantidad*it.precio; return `${it.nombre} ×${it.cantidad} ${it.unidad}`; }).join('\n   '); }
  let extLinea='';
  p.extras.forEach(e=>{ total += e.cantidad*e.precio; extLinea += `\n   ${e.nombre} ×${e.cantidad} (${fmt(e.cantidad*e.precio)})`; });
  const entregaLinea = p.entrega==='Envío'
    ? `📍 Envío a: ${p.direccion}, Barrio ${p.barrio} (+${fmt(CONFIG.PRECIO_ENVIO||0)} envío)`
    : `📍 Retiro en huerta — Ruta 8 km 609, Río Cuarto`;
  if (p.entrega==='Envío') total += CONFIG.PRECIO_ENVIO||0;
  const msg = `¡Hola! Quiero hacer un pedido 🌿\n\n👤 *${p.nombre}*\n📱 ${p.telefono}${p.email?'\n📧 '+p.email:''}\n${entregaLinea}\n📅 Entrega: Viernes\n\n🛒 *Pedido:*\n   ${detalle}${extLinea?'\n\n🍱 *Extras:*'+extLinea:''}${p.notas?'\n\n📝 '+p.notas:''}\n\n💰 *Total estimado:* ${fmt(total)}\n💳 *Pago:* ${p.pago}`;
  window._waMsg = encodeURIComponent(msg);
}
function abrirWhatsApp(){ window.open(`https://wa.me/${CONFIG.WHATSAPP||'5493584209218'}?text=${window._waMsg||''}`,'_blank'); }

cargarCatalogo();
