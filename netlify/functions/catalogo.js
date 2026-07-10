'use strict';

// GET público: devuelve el catálogo (hoja "Detalle por Verdura", solo en stock)
// agrupado por Grupo, más la configuración de precios/extras (hoja "Config").
// Alimenta el formulario de pedidos.

const G = require('./_graph');

const CONFIG_DEFAULTS = {
  PRECIO_BOLSON:    20000,
  PRECIO_ENVIO:     4000,
  PRECIO_SOPERA:    3500,
  PRECIO_ENSALADA:  3500,
  PRECIO_ESCABECHE: 3500,
  SS_SOPERA:        false,
  SS_ENSALADA:      false,
  SS_ESCABECHE:     false,
  WHATSAPP:         '5493584209218',
  // Datos bancarios: se toman de la hoja "Config" del Excel (no se escriben a mano).
  ALIAS:            '',
  BANCO:            '',
  TITULAR:          '',
  CBU:              '',
};

// Claves de texto usadas por el formulario y las variantes aceptadas en la hoja
// "Config" (tolerante a mayúsculas/minúsculas y acentos).
const TEXT_KEY_ALIASES = {
  ALIAS:   ['alias'],
  BANCO:   ['banco'],
  TITULAR: ['titular', 'titular de la cuenta', 'razon social', 'razón social'],
  CBU:     ['cbu', 'cbu/cvu', 'cvu'],
  WHATSAPP:['whatsapp', 'wsp', 'telefono', 'teléfono'],
};

// Localiza el índice de columna por nombre (tolerante a may/min y acentos sueltos).
function findCol(header, ...names) {
  const norm = s => String(s ?? '').trim().toLowerCase();
  const wanted = names.map(norm);
  for (let i = 0; i < header.length; i++) {
    const h = norm(header[i]);
    if (wanted.some(w => h === w || h.startsWith(w))) return i;
  }
  return -1;
}

function parseBool(v, def = true) {
  if (v === null || v === undefined || v === '') return def;
  const s = String(v).trim().toLowerCase();
  if (['false', 'no', '0', 'sin stock', 'sinstock'].includes(s)) return false;
  if (['true', 'si', 'sí', '1', 'stock', 'en stock'].includes(s)) return true;
  return def;
}

// Catálogo de verduras para el formulario. Precio/nombre/unidad y el toggle de
// stock del admin salen de "Detalle por Verdura"; la cantidad cosechada de
// "StockDisponible" y lo ya pedido (pedidasV) de los pedidos de esa cosecha.
// SOLO se incluyen productos disponibles: se OCULTAN los desactivados por el
// admin y los que llegaron al límite (disponible <= 0). Grupos vacíos se omiten.
function buildCatalogo(rows, cosechada, pedidasV) {
  let hIdx = rows.findIndex(r => r.some(c => String(c ?? '').trim().toLowerCase() === 'grupo'));
  if (hIdx === -1) hIdx = 1; // fallback: segunda fila
  const header = rows[hIdx] || [];
  const cN = findCol(header, 'verdura / producto', 'verdura', 'producto');
  const cG = findCol(header, 'grupo');
  const cU = findCol(header, 'unidad');
  const cP = findCol(header, 'precio unit. ($)', 'precio unit', 'precio');
  const cS = findCol(header, 'stock');

  const grupos = [];
  const idx = {};
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const nombre = String(r[cN] ?? '').trim();
    const grupo  = String(r[cG] ?? '').trim();
    if (!nombre || !grupo) continue;
    if (/total/i.test(nombre)) continue;
    const nl = nombre.toLowerCase();
    if (nl === 'bolsón semanal' || nl === 'bolson semanal') continue;           // bolsón aparte
    if (['extras', 'bolsones', 'bolsón', 'bolson'].includes(grupo.toLowerCase())) continue;

    const enStock = cS === -1 ? true : parseBool(r[cS], true);
    const cos     = cosechada[nl] || 0;
    const pedida  = pedidasV[nl] || 0;
    const disp    = Math.max(0, cos - pedida);
    if (!enStock || disp <= 0) continue;                                        // OCULTAR sin stock
    if (!(grupo in idx)) { idx[grupo] = grupos.length; grupos.push({ grupo, items: [] }); }
    grupos[idx[grupo]].items.push({
      nombre,
      unidad: String(r[cU] ?? '').trim(),
      precio: Number(r[cP]) || 0,
      cosechada: cos,
      pedida,
      disponible: disp,
    });
  }
  return grupos;
}

// Nombres (en minúscula) de los extras que maneja el formulario.
const EXTRA_NAMES = ['bandeja sopera', 'bandeja de ensalada', 'escabeche de berenjenas'];

function buildConfig(rows) {
  const cfg = { ...CONFIG_DEFAULTS };
  const norm = s => String(s ?? '').toLowerCase()
    .replace(/[áàäâã]/g, 'a').replace(/[éèëê]/g, 'e').replace(/[íìïî]/g, 'i')
    .replace(/[óòöôõ]/g, 'o').replace(/[úùüû]/g, 'u') // quitar acentos
    .trim();
  for (const r of rows) {
    const k = String(r[0] ?? '').trim();
    if (!k || /clave/i.test(k)) continue;
    const v = r[1];
    if (k.startsWith('SS_')) { cfg[k] = parseBool(v, false); continue; }
    if (k.startsWith('PRECIO_')) { cfg[k] = Number(v) || cfg[k] || 0; continue; }
    // Mapear claves de texto conocidas a su nombre canónico (ALIAS, TITULAR, CBU…).
    const nk = norm(k);
    const canon = Object.keys(TEXT_KEY_ALIASES)
      .find(c => c.toLowerCase() === nk || TEXT_KEY_ALIASES[c].some(a => norm(a) === nk));
    const key = canon || k;
    cfg[key] = (v ?? '') !== '' ? String(v).trim() : cfg[key];
  }
  return cfg;
}

exports.handler = async (event) => {
  const pf = G.preflight(event); if (pf) return pf;
  if (!G.envOk()) return G.json(500, { error: 'Servidor no configurado.' });

  try {
    const token = await G.getToken();
    const [det, conf, stk, ped] = await Promise.all([
      G.readSheet(token, 'Detalle por Verdura'),
      G.readSheet(token, 'Config'),
      G.readSheet(token, 'StockDisponible'),
      G.readSheet(token, 'Pedidos'),
    ]);

    const stock   = G.parseStockDisponible(stk.values || []);
    const pedidas = G.pedidasPorCosecha(ped.values || [], stock.fecha);
    const grupos  = buildCatalogo(det.values || [], stock.cosechada, pedidas.verduras);
    const bolson  = { cosechada: stock.bolson, pedida: pedidas.bolson,
                      disponible: Math.max(0, stock.bolson - pedidas.bolson) };
    const extras  = {};
    for (const n of EXTRA_NAMES) {
      const c = stock.cosechada[n] || 0, p = pedidas.extras[n] || 0;
      extras[n] = { cosechada: c, pedida: p, disponible: Math.max(0, c - p) };
    }
    const config = buildConfig(conf.values || []);

    return G.json(200, { ok: true, catalogo: grupos, bolson, extras, fecha: stock.fecha, config });
  } catch (err) {
    console.error('[catalogo]', err.message);
    return G.json(500, { error: 'Error del servidor: ' + err.message });
  }
};
