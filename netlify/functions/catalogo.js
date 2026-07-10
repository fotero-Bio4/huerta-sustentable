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

// Devuelve { grupos, bolson }. Combina "Detalle por Verdura" (precio, cantidad
// pedida, toggle de stock del admin) con la cosecha manual de "StockDisponible".
// Cada producto lleva: cosechada, pedida, disponible = max(0, cosechada-pedida)
// y sinStock = (admin lo desactivó) || (disponible <= 0). No se oculta ninguno:
// el formulario los muestra atenuados.
function buildCatalogo(rows, cosechada) {
  // Buscar la fila de encabezado (la que contiene "Grupo").
  let hIdx = rows.findIndex(r => r.some(c => String(c ?? '').trim().toLowerCase() === 'grupo'));
  if (hIdx === -1) hIdx = 1; // fallback: segunda fila
  const header = rows[hIdx] || [];
  const cN = findCol(header, 'verdura / producto', 'verdura', 'producto');
  const cG = findCol(header, 'grupo');
  const cU = findCol(header, 'unidad');
  const cP = findCol(header, 'precio unit. ($)', 'precio unit', 'precio');
  const cPed = findCol(header, 'cantidad pedida');
  const cS = findCol(header, 'stock');

  const grupos = [];
  const idx = {};
  let bolson = { cosechada: cosechada['__bolson__'] || 0, pedida: 0, disponible: 0 };
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const nombre = String(r[cN] ?? '').trim();
    const grupo  = String(r[cG] ?? '').trim();
    if (!nombre || !grupo) continue;
    if (/total/i.test(nombre)) continue;
    const nl = nombre.toLowerCase();
    const pedida = cPed === -1 ? 0 : (Number(r[cPed]) || 0);

    // El bolsón se maneja aparte (grupo "Bolsones").
    if (nl === 'bolsón semanal' || nl === 'bolson semanal') {
      bolson.pedida = pedida;
      bolson.disponible = Math.max(0, bolson.cosechada - pedida);
      continue;
    }
    // Extras y otros grupos manejados aparte quedan fuera del catálogo de verduras.
    if (['extras', 'bolsones', 'bolsón', 'bolson'].includes(grupo.toLowerCase())) continue;

    const enStock  = cS === -1 ? true : parseBool(r[cS], true);
    const cos      = cosechada[nl] || 0;
    const disp     = Math.max(0, cos - pedida);
    if (!(grupo in idx)) { idx[grupo] = grupos.length; grupos.push({ grupo, items: [] }); }
    grupos[idx[grupo]].items.push({
      nombre,
      unidad: String(r[cU] ?? '').trim(),
      precio: Number(r[cP]) || 0,
      cosechada: cos,
      pedida,
      disponible: disp,
      sinStock: !enStock || disp <= 0,
    });
  }
  return { grupos, bolson };
}

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
    const [det, conf, stk] = await Promise.all([
      G.readSheet(token, 'Detalle por Verdura'),
      G.readSheet(token, 'Config'),
      G.readSheet(token, 'StockDisponible'),
    ]);

    const stock = G.parseStockDisponible(stk.values || []);
    const cosechada = { ...stock.cosechada, __bolson__: stock.bolson };
    const { grupos, bolson } = buildCatalogo(det.values || [], cosechada);
    const config = buildConfig(conf.values || []);

    return G.json(200, { ok: true, catalogo: grupos, bolson, config });
  } catch (err) {
    console.error('[catalogo]', err.message);
    return G.json(500, { error: 'Error del servidor: ' + err.message });
  }
};
