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
  ALIAS:            'mi.perra.venecia',
  BANCO:            'Banco Nación',
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

function buildCatalogo(rows) {
  // Buscar la fila de encabezado (la que contiene "Grupo").
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
    // Ignorar filas de totales y los grupos manejados aparte (extras y bolsón).
    if (/total/i.test(nombre)) continue;
    if (['extras', 'bolsones', 'bolsón', 'bolson'].includes(grupo.toLowerCase())) continue;
    const enStock = cS === -1 ? true : parseBool(r[cS], true);
    if (!enStock) continue;
    if (!(grupo in idx)) { idx[grupo] = grupos.length; grupos.push({ grupo, items: [] }); }
    grupos[idx[grupo]].items.push({
      nombre,
      unidad: String(r[cU] ?? '').trim(),
      precio: Number(r[cP]) || 0,
    });
  }
  return grupos;
}

function buildConfig(rows) {
  const cfg = { ...CONFIG_DEFAULTS };
  for (const r of rows) {
    const k = String(r[0] ?? '').trim();
    if (!k || /clave/i.test(k)) continue;
    const v = r[1];
    if (k.startsWith('SS_')) cfg[k] = parseBool(v, false);
    else if (k.startsWith('PRECIO_')) cfg[k] = Number(v) || cfg[k] || 0;
    else cfg[k] = v ?? cfg[k];
  }
  return cfg;
}

exports.handler = async (event) => {
  const pf = G.preflight(event); if (pf) return pf;
  if (!G.envOk()) return G.json(500, { error: 'Servidor no configurado.' });

  try {
    const token = await G.getToken();
    const [det, conf] = await Promise.all([
      G.readSheet(token, 'Detalle por Verdura'),
      G.readSheet(token, 'Config'),
    ]);

    const catalogo = buildCatalogo(det.values || []);
    const config   = buildConfig(conf.values || []);

    return G.json(200, { ok: true, catalogo, config });
  } catch (err) {
    console.error('[catalogo]', err.message);
    return G.json(500, { error: 'Error del servidor: ' + err.message });
  }
};
