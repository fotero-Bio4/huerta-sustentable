'use strict';

// POST con login (hoja "Usuarios"): devuelve la lista de pedidos, estadísticas,
// y el Detalle por Verdura / Cosecha Resumen calculados a partir de los pedidos
// reales (suma cantidades, más preciso que las fórmulas SEARCH del Excel).

const G = require('./_graph');

// Índices de columna en la hoja Pedidos (0-indexed, A=0 … S=18).
const COL = {
  nombre: 0, tel: 1, email: 2, barrio: 3, dir: 4, entrega: 5,
  bolsones: 6, precioBolson: 7, detalle: 8, cantItems: 9, totalVerd: 10,
  notas: 11, extras: 12, totalExtras: 13, pago: 14, total: 15,
  estado: 16, fecha: 17, id: 18,
};

function num(v) { const n = Number(v); return Number.isFinite(n) ? n : 0; }

function buildPedidos(rows, firstRow) {
  // Header de datos = fila cuyo col A == "Nombre".
  let hIdx = rows.findIndex(r => String(r[0] ?? '').trim().toLowerCase() === 'nombre');
  if (hIdx === -1) hIdx = 3;
  const pedidos = [];
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const nombre = String(r[COL.nombre] ?? '').trim();
    if (!nombre) continue;
    pedidos.push({
      excelRow:    firstRow + i,
      nombre,
      tel:         String(r[COL.tel] ?? '').trim(),
      email:       String(r[COL.email] ?? '').trim(),
      barrio:      String(r[COL.barrio] ?? '').trim(),
      dir:         String(r[COL.dir] ?? '').trim(),
      entrega:     String(r[COL.entrega] ?? '').trim(),
      bolsones:    num(r[COL.bolsones]),
      precioBolson:num(r[COL.precioBolson]),
      detalle:     String(r[COL.detalle] ?? '').trim(),
      totalVerd:   num(r[COL.totalVerd]),
      notas:       String(r[COL.notas] ?? '').trim(),
      extras:      String(r[COL.extras] ?? '').trim(),
      totalExtras: num(r[COL.totalExtras]),
      pago:        String(r[COL.pago] ?? '').trim(),
      total:       num(r[COL.total]),
      estado:      G.normEstado(r[COL.estado]),
      fechaISO:    G.anyToISO(r[COL.fecha]),
      id:          String(r[COL.id] ?? '').trim(),
    });
  }
  return pedidos;
}

function parseBool(v, def = true) {
  if (v === null || v === undefined || v === '') return def;
  const s = String(v).trim().toLowerCase();
  if (['false','no','0','sin stock'].includes(s)) return false;
  if (['true','si','sí','1','stock'].includes(s)) return true;
  return def;
}

// Catálogo completo desde "Detalle por Verdura": índice por nombre + lista con
// stock (toggle admin), cantidad cosechada (de StockDisponible), pedida (de los
// pedidos de la cosecha activa) y disponible.
function catalogFull(rows, cosechada, pedidasV) {
  cosechada = cosechada || {};
  pedidasV = pedidasV || {};
  let hIdx = rows.findIndex(r => r.some(c => String(c ?? '').trim().toLowerCase() === 'grupo'));
  if (hIdx === -1) hIdx = 1;
  const header = (rows[hIdx] || []).map(c => String(c ?? '').trim().toLowerCase());
  const cS = header.findIndex(h => h === 'stock');
  const index = {};
  const lista = [];
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const nombre = String(r[0] ?? '').trim();
    const grupo  = String(r[1] ?? '').trim();
    if (!nombre || !grupo || /total/i.test(nombre)) continue;
    if (['extras', 'bolsones', 'bolsón', 'bolson'].includes(grupo.toLowerCase())) continue;
    const cos    = cosechada[nombre.toLowerCase()] || 0;
    const pedida = pedidasV[nombre.toLowerCase()] || 0;
    const item = { nombre, grupo, unidad: String(r[2] ?? '').trim(), precio: num(r[3]),
      stock: cS === -1 ? true : parseBool(r[cS], true),
      cosechada: cos, pedida, disponible: Math.max(0, cos - pedida) };
    index[nombre.toLowerCase()] = item;
    lista.push(item);
  }
  return { index, lista };
}

function buildConfig(rows) {
  const cfg = {};
  for (const r of rows) {
    const k = String(r[0] ?? '').trim();
    if (!k || /clave/i.test(k)) continue;
    const v = r[1];
    if (k.startsWith('SS_')) cfg[k] = parseBool(v, false);
    else if (k.startsWith('PRECIO_')) cfg[k] = num(v);
    else cfg[k] = v;
  }
  return cfg;
}

// Calcula Detalle por Verdura (suma cantidades) y Cosecha (bolsones+extras).
function computeCosecha(pedidos, catIdx) {
  const verduras = {};   // nombre → { nombre, grupo, unidad, cantidad }
  const extras   = {};   // nombre → cantidad
  let bolsones = 0;

  for (const p of pedidos) {
    if (p.estado === 'cancelado') continue;
    bolsones += p.bolsones;
    for (const it of G.parseDetalle(p.detalle)) {
      const key = it.nombre.toLowerCase();
      const meta = catIdx[key] || { nombre: it.nombre, grupo: 'Otros', unidad: it.unidad };
      if (!verduras[key]) verduras[key] = { nombre: meta.nombre, grupo: meta.grupo, unidad: meta.unidad, cantidad: 0 };
      verduras[key].cantidad += it.cantidad;
    }
    for (const e of G.parseDetalle(p.extras)) {
      extras[e.nombre] = (extras[e.nombre] || 0) + e.cantidad;
    }
  }
  return {
    bolsones,
    verduras: Object.values(verduras).filter(v => v.cantidad > 0)
      .sort((a, b) => a.grupo.localeCompare(b.grupo) || a.nombre.localeCompare(b.nombre)),
    extras: Object.entries(extras).filter(([, c]) => c > 0).map(([nombre, cantidad]) => ({ nombre, cantidad })),
  };
}

exports.handler = async (event) => {
  const pf = G.preflight(event); if (pf) return pf;
  if (event.httpMethod !== 'POST') return G.json(405, { error: 'Método no permitido' });
  if (!G.envOk())                  return G.json(500, { error: 'Servidor no configurado.' });

  let body;
  try { body = JSON.parse(event.body || '{}'); }
  catch { return G.json(400, { error: 'Cuerpo inválido.' }); }

  try {
    const token = await G.getToken();
    const user = await G.validateUser(token, body.mail, body.pss);
    if (!user) return G.json(401, { error: 'Mail o contraseña incorrectos.' });

    const [ped, det, conf, stk] = await Promise.all([
      G.readSheet(token, 'Pedidos'),
      G.readSheet(token, 'Detalle por Verdura'),
      G.readSheet(token, 'Config'),
      G.readSheet(token, 'StockDisponible'),
    ]);

    const stock   = G.parseStockDisponible(stk.values || []);
    const pedidas = G.pedidasPorCosecha(ped.values || [], stock.fecha);
    const pedidos = buildPedidos(ped.values || [], ped.firstRow);
    const cat     = catalogFull(det.values || [], stock.cosechada, pedidas.verduras);
    const config  = buildConfig(conf.values || []);
    const cosecha = computeCosecha(pedidos, cat.index);

    return G.json(200, { ok: true, nombre: user.nombre, pedidos, cosecha, catalogo: cat.lista, config, fechaCosecha: stock.fecha });
  } catch (err) {
    console.error('[admin]', err.message);
    return G.json(500, { error: 'Error del servidor: ' + err.message });
  }
};
