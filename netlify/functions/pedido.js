'use strict';

// POST público (rate-limited): agrega un pedido nuevo a la hoja "Pedidos".
// Escribe columnas A–O (la P = Total tiene fórmula y se deja intacta) y las
// columnas de estado de la app Q=Estado, R=FechaPedido, S=IDPedido.

const G = require('./_graph');

function num(v) { const n = Number(v); return Number.isFinite(n) ? n : 0; }
function clean(v) { return String(v ?? '').trim().slice(0, 500); }

// Mapa nombre→cantidad pedida (col E "Cantidad pedida") desde Detalle por Verdura.
function pedidaIndex(rows) {
  let hIdx = rows.findIndex(r => r.some(c => String(c ?? '').trim().toLowerCase() === 'grupo'));
  if (hIdx === -1) hIdx = 1;
  const header = (rows[hIdx] || []).map(c => String(c ?? '').trim().toLowerCase());
  let cPed = header.findIndex(h => h === 'cantidad pedida' || h.startsWith('cantidad pedida'));
  if (cPed === -1) cPed = 4;
  const map = {};
  let bolson = 0;
  for (let i = hIdx + 1; i < rows.length; i++) {
    const nombre = String(rows[i][0] ?? '').trim().toLowerCase();
    if (!nombre || /total/i.test(nombre)) continue;
    const ped = num(rows[i][cPed]);
    if (nombre === 'bolsón semanal' || nombre === 'bolson semanal') bolson = ped;
    else map[nombre] = ped;
  }
  return { map, bolson };
}

// Arma el texto de extras: "Bandeja sopera x1 | Escabeche de berenjenas x2"
function formatExtras(extras) {
  return (extras || [])
    .filter(e => e && e.nombre && num(e.cantidad) > 0)
    .map(e => `${e.nombre} x${num(e.cantidad)}`)
    .join(' | ');
}

exports.handler = async (event) => {
  const pf = G.preflight(event); if (pf) return pf;
  if (event.httpMethod !== 'POST') return G.json(405, { error: 'Método no permitido' });
  if (!G.envOk())                  return G.json(500, { error: 'Servidor no configurado.' });

  if (G.checkRate(G.clientIp(event), 12, 15))
    return G.json(429, { error: 'Demasiados pedidos seguidos. Esperá unos minutos.' });

  let body;
  try { body = JSON.parse(event.body || '{}'); }
  catch { return G.json(400, { error: 'Cuerpo inválido.' }); }

  const nombre   = clean(body.nombre);
  const telefono = clean(body.telefono);
  if (!nombre || !telefono) return G.json(400, { error: 'Nombre y teléfono son obligatorios.' });

  const entrega   = body.entrega === 'Retiro' ? 'Retiro' : 'Envío';
  const items     = Array.isArray(body.items)  ? body.items  : [];
  const extras    = Array.isArray(body.extras) ? body.extras : [];
  const bolsones  = num(body.bolsones);
  const hasContenido = bolsones > 0 || items.length > 0 || extras.length > 0;
  if (!hasContenido) return G.json(400, { error: 'El pedido está vacío.' });

  // Cálculos
  const totalVerduras = items.reduce((s, it) => s + num(it.cantidad) * num(it.precio), 0);
  const totalExtras   = extras.reduce((s, e) => s + num(e.cantidad) * num(e.precio), 0);
  const detalleTxt    = G.formatDetalle(items);
  const extrasTxt     = formatExtras(extras);

  try {
    const token = await G.getToken();

    // Leer pedidos (para la próxima fila) + stock cosechado y cantidades ya
    // pedidas (para validar que no se pase del disponible).
    const [ped, stk, det] = await Promise.all([
      G.readSheet(token, 'Pedidos'),
      G.readSheet(token, 'StockDisponible'),
      G.readSheet(token, 'Detalle por Verdura'),
    ]);
    const { values: rows, firstRow } = ped;
    const nextRow = G.nextRowFromRows(rows, firstRow, 0);

    // ── Validación de stock (verduras + bolsón) ──────────────────────────────
    const stock = G.parseStockDisponible(stk.values || []);
    const { map: pedidaMap, bolson: bolsonPedida } = pedidaIndex(det.values || []);
    const dispDe = (nombre) => {
      const nl = String(nombre).trim().toLowerCase();
      return Math.max(0, (stock.cosechada[nl] || 0) - (pedidaMap[nl] || 0));
    };
    const faltantes = [];
    for (const it of items) {
      if (num(it.cantidad) > dispDe(it.nombre)) faltantes.push(it.nombre);
    }
    const bolsonDisp = Math.max(0, stock.bolson - bolsonPedida);
    if (bolsones > bolsonDisp) faltantes.push('Bolsón semanal');
    if (faltantes.length) {
      return G.json(409, { error: 'Nos quedamos sin stock suficiente de: ' + faltantes.join(', ') +
        '. Actualizá la página para ver la disponibilidad y ajustá tu pedido.' });
    }

    const id = 'HS' + Date.now().toString(36).toUpperCase();
    const fechaISO = new Date(new Date().getTime() - new Date().getTimezoneOffset() * 60000)
      .toISOString().slice(0, 10);

    // A..O (15 columnas). NO se escribe P (fórmula de total).
    const rowAO = [
      nombre,                          // A Nombre
      telefono,                        // B Teléfono
      clean(body.email),               // C Email
      clean(body.barrio),              // D Barrio
      clean(body.direccion),           // E Dirección / Retiro
      entrega,                         // F Tipo entrega
      bolsones,                        // G Cant. bolsones
      num(body.precioBolson),          // H Precio bolsón ($)
      detalleTxt,                      // I Detalle verduras
      items.length,                   // J Cant. items
      totalVerduras,                   // K Total verduras ($)
      clean(body.notas),               // L Notas
      extrasTxt,                       // M Extras pedidos
      totalExtras,                     // N Total extras ($)
      clean(body.pago) || 'Transferencia', // O Forma de pago
    ];
    const fmtAO = ['General','@','General','General','General','General','0','0','General','0','0','General','General','0','General'];

    await G.patchRange(token, 'Pedidos', `A${nextRow}:O${nextRow}`, rowAO, fmtAO);

    // Q..S (estado de la app). Se deja P intacta entre medio.
    const rowQS = ['pendiente', G.dateToExcel(fechaISO), id];
    await G.patchRange(token, 'Pedidos', `Q${nextRow}:S${nextRow}`, rowQS, ['General','dd/mm/yyyy','General']);

    return G.json(200, { ok: true, id });
  } catch (err) {
    console.error('[pedido]', err.message);
    return G.json(500, { error: 'Error al guardar el pedido: ' + err.message });
  }
};
