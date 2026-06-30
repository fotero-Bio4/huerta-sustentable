'use strict';

// POST público (rate-limited): agrega un pedido nuevo a la hoja "Pedidos".
// Escribe columnas A–O (la P = Total tiene fórmula y se deja intacta) y las
// columnas de estado de la app Q=Estado, R=FechaPedido, S=IDPedido.

const G = require('./_graph');

function num(v) { const n = Number(v); return Number.isFinite(n) ? n : 0; }
function clean(v) { return String(v ?? '').trim().slice(0, 500); }

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

    // Próxima fila libre (encabezados en fila 4, datos desde fila 5; col A = clave).
    const { values: rows, firstRow } = await G.readSheet(token, 'Pedidos');
    const nextRow = G.nextRowFromRows(rows, firstRow, 0);

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
