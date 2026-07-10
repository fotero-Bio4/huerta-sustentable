'use strict';

// POST con login: actualiza el estado (col Q) y/o el precio final (col P) de un
// pedido, o cambia el estado de varios pedidos a la vez (body.excelRows[]).

const G = require('./_graph');

const ESTADOS = ['pendiente', 'confirmado', 'pagado', 'cancelado'];

exports.handler = async (event) => {
  const pf = G.preflight(event); if (pf) return pf;
  if (event.httpMethod !== 'POST') return G.json(405, { error: 'Método no permitido' });
  if (!G.envOk())                  return G.json(500, { error: 'Servidor no configurado.' });

  let body;
  try { body = JSON.parse(event.body || '{}'); }
  catch { return G.json(400, { error: 'Cuerpo inválido.' }); }

  // Filas a actualizar: una sola (excelRow) o varias (excelRows[]).
  const rowsIn = Array.isArray(body.excelRows) ? body.excelRows : [body.excelRow];
  const rows = rowsIn.map(r => parseInt(r, 10)).filter(r => r && r >= 5);
  if (!rows.length) return G.json(400, { error: 'Fila inválida.' });

  try {
    const token = await G.getToken();
    const user = await G.validateUser(token, body.mail, body.pss);
    if (!user) return G.json(401, { error: 'Mail o contraseña incorrectos.' });

    if (body.estado) {
      const raw = String(body.estado).trim().toLowerCase();
      const permitidos = [...ESTADOS, 'preparando', 'entregado']; // acepta legacy
      if (!permitidos.includes(raw)) return G.json(400, { error: 'Estado inválido.' });
      const est = G.normEstado(raw);
      for (const row of rows) await G.patchCell(token, 'Pedidos', `Q${row}`, est);
    }

    if (body.total !== undefined && body.total !== null && body.total !== '') {
      const total = Number(body.total);
      if (!Number.isFinite(total)) return G.json(400, { error: 'Total inválido.' });
      // Sobrescribe la fórmula de la fila con el precio final confirmado (solo single).
      await G.patchCell(token, 'Pedidos', `P${rows[0]}`, total);
    }

    return G.json(200, { ok: true });
  } catch (err) {
    console.error('[estado]', err.message);
    return G.json(500, { error: 'Error del servidor: ' + err.message });
  }
};
