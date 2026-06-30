'use strict';

// POST con login: actualiza el estado (col Q) y/o el precio final (col P) de un
// pedido identificado por su fila de Excel.

const G = require('./_graph');

const ESTADOS = ['pendiente', 'preparando', 'entregado', 'cancelado'];

exports.handler = async (event) => {
  const pf = G.preflight(event); if (pf) return pf;
  if (event.httpMethod !== 'POST') return G.json(405, { error: 'Método no permitido' });
  if (!G.envOk())                  return G.json(500, { error: 'Servidor no configurado.' });

  let body;
  try { body = JSON.parse(event.body || '{}'); }
  catch { return G.json(400, { error: 'Cuerpo inválido.' }); }

  const row = parseInt(body.excelRow, 10);
  if (!row || row < 5) return G.json(400, { error: 'Fila inválida.' });

  try {
    const token = await G.getToken();
    const user = await G.validateUser(token, body.mail, body.pss);
    if (!user) return G.json(401, { error: 'Mail o contraseña incorrectos.' });

    if (body.estado) {
      const est = String(body.estado).trim().toLowerCase();
      if (!ESTADOS.includes(est)) return G.json(400, { error: 'Estado inválido.' });
      await G.patchCell(token, 'Pedidos', `Q${row}`, est);
    }

    if (body.total !== undefined && body.total !== null && body.total !== '') {
      const total = Number(body.total);
      if (!Number.isFinite(total)) return G.json(400, { error: 'Total inválido.' });
      // Sobrescribe la fórmula de la fila con el precio final confirmado.
      await G.patchCell(token, 'Pedidos', `P${row}`, total);
    }

    return G.json(200, { ok: true });
  } catch (err) {
    console.error('[estado]', err.message);
    return G.json(500, { error: 'Error del servidor: ' + err.message });
  }
};
