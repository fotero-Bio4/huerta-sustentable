'use strict';

// POST con login: guarda configuración general (hoja "Config", clave/valor) y
// precios/stock de cada verdura (hoja "Detalle por Verdura").

const G = require('./_graph');

function findCol(header, ...names) {
  const norm = s => String(s ?? '').trim().toLowerCase();
  const wanted = names.map(norm);
  for (let i = 0; i < header.length; i++) {
    const h = norm(header[i]);
    if (wanted.some(w => h === w || h.startsWith(w))) return i;
  }
  return -1;
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

    // ── Config general (clave/valor) ──────────────────────────────────────────
    if (body.config && typeof body.config === 'object') {
      const { values: rows, firstRow } = await G.readSheet(token, 'Config');
      const keyRow = {};
      let lastRow = firstRow;
      for (let i = 0; i < rows.length; i++) {
        const k = String(rows[i][0] ?? '').trim();
        if (k) { keyRow[k] = firstRow + i; lastRow = firstRow + i; }
      }
      for (const [k, v] of Object.entries(body.config)) {
        const val = typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : v;
        if (keyRow[k]) {
          await G.patchCell(token, 'Config', `B${keyRow[k]}`, val);
        } else {
          lastRow += 1;
          await G.patchRange(token, 'Config', `A${lastRow}:B${lastRow}`, [k, val]);
          keyRow[k] = lastRow;
        }
      }
    }

    // ── Precios / stock por verdura ───────────────────────────────────────────
    const wantsVerduras = Array.isArray(body.precios) || Array.isArray(body.stock);
    if (wantsVerduras) {
      const { values: rows, firstRow } = await G.readSheet(token, 'Detalle por Verdura');
      let hIdx = rows.findIndex(r => r.some(c => String(c ?? '').trim().toLowerCase() === 'grupo'));
      if (hIdx === -1) hIdx = 1;
      const header = rows[hIdx] || [];
      const cN = findCol(header, 'verdura / producto', 'verdura', 'producto');
      const cP = findCol(header, 'precio unit. ($)', 'precio unit', 'precio');
      const cS = findCol(header, 'stock');

      const rowByName = {};
      for (let i = hIdx + 1; i < rows.length; i++) {
        const nombre = String(rows[i][cN] ?? '').trim().toLowerCase();
        if (nombre) rowByName[nombre] = firstRow + i;
      }

      for (const p of (body.precios || [])) {
        const er = rowByName[String(p.nombre || '').trim().toLowerCase()];
        if (er && cP !== -1 && Number.isFinite(Number(p.precio)))
          await G.patchCell(token, 'Detalle por Verdura', `${G.colLetter(cP + 1)}${er}`, Number(p.precio));
      }
      if (cS !== -1) {
        for (const s of (body.stock || [])) {
          const er = rowByName[String(s.nombre || '').trim().toLowerCase()];
          if (er) await G.patchCell(token, 'Detalle por Verdura', `${G.colLetter(cS + 1)}${er}`, s.enStock ? 'TRUE' : 'FALSE');
        }
      }
    }

    return G.json(200, { ok: true });
  } catch (err) {
    console.error('[config]', err.message);
    return G.json(500, { error: 'Error del servidor: ' + err.message });
  }
};
