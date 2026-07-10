'use strict';

// POST con login: genera el archivo PDV (.xls) para importar al ERP, con los
// pedidos cuyo FechaPedido cae entre {desde, hasta} (inclusive) y estado = pagado.
// Cada ítem (bolsón / verdura / extra) es una fila. El mapeo de códigos y los
// valores fijos de cabecera se leen de la hoja "Mapeo PDV".

const G = require('./_graph');
const XLSX = require('xlsx');

const COL = {
  nombre: 0, tel: 1, email: 2, barrio: 3, dir: 4, entrega: 5,
  bolsones: 6, precioBolson: 7, detalle: 8, totalVerd: 10,
  extras: 12, pago: 14, total: 15, estado: 16, fecha: 17, id: 18,
};
function num(v) { const n = Number(v); return Number.isFinite(n) ? n : 0; }

function buildPedidos(rows, firstRow) {
  let hIdx = rows.findIndex(r => String(r[0] ?? '').trim().toLowerCase() === 'nombre');
  if (hIdx === -1) hIdx = 3;
  const out = [];
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const nombre = String(r[COL.nombre] ?? '').trim();
    if (!nombre) continue;
    out.push({
      nombre,
      barrio:   String(r[COL.barrio] ?? '').trim(),
      bolsones: num(r[COL.bolsones]),
      precioBolson: num(r[COL.precioBolson]),
      detalle:  String(r[COL.detalle] ?? '').trim(),
      extras:   String(r[COL.extras] ?? '').trim(),
      pago:     String(r[COL.pago] ?? '').trim(),
      estado:   G.normEstado(r[COL.estado]),
      fechaISO: G.anyToISO(r[COL.fecha]),
    });
  }
  return out;
}

// Precio por verdura desde "Detalle por Verdura".
function priceIndex(rows) {
  let hIdx = rows.findIndex(r => r.some(c => String(c ?? '').trim().toLowerCase() === 'grupo'));
  if (hIdx === -1) hIdx = 1;
  const map = {};
  for (let i = hIdx + 1; i < rows.length; i++) {
    const nombre = String(rows[i][0] ?? '').trim();
    if (nombre && !/total/i.test(nombre)) map[nombre.toLowerCase()] = num(rows[i][3]);
  }
  return map;
}

// Parsea la hoja "Mapeo PDV": tabla de productos + valores fijos de cabecera.
function parseMapeo(rows) {
  const productMap = {}; // nombre.toLowerCase() → { producto, descitem, tipoitem }
  const fixed = {};      // CAMPO (mayúsc) → valor
  let mode = 'fixed';
  for (const r of rows) {
    const a = String(r[0] ?? '').trim();
    const al = a.toLowerCase();
    if (!a) continue;
    if (al.startsWith('nombre app') || al === 'productos') { mode = 'prod'; continue; }
    if (al.startsWith('campo') || al === 'cabecera' || al.startsWith('valores fijos')) { mode = 'fixed'; continue; }
    if (al.startsWith('🗺') || al.startsWith('mapeo')) continue; // título
    if (mode === 'prod') {
      productMap[al] = {
        producto: String(r[1] ?? '').trim(),
        descitem: String(r[2] ?? '').trim() || a,
        tipoitem: String(r[3] ?? '').trim(),
      };
    } else {
      fixed[a.toUpperCase()] = r[1] ?? '';
    }
  }
  return { productMap, fixed };
}

function fechaLegible(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

exports.handler = async (event) => {
  const pf = G.preflight(event); if (pf) return pf;
  if (event.httpMethod !== 'POST') return G.json(405, { error: 'Método no permitido' });
  if (!G.envOk())                  return G.json(500, { error: 'Servidor no configurado.' });

  let body;
  try { body = JSON.parse(event.body || '{}'); }
  catch { return G.json(400, { error: 'Cuerpo inválido.' }); }

  const desde = G.anyToISO(body.desde);
  const hasta = G.anyToISO(body.hasta);
  if (!desde || !hasta) return G.json(400, { error: 'Indicá fecha desde y hasta.' });

  try {
    const token = await G.getToken();
    const user = await G.validateUser(token, body.mail, body.pss);
    if (!user) return G.json(401, { error: 'Mail o contraseña incorrectos.' });

    const [ped, det, pdv, map, conf] = await Promise.all([
      G.readSheet(token, 'Pedidos'),
      G.readSheet(token, 'Detalle por Verdura'),
      G.readSheet(token, 'PDV'),
      G.readSheet(token, 'Mapeo PDV'),
      G.readSheet(token, 'Config'),
    ]);

    // Precios de extras desde Config (misma fuente que el formulario).
    const cfg = {};
    for (const r of (conf.values || [])) { const k = String(r[0] ?? '').trim(); if (k && !/clave/i.test(k)) cfg[k] = num(r[1]); }
    const EXTRA_KEY = { 'bandeja sopera':'PRECIO_SOPERA', 'bandeja de ensalada':'PRECIO_ENSALADA', 'escabeche de berenjenas':'PRECIO_ESCABECHE' };
    const precioExtra = (nombre) => cfg[EXTRA_KEY[nombre.toLowerCase()]] || precios[nombre.toLowerCase()] || num(fixed['PRECIO_' + nombre.toUpperCase()]) || 0;

    const header = (pdv.values && pdv.values[0]) ? pdv.values[0].map(h => String(h ?? '').trim()) : [];
    if (!header.length) return G.json(500, { error: 'La hoja PDV no tiene encabezados.' });

    const pedidos = buildPedidos(ped.values || [], ped.firstRow)
      .filter(p => p.estado === 'pagado' && p.fechaISO && p.fechaISO >= desde && p.fechaISO <= hasta);

    if (!pedidos.length) return G.json(404, { error: 'No hay pedidos pagados en ese rango de fechas.' });

    const precios = priceIndex(det.values || []);
    const { productMap, fixed } = parseMapeo(map.values || []);

    const prod = (nombre) => productMap[String(nombre).toLowerCase()] || null;
    const codigo = (nombre) => { const p = prod(nombre); return p && p.producto ? p.producto : `[FALTA CODIGO: ${nombre}]`; };
    const descItem = (nombre) => { const p = prod(nombre); return p && p.descitem ? p.descitem : nombre; };
    const tipoItem = (nombre) => { const p = prod(nombre); return (p && p.tipoitem) || fixed.TIPOITEM || ''; };

    // Expandir pedidos → ítems
    const numInicial = num(fixed.NUMERO_INICIAL) || 1;
    const dataRows = [];
    let numero = numInicial;
    let warnings = 0;

    for (const p of pedidos) {
      const items = [];
      if (p.bolsones > 0) items.push({ nombre: 'Bolsón semanal', cantidad: p.bolsones, precio: p.precioBolson });
      for (const it of G.parseDetalle(p.detalle)) items.push({ nombre: it.nombre, cantidad: it.cantidad, precio: precios[it.nombre.toLowerCase()] || 0 });
      for (const e of G.parseDetalle(p.extras))  items.push({ nombre: e.nombre, cantidad: e.cantidad, precio: precioExtra(e.nombre) });
      if (!items.length) continue;

      const condPago = fixed['CONDICIONPAGO_' + String(p.pago).toUpperCase()] || fixed.CONDICIONPAGO || p.pago;

      for (const it of items) {
        if (codigo(it.nombre).startsWith('[FALTA')) warnings++;
        const computed = {
          NUMERO:         numero,
          FECHA:          fechaLegible(p.fechaISO),
          CLIENTE:        fixed.CLIENTE || p.nombre,
          CONDICIONPAGO:  condPago,
          PRODUCTO:       codigo(it.nombre),
          DESCRIPCIONITEM:descItem(it.nombre),
          CANTIDAD:       it.cantidad,
          PRECIO:         it.precio,
          TIPOITEM:       tipoItem(it.nombre),
          FECHAENTREGA:   fixed.FECHAENTREGA || fechaLegible(p.fechaISO),
        };
        const row = header.map(h => {
          const key = h.toUpperCase();
          if (key in computed && computed[key] !== '' && computed[key] !== undefined) return computed[key];
          if (key in fixed) return fixed[key];
          if (key in computed) return computed[key];
          return '';
        });
        dataRows.push(row);
      }
      numero++;
    }

    // Generar .xls (BIFF8) en base64
    const aoa = [header, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet0');
    const base64 = XLSX.write(wb, { type: 'base64', bookType: 'biff8' });

    const filename = `PDV_huerta_${desde}_a_${hasta}.xls`;
    return G.json(200, {
      ok: true,
      filename,
      base64,
      pedidos: pedidos.length,
      filas: dataRows.length,
      warnings,
    });
  } catch (err) {
    console.error('[export-pdv]', err.message);
    return G.json(500, { error: 'Error al generar el PDV: ' + err.message });
  }
};
