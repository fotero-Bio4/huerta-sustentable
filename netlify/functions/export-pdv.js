'use strict';

// POST con login: genera el archivo PDV (.xls) para importar al ERP, con los
// pedidos PAGADOS cuyo FechaPedido cae entre {desde, hasta} (inclusive).
// Cada ítem (bolsón / verdura / extra) es una fila.
//
// La fila 1 de la hoja "PDV" es la PLANTILLA con los valores fijos del documento
// (CLIENTE, CONDICIONPAGO, PRODUCTO = "BOLSON AGROECOLOGICO HUERTA", WORKFLOW,
// DIMENSION, etc.). Esos valores se repiten en TODAS las filas sin modificarse.
// Solo se completan por pedido/ítem: NUMERO, FECHA, DESCRIPCION (cliente + datos),
// DESCRIPCIONITEM (cada ítem), CANTIDAD y PRECIO. La hoja "Mapeo PDV" aporta, como
// respaldo, valores fijos que no estén en la plantilla y el NUMERO_INICIAL.

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
      tel:      String(r[COL.tel] ?? '').trim(),
      email:    String(r[COL.email] ?? '').trim(),
      barrio:   String(r[COL.barrio] ?? '').trim(),
      dir:      String(r[COL.dir] ?? '').trim(),
      entrega:  String(r[COL.entrega] ?? '').trim(),
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

// Valores fijos de respaldo desde "Mapeo PDV" (sección clave/valor).
function parseMapeoFixed(rows) {
  const fixed = {};
  let mode = 'fixed';
  for (const r of (rows || [])) {
    const a = String(r[0] ?? '').trim();
    const al = a.toLowerCase();
    if (!a) continue;
    if (al.startsWith('nombre app') || al === 'productos') { mode = 'prod'; continue; }
    if (al.startsWith('campo') || al === 'cabecera' || al.startsWith('valores fijos')) { mode = 'fixed'; continue; }
    if (al.startsWith('🗺') || al.startsWith('mapeo')) continue;
    if (mode === 'fixed') fixed[a.toUpperCase()] = r[1] ?? '';
  }
  return fixed;
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

    // Encabezado y plantilla de valores fijos (fila 1 de PDV).
    const header = (pdv.values && pdv.values[0]) ? pdv.values[0].map(h => String(h ?? '').trim()) : [];
    if (!header.length) return G.json(500, { error: 'La hoja PDV no tiene encabezados.' });
    const tmpl = (pdv.values && pdv.values[1]) ? pdv.values[1] : [];
    const fixed = { ...parseMapeoFixed(map.values || []) };   // respaldo
    header.forEach((h, i) => {                                // la plantilla PDV manda
      const v = tmpl[i];
      if (v !== '' && v !== null && v !== undefined) fixed[h.toUpperCase()] = v;
    });

    // Precios de extras desde Config (misma fuente que el formulario).
    const cfg = {};
    for (const r of (conf.values || [])) { const k = String(r[0] ?? '').trim(); if (k && !/clave/i.test(k)) cfg[k] = num(r[1]); }
    const EXTRA_KEY = { 'bandeja sopera':'PRECIO_SOPERA', 'bandeja de ensalada':'PRECIO_ENSALADA', 'escabeche de berenjenas':'PRECIO_ESCABECHE' };
    const precios = priceIndex(det.values || []);
    const precioExtra = (nombre) => cfg[EXTRA_KEY[nombre.toLowerCase()]] || precios[nombre.toLowerCase()] || 0;

    const pedidos = buildPedidos(ped.values || [], ped.firstRow)
      .filter(p => p.estado === 'pagado' && p.fechaISO && p.fechaISO >= desde && p.fechaISO <= hasta);
    if (!pedidos.length) return G.json(404, { error: 'No hay pedidos pagados en ese rango de fechas.' });

    // Descripción de cabecera: nombre del cliente + sus datos.
    const descripcionCliente = (p) => {
      const entrega = p.entrega === 'Retiro'
        ? 'Retiro en huerta'
        : [p.dir, p.barrio].filter(Boolean).join(', ');
      return [p.nombre, p.tel ? 'Tel: ' + p.tel : '', p.email, entrega].filter(Boolean).join(' | ');
    };

    const numInicial = num(fixed.NUMERO_INICIAL) || num(tmpl[header.findIndex(h => h.toUpperCase() === 'NUMERO')]) || 1;
    const dataRows = [];
    let numero = numInicial;

    for (const p of pedidos) {
      const items = [];
      if (p.bolsones > 0) items.push({ nombre: 'Bolsón semanal', cantidad: p.bolsones, precio: p.precioBolson });
      for (const it of G.parseDetalle(p.detalle)) items.push({ nombre: it.nombre, cantidad: it.cantidad, precio: precios[it.nombre.toLowerCase()] || 0 });
      for (const e of G.parseDetalle(p.extras))  items.push({ nombre: e.nombre, cantidad: e.cantidad, precio: precioExtra(e.nombre) });
      if (!items.length) continue;

      const descCliente = descripcionCliente(p);
      for (const it of items) {
        const computed = {
          NUMERO:          numero,
          FECHA:           fechaLegible(p.fechaISO),
          DESCRIPCION:     descCliente,
          DESCRIPCIONITEM: it.nombre,
          CANTIDAD:        it.cantidad,
          PRECIO:          it.precio,
        };
        const row = header.map(h => {
          const key = h.toUpperCase();
          if (key in computed && computed[key] !== '' && computed[key] !== undefined && computed[key] !== null) return computed[key];
          if (key in fixed) return fixed[key];
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
    return G.json(200, { ok: true, filename, base64, pedidos: pedidos.length, filas: dataRows.length, warnings: 0 });
  } catch (err) {
    console.error('[export-pdv]', err.message);
    return G.json(500, { error: 'Error al generar el PDV: ' + err.message });
  }
};
