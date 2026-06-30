'use strict';

// ════════════════════════════════════════════════════════════════════════════
//  Helpers compartidos para Microsoft Graph (Excel SharePoint como base de datos)
//  Reusa el patrón de RPA0045_VariablesDestileria (client_credentials, app-only).
// ════════════════════════════════════════════════════════════════════════════

const GRAPH = 'https://graph.microsoft.com/v1.0';

const CORS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Content-Type':                 'application/json',
};

// ── Respuestas JSON estándar ───────────────────────────────────────────────
function json(statusCode, obj) {
  return { statusCode, headers: CORS, body: JSON.stringify(obj) };
}
function preflight(event) {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: CORS, body: '' };
  return null;
}

// ── Rate limiting en memoria (por IP) ───────────────────────────────────────
const rateMap = new Map();
function checkRate(ip, max = 15, windowMin = 15) {
  const now = Date.now();
  const e = rateMap.get(ip);
  if (!e || now - e.t > windowMin * 60 * 1000) { rateMap.set(ip, { t: now, n: 1 }); return false; }
  e.n++;
  return e.n > max;
}
function clientIp(event) {
  return ((event.headers['x-forwarded-for'] || '').split(',')[0].trim()) || 'unknown';
}

// ── Auth Graph ──────────────────────────────────────────────────────────────
async function getToken() {
  const { GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET } = process.env;
  const r = await fetch(
    `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/oauth2/v2.0/token`,
    {
      method:  'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body:    new URLSearchParams({
        grant_type:    'client_credentials',
        client_id:     GRAPH_CLIENT_ID,
        client_secret: GRAPH_CLIENT_SECRET,
        scope:         'https://graph.microsoft.com/.default',
      }),
    }
  );
  const d = await r.json();
  if (!d.access_token) throw new Error('Auth Graph: ' + (d.error_description || JSON.stringify(d)));
  return d.access_token;
}

// ── Config del archivo (drive + item) ────────────────────────────────────────
function fileBase() {
  const { HUERTA_DRIVE_ID, HUERTA_ITEM_ID } = process.env;
  return `${GRAPH}/drives/${HUERTA_DRIVE_ID}/items/${HUERTA_ITEM_ID}/workbook`;
}
function envOk() {
  const { GRAPH_CLIENT_SECRET, HUERTA_DRIVE_ID, HUERTA_ITEM_ID } = process.env;
  return Boolean(GRAPH_CLIENT_SECRET && HUERTA_DRIVE_ID && HUERTA_ITEM_ID);
}

// ── Lectura / escritura de rangos ─────────────────────────────────────────────
function parseFirstRow(address) {
  // "Pedidos!A4:P104" → 4
  const rangePart = (address || '').split('!').pop();
  const m = rangePart.match(/\$?[A-Z]+\$?(\d+)/);
  return m ? parseInt(m[1], 10) : 1;
}

// Lee el usedRange de una hoja. Devuelve { values, firstRow }. Si la hoja no existe → vacío.
async function readSheet(token, sheet) {
  const enc = encodeURIComponent(sheet);
  const r = await fetch(`${fileBase()}/worksheets/${enc}/usedRange(valuesOnly=true)`,
    { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) return { values: [], firstRow: 1, ok: false };
  const d = await r.json();
  return { values: d.values || [], firstRow: parseFirstRow(d.address), ok: true };
}

// PATCH de un rango con valores (y opcionalmente formatos de número).
async function patchRange(token, sheet, addr, rowValues, rowFormats) {
  const enc = encodeURIComponent(sheet);
  const body = { values: [rowValues] };
  if (rowFormats) body.numberFormat = [rowFormats];
  const r = await fetch(`${fileBase()}/worksheets/${enc}/range(address='${addr}')`, {
    method:  'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body:    JSON.stringify(body),
  });
  if (!r.ok) throw new Error(`PATCH ${sheet}!${addr}: ${r.status} ${await r.text()}`);
}

// PATCH de una sola celda.
async function patchCell(token, sheet, cellAddr, value) {
  const enc = encodeURIComponent(sheet);
  const r = await fetch(`${fileBase()}/worksheets/${enc}/range(address='${cellAddr}')`, {
    method:  'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body:    JSON.stringify({ values: [[value]] }),
  });
  if (!r.ok) throw new Error(`PATCH ${sheet}!${cellAddr}: ${r.status} ${await r.text()}`);
}

// ── Utilidades de columnas / fechas ───────────────────────────────────────────
function colLetter(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
// YYYY-MM-DD → serial Excel (25569 = días entre 1900-01-01 y 1970-01-01).
function dateToExcel(dateStr) {
  return Math.floor(Date.parse(dateStr + 'T00:00:00Z') / 86400000) + 25569;
}
// serial Excel → YYYY-MM-DD
function excelToDate(serial) {
  const ms = (Number(serial) - 25569) * 86400000;
  return new Date(ms).toISOString().slice(0, 10);
}
// Acepta serial Excel o string ISO/legible y devuelve YYYY-MM-DD (o '' si no se puede).
function anyToISO(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return excelToDate(v);
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);          // ISO
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);        // dd/mm/yyyy
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  return '';
}

// Próxima fila Excel libre mirando una columna clave (1-indexed).
function nextRowFromRows(rows, firstRow, keyColIdx = 0) {
  let lastDataRow = firstRow;
  for (let i = 1; i < rows.length; i++) {
    const v = rows[i][keyColIdx];
    if (v !== null && v !== undefined && v !== '' && v !== 0) lastDataRow = firstRow + i;
  }
  return lastDataRow + 1;
}

// ── Detalle de verduras: formato canónico ──────────────────────────────────────
// items: [{ nombre, cantidad, unidad }]  →  "Lechuga crespa x2 planta | Mizuna x1 atado"
function formatDetalle(items) {
  return items
    .filter(it => it && it.nombre && Number(it.cantidad) > 0)
    .map(it => `${it.nombre} x${Number(it.cantidad)}${it.unidad ? ' ' + it.unidad : ''}`)
    .join(' | ');
}
// Parsea el texto canónico de vuelta a items. Tolera "x", "×" y separador "|" o "/".
function parseDetalle(text) {
  if (!text) return [];
  return String(text)
    .split(/\s*[|/]\s*/)
    .map(s => s.trim())
    .filter(Boolean)
    .map(seg => {
      const m = seg.match(/^(.*?)\s*[x×]\s*(\d+(?:[.,]\d+)?)\s*(.*)$/i);
      if (!m) return { nombre: seg, cantidad: 1, unidad: '' };
      return {
        nombre:   m[1].trim(),
        cantidad: Number(String(m[2]).replace(',', '.')),
        unidad:   (m[3] || '').trim(),
      };
    });
}

// ── Login (hoja Usuarios: Mail | Nombre | Pss) ──────────────────────────────────
async function validateUser(token, mail, pss) {
  if (!mail || !pss) return null;
  const { values: rows } = await readSheet(token, 'Usuarios');
  // Encabezado en la primera fila; datos debajo.
  const row = rows.slice(1).find(r =>
    String(r[0] ?? '').trim().toLowerCase() === String(mail).trim().toLowerCase() &&
    String(r[2] ?? '').trim() === String(pss).trim()
  );
  if (!row) return null;
  return { nombre: String(row[1] ?? '').trim() || String(row[0]).trim(), mail: String(row[0]).trim() };
}

module.exports = {
  GRAPH, CORS, json, preflight,
  checkRate, clientIp,
  getToken, fileBase, envOk,
  readSheet, patchRange, patchCell,
  colLetter, dateToExcel, excelToDate, anyToISO, nextRowFromRows,
  formatDetalle, parseDetalle,
  validateUser,
};
