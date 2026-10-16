/**
 * LIBRO DE ASISTENCIA — Servidor de sincronización en Google Drive
 * ----------------------------------------------------------------
 * Guarda la información de todos los equipos en esta hoja de cálculo
 * (pestaña "datos") y la entrega a cada equipo conectado.
 *
 * 1. Cambie la CLAVE de abajo por una frase secreta propia.
 * 2. Implementar → Nueva implementación → Aplicación web
 *    Ejecutar como: Yo   ·   Quién tiene acceso: Cualquier usuario
 * 3. Copie la URL que termina en /exec y péguela en el sistema
 *    (Panel de administrador → Sincronización con Google Drive).
 *
 * No edite la pestaña "datos" a mano.
 */
const CLAVE = 'CAMBIE-ESTA-CLAVE';

const HOJA = 'datos';
const TROZO = 45000;           // una celda de Sheets admite hasta 50 000 caracteres
const FUSIONAR = ['asis|', 'lista|'];

function doGet(e) { return responder((e && e.parameter) || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) {}
  return responder(p);
}

function responder(p) {
  let r;
  try {
    if (CLAVE === 'CAMBIE-ESTA-CLAVE') r = { ok: false, error: 'Cambie la CLAVE en Codigo.gs y vuelva a implementar.' };
    else if (!p.clave || p.clave !== CLAVE) r = { ok: false, error: 'clave' };
    else if (p.op === 'ping') r = { ok: true, now: ultimo(), filas: Math.max(0, hoja().getLastRow() - 1) };
    else if (p.op === 'pull') r = pull(Number(p.since) || 0);
    else if (p.op === 'push') r = push(p.items || []);
    else r = { ok: false, error: 'Operación desconocida' };
  } catch (err) {
    r = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

function hoja() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJA);
  if (!h) {
    h = ss.insertSheet(HOJA);
    h.getRange(1, 1, 1, 3).setValues([['clave', 'actualizado', 'valor']]);
    h.setFrozenRows(1);
    h.getRange('A:Z').setNumberFormat('@');
  }
  return h;
}
function ultimo() { return Number(PropertiesService.getScriptProperties().getProperty('ULT') || 0); }
function ancho(h) { return Math.max(1, h.getLastColumn() - 2); }

/* Entrega lo que cambió después de "since" */
function pull(since) {
  const ult = ultimo();
  if (since && since >= ult) return { ok: true, now: ult, rows: [] };
  const h = hoja(), n = h.getLastRow() - 1;
  if (n < 1) return { ok: true, now: ult, rows: [] };
  const ab = h.getRange(2, 1, n, 2).getValues();
  const idx = [];
  ab.forEach((r, i) => { if (Number(r[1]) > since) idx.push(i); });
  let rows = [];
  if (idx.length) {
    const w = ancho(h);
    if (idx.length > 120) {
      const todo = h.getRange(2, 3, n, w).getValues();
      rows = idx.map(i => ({ k: String(ab[i][0]), t: Number(ab[i][1]), v: todo[i].join('') }));
    } else {
      rows = idx.map(i => ({ k: String(ab[i][0]), t: Number(ab[i][1]), v: h.getRange(i + 2, 3, 1, w).getValues()[0].join('') }));
    }
  }
  return { ok: true, now: ult, rows: rows };
}

/* Recibe cambios de un equipo */
function push(items) {
  const lock = LockService.getScriptLock();
  lock.waitLock(28000);
  try {
    const h = hoja();
    let n = h.getLastRow() - 1, w = ancho(h);
    const fila = {};
    if (n > 0) h.getRange(2, 1, n, 1).getValues().forEach((r, i) => { fila[String(r[0])] = i + 2; });
    let t = Math.max(Date.now(), ultimo() + 1);
    const nuevos = [], cambios = [];

    items.forEach(it => {
      if (!it || !it.k) return;
      let v = it.v == null ? '' : (typeof it.v === 'string' ? it.v : JSON.stringify(it.v));
      const f = fila[it.k];
      if (f && v && FUSIONAR.some(p => it.k.indexOf(p) === 0)) {
        const prev = h.getRange(f, 3, 1, w).getValues()[0].join('');
        if (prev) { try { v = JSON.stringify(fusionar(it.k, JSON.parse(prev), JSON.parse(v))); } catch (x) {} }
      }
      const trozos = [];
      for (let i = 0; i < v.length; i += TROZO) trozos.push(v.slice(i, i + TROZO));
      if (!trozos.length) trozos.push('');
      if (trozos.length > w) w = trozos.length;
      t++;
      if (f) cambios.push([f, it.k, t, trozos]);
      else if (fila[it.k] !== -1) { nuevos.push([it.k, t, trozos]); fila[it.k] = -1; }
    });

    if (h.getMaxColumns() < 2 + w) h.insertColumnsAfter(h.getMaxColumns(), 2 + w - h.getMaxColumns());
    const fil = (k, tt, tr) => { const r = [k, tt].concat(tr); while (r.length < 2 + w) r.push(''); return r; };

    if (cambios.length > 25 && n > 0) {
      const todo = h.getRange(2, 1, n, 2 + w).getValues();
      cambios.forEach(c => { todo[c[0] - 2] = fil(c[1], c[2], c[3]); });
      h.getRange(2, 1, n, 2 + w).setValues(todo);
    } else {
      cambios.forEach(c => h.getRange(c[0], 1, 1, 2 + w).setValues([fil(c[1], c[2], c[3])]));
    }
    if (nuevos.length) {
      const r0 = h.getLastRow() + 1;
      const rg = h.getRange(r0, 1, nuevos.length, 2 + w);
      rg.setNumberFormat('@');
      rg.setValues(nuevos.map(x => fil(x[0], x[1], x[2])));
    }
    PropertiesService.getScriptProperties().setProperty('ULT', String(t));
    return { ok: true, now: t, guardados: cambios.length + nuevos.length };
  } finally {
    lock.releaseLock();
  }
}

/* La asistencia de un día se combina estudiante por estudiante: gana el registro más reciente */
function fusionar(k, a, b) {
  const out = {};
  Object.keys(a || {}).forEach(id => { out[id] = a[id]; });
  Object.keys(b || {}).forEach(id => {
    const x = out[id], y = b[id];
    if (!x) { out[id] = y; return; }
    if (k.indexOf('asis|') === 0) { if ((Number(y.u) || 0) >= (Number(x.u) || 0)) out[id] = y; }
    else { if (String(y.at || '') >= String(x.at || '')) out[id] = y; }
  });
  return out;
}
