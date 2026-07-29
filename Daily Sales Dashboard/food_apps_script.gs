/**
 * ACE Brands — Daily Food Variance API  (v2 — lock + trigger)
 *
 * SETUP (once):
 *   1. Paste this over the old code, Ctrl+S.
 *   2. Run > setupTrigger  (creates the 10-min cache refresh, removes old triggers)
 *   3. Deploy > Manage deployments > pencil > New version > Deploy
 *
 * Sheet: https://docs.google.com/spreadsheets/d/1rM8b5xiKIQ0y12Z-dOi7yxzeurw_VOuDknt_ARDpK7s
 */

const FOOD_SHEET_ID = '1rM8b5xiKIQ0y12Z-dOi7yxzeurw_VOuDknt_ARDpK7s';
const CACHE_KEY = 'food_v2';
const CACHE_TTL = 21600; // 6h — trigger refreshes every 10 min anyway; long TTL = stale beats nothing

/** One-time: install the refresh trigger (removes any old ones first). */
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshCache').timeBased().everyMinutes(10).create();
  refreshCache();
}

/** Called by the timer. Lock ensures only one rebuild ever runs at a time. */
function refreshCache() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // another rebuild is already running — skip
  try {
    writeCache(buildJson());
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const p = e && e.parameter || {};

  if (p.diag) return diag();

  // 1. Serve from cache — the only path users should ever hit.
  if (!p.bust) {
    const cached = readCache();
    if (cached) {
      return ContentService.createTextOutput(cached)
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 2. Cache empty. Only ONE request may rebuild; everyone else gets
  //    a tiny "warming" response immediately instead of piling up.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) {
    return ContentService.createTextOutput(JSON.stringify({warming: true}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    // Double-check: the build that held the lock may have just filled the cache.
    const cached2 = p.bust ? null : readCache();
    if (cached2) {
      return ContentService.createTextOutput(cached2)
        .setMimeType(ContentService.MimeType.JSON);
    }
    const json = buildJson();
    writeCache(json);
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/** ?diag=1 — timing + dimensions, to see where the seconds go. */
function diag() {
  const t0 = Date.now();
  const ss = SpreadsheetApp.openById(FOOD_SHEET_ID);
  const t1 = Date.now();
  const sh = pickSheet(ss);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  const maxRows = sh.getMaxRows(), maxCols = sh.getMaxColumns();
  const t2 = Date.now();
  const values = sh.getRange(1, 1, lastRow, Math.min(lastCol, 16)).getValues();
  const t3 = Date.now();
  return ContentService.createTextOutput(JSON.stringify({
    sheetName: sh.getName(),
    lastRow, lastCol, maxRows, maxCols,
    openMs: t1 - t0, metaMs: t2 - t1, readMs: t3 - t2,
    cacheChunks: (readCache() || '').length,
  })).setMimeType(ContentService.MimeType.JSON);
}

function pickSheet(ss) {
  let sh = null;
  ss.getSheets().forEach(s => { if (s.getSheetId() === 0) sh = s; });
  if (!sh) {
    ss.getSheets().forEach(s => {
      const n = s.getName().toLowerCase();
      if (!sh && (n.indexOf('inventory') !== -1 || n.indexOf('shift') !== -1 || n.indexOf('count') !== -1)) sh = s;
    });
  }
  return sh || ss.getSheets()[0];
}

function buildJson() {
  const ss = SpreadsheetApp.openById(FOOD_SHEET_ID);
  const sh = pickSheet(ss);

  // Read only the used rows and at most 16 columns — never the whole grid.
  const lastRow = sh.getLastRow();
  const lastCol = Math.min(sh.getLastColumn(), 16);
  if (lastRow < 2) return JSON.stringify({generated: new Date().toISOString(), rows: []});
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();

  // Find header row (contains 'Item')
  let hdrRow = -1;
  for (let i = 0; i < Math.min(values.length, 5); i++) {
    if (values[i].indexOf('Item') !== -1) hdrRow = i;
  }
  if (hdrRow === -1) return JSON.stringify({error: 'Header row not found'});

  const hdrs = values[hdrRow].map(h => String(h).trim().toLowerCase());
  const col = name => hdrs.indexOf(name.toLowerCase());

  const iDt   = col('Date');
  const iSn   = col('Stores');
  const iItem = col('Item');
  const iBeg  = col('Beg Count');
  const iPur  = col('Purchases');
  const iEnd  = col('Ending Count');
  const iAct  = col('Actual Ct');
  const iTgt  = col('Target #');
  const iActD = col('Actual $');
  const iTgtD = col('Target $');

  const num = v => {
    if (v === '' || v === null || v === undefined) return null;
    const n = Number(v);
    return isNaN(n) ? null : Math.round(n * 100) / 100;
  };

  const rows = [];
  for (let i = hdrRow + 1; i < values.length; i++) {
    const r = values[i];
    const dt = normalizeDate(r[iDt]);
    if (!dt || !r[iItem]) continue;
    rows.push({
      dt:   dt,
      sn:   String(r[iSn] || '').replace(/^\d+\s*-\s*/, '').trim(),
      item: String(r[iItem]).trim(),
      beg:  num(r[iBeg]),
      pur:  num(r[iPur]),
      end:  num(r[iEnd]),
      act:  num(r[iAct]),
      tgt:  num(r[iTgt]),
      actd: num(r[iActD]),
      tgtd: num(r[iTgtD]),
    });
  }

  return JSON.stringify({generated: new Date().toISOString(), rows: rows});
}

function readCache() {
  const cache = CacheService.getScriptCache();
  const parts = [];
  for (let i = 0; ; i++) {
    const chunk = cache.get(CACHE_KEY + '_' + i);
    if (chunk === null) break;
    parts.push(chunk);
  }
  return parts.length ? parts.join('') : null;
}

function writeCache(json) {
  // CacheService max 100KB per key — split into chunks
  const chunks = [];
  for (let i = 0; i < json.length; i += 90000) chunks.push(json.substr(i, 90000));
  const toStore = {};
  chunks.forEach((c, i) => toStore[CACHE_KEY + '_' + i] = c);
  try { CacheService.getScriptCache().putAll(toStore, CACHE_TTL); } catch (err) {}
}

function normalizeDate(v) {
  if (!v) return null;
  const num = Number(v);
  if (!isNaN(num) && num > 40000 && num < 60000) {
    const d = new Date(Date.UTC(1899, 11, 30) + num * 86400000);
    return fmtYMD(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
  }
  if (v instanceof Date) {
    const y = v.getFullYear();
    if (y < 2000 || y > 2100) return null;
    return fmtYMD(y, v.getMonth() + 1, v.getDate());
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return fmtYMD(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return fmtYMD(+m[3], +m[1], +m[2]);
  const d = new Date(s);
  if (!isNaN(d) && d.getFullYear() >= 2000 && d.getFullYear() <= 2100) {
    return fmtYMD(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }
  return null;
}

function fmtYMD(y, m, d) {
  return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2);
}
