/**
 * ACE Brands — Daily Food Variance API
 * Deploy: Extensions > Apps Script > paste this > Deploy > New deployment
 *         Type: Web app · Execute as: Me · Who has access: Anyone
 * Sheet:  https://docs.google.com/spreadsheets/d/1rM8b5xiKIQ0y12Z-dOi7yxzeurw_VOuDknt_ARDpK7s
 */

const FOOD_SHEET_ID = '1rM8b5xiKIQ0y12Z-dOi7yxzeurw_VOuDknt_ARDpK7s';
const CACHE_KEY = 'food_v1';
const CACHE_TTL = 1500; // 25 min (refreshed every 10 by trigger)

/**
 * ONE-TIME SETUP: run this function once from the editor (Run > setupTrigger).
 * It creates a timer that refreshes the cache every 10 minutes,
 * so dashboard users never hit the slow cold build.
 */
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshCache').timeBased().everyMinutes(10).create();
  refreshCache(); // warm it right now too
}

function refreshCache() {
  writeCache(buildJson());
}

function doGet(e) {
  const bust = e && e.parameter && e.parameter.bust;

  if (!bust) {
    const cached = readCache();
    if (cached) {
      return ContentService.createTextOutput(cached)
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  const json = buildJson();
  writeCache(json);

  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
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

function buildJson() {
  const ss = SpreadsheetApp.openById(FOOD_SHEET_ID);
  // Target the tab: gid=0, or a tab whose name mentions inventory/shift/count
  let sh = null;
  ss.getSheets().forEach(s => {
    if (s.getSheetId() === 0) sh = s;
  });
  if (!sh) {
    ss.getSheets().forEach(s => {
      const n = s.getName().toLowerCase();
      if (!sh && (n.indexOf('inventory') !== -1 || n.indexOf('shift') !== -1 || n.indexOf('count') !== -1)) sh = s;
    });
  }
  if (!sh) sh = ss.getSheets()[0];
  const values = sh.getDataRange().getValues();

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
