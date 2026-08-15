/**
 * ACE Brands — Food + Labor API  (v3)
 *
 * Endpoints (same URL):
 *   ?              -> food   (Inventory Shift Count)  — unchanged
 *   ?data=labor    -> labor  (PAR + Actual Timecard + Employee Labor Analysis)
 *   ?diag=1        -> timings
 *   &bust=1        -> bypass cache
 *
 * SETUP (once):
 *   1. Paste this over the old code, Ctrl+S.
 *   2. Run > setupTrigger  (installs the 10-min refresh for BOTH datasets)
 *   3. Deploy > Manage deployments > pencil > New version > Deploy
 *
 * Sheet: https://docs.google.com/spreadsheets/d/1rM8b5xiKIQ0y12Z-dOi7yxzeurw_VOuDknt_ARDpK7s
 */

const FOOD_SHEET_ID  = '1rM8b5xiKIQ0y12Z-dOi7yxzeurw_VOuDknt_ARDpK7s';
const LABOR_SHEET_ID = '1m3aYo6J8wYDKY_1S-TIVpU-JYoZ8VzcfV7qvT6ZnFlY';
const CACHE_KEY = 'food_v2';
const LABOR_KEY = 'labor_v2';
const CACHE_TTL = 21600; // 6h — trigger refreshes every 10 min; long TTL = stale beats nothing
const LABOR_START = '2026-07-13';   // labor history begins here

/** One-time: install the refresh trigger (removes any old ones first). */
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshCache').timeBased().everyMinutes(10).create();
  refreshCache();
}

/** Called by the timer. Rebuilds both datasets; lock prevents overlap. */
function refreshCache() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // another rebuild is already running — skip
  try {
    writeCache(buildJson());
    writeCacheKey(LABOR_KEY, buildLaborJson());
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const p = e && e.parameter || {};
  if (p.diag === 'labor') return diagLabor_();
  if (p.diag) return diag();

  const isLabor = p.data === 'labor';
  const key     = isLabor ? LABOR_KEY : CACHE_KEY;
  const build   = isLabor ? buildLaborJson : buildJson;

  // 1. Serve from cache — the only path users should ever hit.
  if (!p.bust) {
    const cached = readCacheKey(key);
    if (cached) return json_(cached);
  }

  // 2. Cache empty. Only ONE request may rebuild; everyone else gets
  //    a tiny "warming" response immediately instead of piling up.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) return json_(JSON.stringify({warming: true}));
  try {
    const again = p.bust ? null : readCacheKey(key);
    if (again) return json_(again);
    const out = build();
    writeCacheKey(key, out);
    return json_(out);
  } finally {
    lock.releaseLock();
  }
}

function json_(s) {
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}

/** ?diag=labor — list every tab in the labor workbook with its headers. */
function diagLabor_() {
  const out = {sheetId: LABOR_SHEET_ID, tabs: []};
  let ss;
  try { ss = SpreadsheetApp.openById(LABOR_SHEET_ID); }
  catch (err) { return json_(JSON.stringify({error: 'Cannot open labor sheet: ' + err.message})); }
  ss.getSheets().forEach(function (s) {
    const lastRow = s.getLastRow(), lastCol = s.getLastColumn();
    let hdr = [];
    if (lastRow > 0 && lastCol > 0) {
      const scan = s.getRange(1, 1, Math.min(4, lastRow), Math.min(30, lastCol)).getValues();
      // pick the row with the most non-empty cells as the likely header
      let best = 0, bestN = -1;
      scan.forEach(function (r, i) {
        const n = r.filter(function (c) { return String(c).trim() !== ''; }).length;
        if (n > bestN) { bestN = n; best = i; }
      });
      hdr = scan[best].map(function (c) { return String(c).trim(); }).filter(String);
    }
    out.tabs.push({name: s.getName(), gid: s.getSheetId(), rows: lastRow, cols: lastCol, headers: hdr});
  });
  return json_(JSON.stringify(out));
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

function readCache()      { return readCacheKey(CACHE_KEY); }
function writeCache(json) { writeCacheKey(CACHE_KEY, json); }

/* CacheService caps a value at 100KB, so payloads are split across numbered
   keys. The count is stored alongside them: reading "until a gap" used to pick
   up leftover chunks from a previous, larger payload and append them to valid
   JSON, which produced a parse error at the join. Read exactly n, or nothing. */
function readCacheKey(key) {
  const cache = CacheService.getScriptCache();
  const n = parseInt(cache.get(key + '_n'), 10);
  if (!n || n < 1) return null;
  const names = [];
  for (let i = 0; i < n; i++) names.push(key + '_' + i);
  const got = cache.getAll(names);
  let out = '';
  for (let i = 0; i < n; i++) {
    const c = got[key + '_' + i];
    if (c === null || c === undefined) return null;  // partial: rebuild instead
    out += c;
  }
  return out;
}

function writeCacheKey(key, json) {
  const chunks = [];
  for (let i = 0; i < json.length; i += 90000) chunks.push(json.substr(i, 90000));
  const toStore = {};
  chunks.forEach(function (c, i) { toStore[key + '_' + i] = c; });
  const cache = CacheService.getScriptCache();
  try {
    // drop any chunks left by a previous, longer payload
    const stale = [];
    for (let i = chunks.length; i < chunks.length + 12; i++) stale.push(key + '_' + i);
    cache.removeAll(stale);
  } catch (err) {}
  // write the count last, so a half-written cache never looks complete
  try {
    cache.putAll(toStore, CACHE_TTL);
    cache.put(key + '_n', String(chunks.length), CACHE_TTL);
  } catch (err) {}
}

/* ══════════════════ LABOR ══════════════════ */

/** Locate a tab by name fragment, falling back to a column signature. */
function findSheet_(ss, nameFrag, mustHave) {
  const sheets = ss.getSheets();
  for (const s of sheets) {
    if (s.getName().toLowerCase().indexOf(nameFrag) !== -1) return s;
  }
  for (const s of sheets) {
    const hdr = s.getRange(1, 1, Math.min(3, s.getLastRow()), Math.min(30, s.getLastColumn()))
                 .getValues().map(r => r.map(c => String(c).trim().toLowerCase()));
    const flat = [].concat.apply([], hdr);
    if (mustHave.every(h => flat.indexOf(h.toLowerCase()) !== -1)) return s;
  }
  return null;
}

function headerRow_(values, marker) {
  for (let i = 0; i < Math.min(values.length, 6); i++) {
    const row = values[i].map(c => String(c).trim().toLowerCase());
    if (row.indexOf(marker.toLowerCase()) !== -1) return i;
  }
  return -1;
}

function readAll_(sh) {
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];
  return sh.getRange(1, 1, lastRow, lastCol).getValues();
}

function sid_(v) {
  const d = String(v == null ? '' : v).split('-')[0].replace(/\D/g, '');
  return d ? String(parseInt(d, 10)) : '';
}
function num_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  const n = Number(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function buildLaborJson() {
  const ss = SpreadsheetApp.openById(LABOR_SHEET_ID);

  const shPar = findSheet_(ss, 'par schedule', ['Sched Hrs', 'Forecast']);
  const shTc  = findSheet_(ss, 'actual timecard', ['Total Pay', 'Reg Hours']);
  const shEla = findSheet_(ss, 'employee labor', ['Schedule Hours', 'Actual Hours']);
  if (!shPar || !shTc || !shEla) {
    return JSON.stringify({error: 'Labor tab not found: ' +
      [shPar ? '' : 'PAR', shTc ? '' : 'Timecard', shEla ? '' : 'EmpLabor'].filter(String).join(', ') +
      '. Tabs present: ' + ss.getSheets().map(function(s){return s.getName();}).join(' | ')});
  }

  const names = {}, types = {}, days = {}, emp = {};
  const keyOf = (d, s) => d + '|' + s;

  // ---- PAR: forecast, scheduled hours, scheduled $ (one row per store/day)
  {
    const v = readAll_(shPar); const hi = headerRow_(v, 'Sched Hrs');
    if (hi < 0) return JSON.stringify({error: 'PAR header not found'});
    const H = v[hi].map(c => String(c).trim().toLowerCase());
    const col = n => H.indexOf(n.toLowerCase());
    const iD = col('Date'), iS = col('Stores'), iF = col('Forecast'),
          iSH = col('Sched Hrs'), iSD = col('Sched $');
    for (let i = hi + 1; i < v.length; i++) {
      const d = normalizeDate(v[i][iD]); const s = sid_(v[i][iS]);
      if (!d || !s || d < LABOR_START) continue;
      names[s] = names[s] || String(v[i][iS] || '').split('-').pop().trim();
      const k = keyOf(d, s);
      const o = days[k] || (days[k] = {dt: d, sid: s, fc: 0, sh: 0, sd: 0, sa: 0, ah: 0, ot: 0, pay: 0});
      o.fc += num_(v[i][iF]); o.sh += num_(v[i][iSH]); o.sd += num_(v[i][iSD]);
    }
  }

  // ---- Actual Timecard Details: actual hours, OT, pay (punch level)
  {
    const v = readAll_(shTc); const hi = headerRow_(v, 'Total Pay');
    if (hi < 0) return JSON.stringify({error: 'Timecard header not found'});
    const H = v[hi].map(c => String(c).trim().toLowerCase());
    const col = n => H.indexOf(n.toLowerCase());
    const iD = col('Date'), iS = col('Store Id'), iSN = col('Stores'), iE = col('Employee'),
          iJ = col('Job'), iOT = col('OT Hours'), iTH = col('Total Hours'),
          iTP = col('Total Pay'), iST = col('Store Type');
    for (let i = hi + 1; i < v.length; i++) {
      const d = normalizeDate(v[i][iD]); const s = sid_(v[i][iS]);
      if (!d || !s || d < LABOR_START) continue;
      names[s] = names[s] || String(v[i][iSN] || '').split('-').pop().trim();
      if (iST >= 0 && v[i][iST]) types[s] = String(v[i][iST]).trim();
      const k = keyOf(d, s);
      const o = days[k] || (days[k] = {dt: d, sid: s, fc: 0, sh: 0, sd: 0, sa: 0, ah: 0, ot: 0, pay: 0});
      o.ah += num_(v[i][iTH]); o.ot += num_(v[i][iOT]); o.pay += num_(v[i][iTP]);
      const nm = String(v[i][iE] || '').trim(); if (!nm) continue;
      const ek = d + '|' + s + '|' + nm;
      const e = emp[ek] || (emp[ek] = {dt: d, sid: s, n: nm, job: '', sh: 0, ah: 0, ot: 0, pay: 0});
      e.ah += num_(v[i][iTH]); e.ot += num_(v[i][iOT]); e.pay += num_(v[i][iTP]);
      if (v[i][iJ]) e.job = String(v[i][iJ]).trim();
    }
  }

  // ---- Employee Labor Analysis: per-employee scheduled hours
  {
    const v = readAll_(shEla); const hi = headerRow_(v, 'Schedule Hours');
    if (hi < 0) return JSON.stringify({error: 'Employee Labor header not found'});
    const H = v[hi].map(c => String(c).trim().toLowerCase());
    const col = n => H.indexOf(n.toLowerCase());
    const iD = col('Date'), iS = col('Store Id'), iE = col('Employee'),
          iJ = col('Job'), iSH = col('Schedule Hours');
    for (let i = hi + 1; i < v.length; i++) {
      const d = normalizeDate(v[i][iD]); const s = sid_(v[i][iS]);
      if (!d || !s || d < LABOR_START) continue;
      const nm = String(v[i][iE] || '').trim(); if (!nm) continue;
      const ek = d + '|' + s + '|' + nm;
      const e = emp[ek] || (emp[ek] = {dt: d, sid: s, n: nm, job: '', sh: 0, ah: 0, ot: 0, pay: 0});
      e.sh += num_(v[i][iSH]);
      if (!e.job && v[i][iJ]) e.job = String(v[i][iJ]).trim();
    }
  }

  const r2 = n => Math.round(n * 100) / 100;
  const dayArr = Object.keys(days).map(k => {
    const o = days[k];
    return {dt: o.dt, sid: o.sid, fc: r2(o.fc), sh: r2(o.sh), sd: r2(o.sd),
            sa: 0, ah: r2(o.ah), ot: r2(o.ot), pay: r2(o.pay)};
  }).sort((a, b) => a.dt < b.dt ? -1 : a.dt > b.dt ? 1 : (+a.sid) - (+b.sid));

  const empArr = Object.keys(emp).map(k => {
    const e = emp[k];
    return {dt: e.dt, sid: e.sid, n: e.n, job: e.job, sh: r2(e.sh), ah: r2(e.ah), ot: r2(e.ot), pay: r2(e.pay)};
  }).sort((a, b) => a.dt < b.dt ? -1 : a.dt > b.dt ? 1 : a.n < b.n ? -1 : 1);

  const stores = Object.keys(names).sort(function (a, b) { return (+a) - (+b); })
    .map(function (s) { return {id: s, name: names[s], type: types[s] || ''}; });

  return JSON.stringify({generated: new Date().toISOString(),
                         stores: stores, days: dayArr, emp: empArr});
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
