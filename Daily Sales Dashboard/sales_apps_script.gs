/**
 * Daily Sales Dashboard — sales feed.
 *
 * Serves the "Daily Sales Summary Consolidated" workbook as JSON.
 *   ?           → { ok, data:[ … ], meta:{ … } }
 *   ?diag=1     → lists every tab and its headers, without reading data
 *
 * History: this used to call getActiveSpreadsheet() and read getSheets()[0].
 * Both assumptions broke at once when the workbook was reorganised — the binding
 * went away and an "Index" tab appeared in front of the data — and the old code
 * reported it as an empty array with ok:true, a silent outage. It now opens the
 * book by ID, finds the tab by name with a header-signature fallback, and
 * refuses to return an empty payload quietly.
 */

/* "Daily Sales Summary Consolidated". Opened by ID rather than via
   getActiveSpreadsheet(): this project is standalone, so there is no active
   spreadsheet and the binding cannot be relied on. */
var SHEET_ID   = '1IXE6L8oELtaRLSHpM4RXRSplqRVIIaIqSbFaHnwexCw';
var SHEET_NAME = 'Sales';
var REQUIRED   = ['Date', 'Total Net Sales'];   // header signature for the fallback

function openBook() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  if (!ss) throw new Error('Could not open spreadsheet ' + SHEET_ID);
  return ss;
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.diag) return json(diagnose());
    if (p.data === 'scorecard') return json(buildScorecard());
    if (p.data === 'stores') return json(buildStores());
    return json(buildPayload());
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

/* ── Brand scorecard ───────────────────────────────────────────────────────
   One row per store per period, in two flavours: monthly and round-to-date.
   The tab also carries a "Selected Group Average" pseudo-store, kept and
   flagged rather than dropped — it is the brand's own benchmark, so it beats
   inventing thresholds of our own.
   Direction of travel, confirmed against Overall Rating across 500 rows:
     rating, cert, rev      higher is better
     acr, sos, mi           lower is better
     rv, msd, ltec  (Y/N)   Y is worse                                       */
var SCORE_NAME = 'Scorecard';
var SCORE_REQUIRED = ['Restaurant', 'Overall Rating'];

function buildScorecard() {
  var ss = openBook();
  var sheet = ss.getSheetByName(SCORE_NAME);
  if (!sheet || !hasCols(sheet, SCORE_REQUIRED)) {
    sheet = null;
    var all = ss.getSheets();
    for (var i = 0; i < all.length; i++) {
      if (hasCols(all[i], SCORE_REQUIRED)) { sheet = all[i]; break; }
    }
  }
  if (!sheet) throw new Error('No scorecard sheet found. Tabs: ' +
    ss.getSheets().map(function (s) { return s.getName(); }).join(', '));

  var v = sheet.getDataRange().getValues();
  var H = v[0].map(function (h) { return String(h).trim(); });
  var col = function (n) { return H.indexOf(n); };

  var iSK = col('Sort Key'), iPer = col('Period'), iPT = col('Period Type'),
      iRnd = col('Round'), iEnd = col('Period End'), iR = col('Restaurant'),
      iGM = col('GM Name'), iRat = col('Overall Rating'),
      iACR = col('Guest Satisfaction (ACR)'), iSOS = col('Window Time (SOS)'),
      iMI = col('Missing & Incorrect (M&I)'), iCert = col('Station Certification'),
      iREV = col('REV'), iRV = col('Roster Variance'),
      iMSD = col('Missing Speed Data'), iLT = col('LTEC Failure');
  if (iR < 0 || iRat < 0) throw new Error('Scorecard columns not found on "' + sheet.getName() + '"');

  var n  = function (x) { var t = String(x == null ? '' : x).replace(/[^0-9.\-]/g, ''); return t === '' ? null : Number(t); };
  var yn = function (x) { var t = String(x == null ? '' : x).trim().toUpperCase(); return t === 'Y' ? 1 : t === 'N' ? 0 : null; };

  var rows = [];
  for (var r = 1; r < v.length; r++) {
    var who = String(v[r][iR] == null ? '' : v[r][iR]).trim();
    if (!who) continue;
    var isAvg = /average/i.test(who);
    var sid = isAvg ? 'avg' : who.replace(/\D/g, '');
    if (!sid) continue;
    var pt = String(v[r][iPT] == null ? '' : v[r][iPT]).trim();
    rows.push({
      sk:   n(v[r][iSK]),
      per:  String(v[r][iPer] == null ? '' : v[r][iPer]).trim(),
      pt:   /round/i.test(pt) ? 'rtd' : 'mo',
      rnd:  iRnd  >= 0 ? String(v[r][iRnd] || '').trim() : '',
      end:  iEnd  >= 0 ? String(v[r][iEnd] || '').trim() : '',
      sid:  sid,
      avg:  isAvg ? 1 : 0,
      gm:   iGM   >= 0 ? String(v[r][iGM] || '').trim() : '',
      rating: n(v[r][iRat]),
      acr:  iACR  >= 0 ? n(v[r][iACR])  : null,
      sos:  iSOS  >= 0 ? n(v[r][iSOS])  : null,
      mi:   iMI   >= 0 ? n(v[r][iMI])   : null,
      cert: iCert >= 0 ? n(v[r][iCert]) : null,
      rev:  iREV  >= 0 ? n(v[r][iREV])  : null,
      rv:   iRV   >= 0 ? yn(v[r][iRV])  : null,
      msd:  iMSD  >= 0 ? yn(v[r][iMSD]) : null,
      ltec: iLT   >= 0 ? yn(v[r][iLT])  : null
    });
  }
  if (!rows.length) throw new Error('Scorecard tab "' + sheet.getName() + '" produced 0 rows');

  var sks = rows.map(function (o) { return o.sk; }).filter(function (x) { return x; }).sort();
  return {
    ok: true, rows: rows,
    meta: {
      tab: sheet.getName(), rows: rows.length,
      first: sks[0], last: sks[sks.length - 1],
      generated: new Date().toISOString()
    }
  };
}


/* ── Store roster ──────────────────────────────────────────────────────────
   Area Director assignments, store names and closures, read from the
   "Store Details" tab and joined to the store-type lookup.

   That tab also carries Federal Tax IDs, store/AD/GM email addresses and
   phone numbers. The dashboard is a public page, so none of that leaves
   here: only id, name, Area Director, type and a closed flag.             */
var ROSTER_NAME = 'Store Details';
var TYPES_NAME  = 'Sheet9';

function buildStores() {
  var ss = openBook();

  // store id -> type, from the little lookup tab
  var types = {};
  var tSheet = ss.getSheetByName(TYPES_NAME);
  if (tSheet && tSheet.getLastRow() > 1) {
    var tv = tSheet.getDataRange().getValues();
    var tH = tv[0].map(function (h) { return String(h).trim(); });
    var iID = tH.indexOf('StoreID'), iTY = tH.indexOf('StoreType');
    if (iID >= 0 && iTY >= 0) {
      for (var t = 1; t < tv.length; t++) {
        var tid = String(tv[t][iID] == null ? '' : tv[t][iID]).replace(/\D/g, '');
        if (tid) types[String(parseInt(tid, 10))] = String(tv[t][iTY] || '').trim();
      }
    }
  }

  var sheet = ss.getSheetByName(ROSTER_NAME);
  if (!sheet || !hasCols(sheet, ['Store', 'Area Director'])) {
    var all = ss.getSheets();
    sheet = null;
    for (var i = 0; i < all.length; i++) {
      if (hasCols(all[i], ['Store', 'Area Director'])) { sheet = all[i]; break; }
    }
  }
  if (!sheet) throw new Error('No store roster found. Tabs: ' +
    ss.getSheets().map(function (x) { return x.getName(); }).join(', '));

  var v = sheet.getDataRange().getValues();
  var H = v[0].map(function (h) { return String(h).trim(); });
  var col = function (n) { return H.indexOf(n); };
  var iStore = col('Store'), iAD = col('Area Director'), iClosed = col('Permanently Closed');
  if (iStore < 0) throw new Error('No "Store" column on "' + sheet.getName() + '"');

  var rows = [];
  for (var r = 1; r < v.length; r++) {
    var raw = String(v[r][iStore] == null ? '' : v[r][iStore]).trim();
    if (!raw) continue;
    // "Bardstown - 5671" -> name + id. Rows without an id are Corporate, not stores.
    var m = raw.match(/^(.*?)\s*-\s*(\d+)\s*$/);
    if (!m) continue;
    var name = m[1].trim(), sid = String(parseInt(m[2], 10));
    if (!name || !sid) continue;
    var ad = iAD >= 0 ? String(v[r][iAD] == null ? '' : v[r][iAD]).trim() : '';
    var closedRaw = iClosed >= 0 ? v[r][iClosed] : '';
    rows.push({
      sid:    sid,
      name:   name,
      ad:     ad || null,
      type:   types[sid] || null,
      closed: closedRaw ? 1 : 0
    });
  }
  if (!rows.length) throw new Error('Roster tab "' + sheet.getName() + '" produced 0 stores');

  var ads = {};
  rows.forEach(function (o) { if (o.ad && !o.closed) ads[o.ad] = (ads[o.ad] || 0) + 1; });

  return {
    ok: true, stores: rows,
    meta: {
      tab: sheet.getName(), stores: rows.length,
      open: rows.filter(function (o) { return !o.closed; }).length,
      directors: ads,
      generated: new Date().toISOString()
    }
  };
}

function hasCols(sheet, need) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return false;
  var hdrs = headersOf(sheet);
  for (var i = 0; i < need.length; i++) if (hdrs.indexOf(need[i]) === -1) return false;
  return true;
}

/** Locate the data sheet by name, then by header signature. Never by position. */
function findSheet(ss) {
  var byName = ss.getSheetByName(SHEET_NAME);
  if (byName && hasSignature(byName)) return byName;

  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (hasSignature(sheets[i])) return sheets[i];
  }
  throw new Error('No sheet found with columns: ' + REQUIRED.join(', ') +
                  '. Tabs present: ' + sheets.map(function (s) { return s.getName(); }).join(', '));
}

function hasSignature(sheet) {
  if (sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return false;
  var hdrs = headersOf(sheet);
  for (var i = 0; i < REQUIRED.length; i++) {
    if (hdrs.indexOf(REQUIRED[i]) === -1) return false;
  }
  return true;
}

function headersOf(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn())
              .getValues()[0]
              .map(function (h) { return String(h).trim(); });
}

function buildPayload() {
  var ss    = openBook();
  var sheet = findSheet(ss);
  var vals  = sheet.getDataRange().getValues();
  var hdrs  = vals[0].map(function (h) { return String(h).trim(); });

  var ci = function (name) {
    var i = hdrs.indexOf(name);
    if (i === -1) throw new Error('Column "' + name + '" not found on tab "' + sheet.getName() + '"');
    return i;
  };

  var iDate  = ci('Date');
  var iSID   = ci('Stores');
  var iSName = ci('Store Name');
  var iSType = ci('Store Type');
  var iSales = ci('Total Net Sales');
  var iCount = ci('Customer Count');
  var iAC    = ci('Average Check');
  var iPer   = ci('Period');
  var iPY    = ci('Period Year');
  var iPN    = ci('Period Number');
  var iPW    = ci('Period Week');
  var iWN    = ci('Week Number');

  var tz     = Session.getScriptTimeZone();
  var rows   = [];
  var skipped = 0;

  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    if (!r[iDate]) continue;

    var py = Number(r[iPY]);
    var pn = Number(r[iPN]);
    if (py < 2026) { skipped++; continue; }      // P1 2026 onwards only

    var dv = r[iDate];
    var dt = dv instanceof Date
      ? Utilities.formatDate(dv, tz, 'yyyy-MM-dd')
      : String(dv);

    rows.push({
      dt:    dt,
      sid:   r[iSID],
      sn:    r[iSName],
      st:    r[iSType],
      sales: Number(r[iSales]) || 0,
      cnt:   Number(r[iCount]) || 0,
      ac:    Number(r[iAC])    || 0,
      per:   r[iPer],
      py:    py,
      pn:    pn,
      pw:    Number(r[iPW]) || 0,
      wn:    Number(r[iWN]) || 0
    });
  }

  // An empty payload is never normal here. Fail loudly so the dashboard shows an
  // error instead of an empty dashboard that looks like "no sales today".
  if (!rows.length) {
    throw new Error('Read tab "' + sheet.getName() + '" (' + (vals.length - 1) +
                    ' data rows) but produced 0 usable rows; ' + skipped +
                    ' were pre-2026. Check the Date / Period Year columns.');
  }

  var dates = rows.map(function (r) { return r.dt; }).sort();

  return {
    ok: true,
    data: rows,
    meta: {
      tab: sheet.getName(),
      rows: rows.length,
      skippedPre2026: skipped,
      first: dates[0],
      last: dates[dates.length - 1],
      generated: new Date().toISOString()
    }
  };
}

/** Tab lister — tells us what the script can actually see. */
function diagnose() {
  var ss = openBook();
  var out = ss.getSheets().map(function (s, i) {
    var hdrs = s.getLastColumn() ? headersOf(s) : [];
    return {
      position: i,
      name: s.getName(),
      rows: s.getLastRow(),
      cols: s.getLastColumn(),
      matchesSignature: hasSignature(s),
      headers: hdrs.slice(0, 20)
    };
  });
  var chosen = null;
  try { chosen = findSheet(ss).getName(); } catch (err) { chosen = 'ERROR: ' + err.message; }
  return { ok: true, workbook: ss.getName(), sheetId: SHEET_ID, wouldRead: chosen, sheets: out };
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
