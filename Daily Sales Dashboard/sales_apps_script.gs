/**
 * Daily Sales Dashboard — sales feed.
 *
 * Serves the "Daily Sales Summary Consolidated" workbook as JSON.
 *   ?           → { ok, data:[ … ], meta:{ … } }
 *   ?diag=1     → lists every tab and its headers, without reading data
 *
 * History: this used to read ss.getSheets()[0]. When an "Index" tab was added
 * in front of the data, the script happily read the index, found no columns,
 * and returned an empty array with ok:true — a silent outage. It now finds the
 * sheet by name, falls back to a header signature, and refuses to return an
 * empty payload quietly.
 */

var SHEET_NAME = 'Sales';
var REQUIRED   = ['Date', 'Total Net Sales'];   // header signature for the fallback

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.diag) return json(diagnose());
    return json(buildPayload());
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
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
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  return { ok: true, workbook: ss.getName(), wouldRead: chosen, sheets: out };
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
