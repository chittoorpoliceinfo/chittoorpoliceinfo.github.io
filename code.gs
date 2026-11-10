/**
 * CHITTOOR POLICE INFORMATION – JSON API for GitHub Pages (password protected)
 * The password is stored ONLY here, in Script Properties – never in GitHub.
 *
 * Setup:  Project Settings → Script Properties → add  ACCESS_PASSWORD = <your password>
 * Deploy: Deploy → New deployment → Web app
 *         Execute as: Me   |   Who has access: Anyone
 */

var SHEET_ID   = '1XPKZBTKYi6PrQTrmXI89M-pbTFNfihVyjdNETlXQJPo';
var SHEET_NAME = 'CHITTOOR_POLICE';

var MIN_QUERY_LEN   = 3;    // must type at least 3 characters
var MAX_RESULTS     = 100;  // never send the whole sheet at once
var MASK_AADHAAR    = true; // show only last 4 digits of Aadhaar
var MAX_FAILS       = 20;   // wrong-password attempts allowed ...
var FAIL_WINDOW_SEC = 600;  // ... per 10 minutes, then login locks

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    var cache = CacheService.getScriptCache();
    var fails = Number(cache.get('fails') || 0);
    if (fails >= MAX_FAILS) {
      return json_({ ok: false, auth: false, error: 'Too many wrong attempts. Try again after 10 minutes.' });
    }

    var pwd = PropertiesService.getScriptProperties().getProperty('ACCESS_PASSWORD');
    if (!pwd || String(body.password || '') !== pwd) {
      cache.put('fails', String(fails + 1), FAIL_WINDOW_SEC);
      Utilities.sleep(1500);
      return json_({ ok: false, auth: false, error: 'Wrong password' });
    }

    if (body.action === 'login')  return json_({ ok: true });
    if (body.action === 'search') return json_(search_(String(body.q || '')));
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

// Plain GET just confirms the API is alive – it never returns data.
function doGet() {
  return json_({ ok: true, service: 'Chittoor Police Info API' });
}

function search_(q) {
  q = q.trim().toLowerCase();
  if (q.length < MIN_QUERY_LEN) {
    return { ok: false, error: 'Type at least ' + MIN_QUERY_LEN + ' characters.' };
  }

  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sh) return { ok: false, error: 'Sheet not found' };

  var values = sh.getDataRange().getDisplayValues(); // dates exactly as shown in sheet
  if (!values.length) return { ok: true, headers: [], rows: [], total: 0 };

  var headers = values[0];
  var aadhaarCols = [];
  headers.forEach(function (h, i) {
    if (/aadh?a?ar/i.test(String(h))) aadhaarCols.push(i);
  });

  var rows = [], total = 0;
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (row.join(' | ').toLowerCase().indexOf(q) === -1) continue;
    total++;
    if (rows.length >= MAX_RESULTS) continue;
    var out = row.slice();
    if (MASK_AADHAAR) aadhaarCols.forEach(function (c) { out[c] = mask_(out[c]); });
    rows.push(out);
  }
  return { ok: true, headers: headers, rows: rows, total: total };
}

function mask_(v) {
  var s = String(v || '').replace(/\s/g, '');
  return s.length > 4 ? 'XXXX XXXX ' + s.slice(-4) : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}